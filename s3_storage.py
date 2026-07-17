"""
Общий модуль работы с хранилищем сканов (Cloud.ru S3). Вынесен из webforms.py, чтобы им мог
пользоваться и bot.py (скачивание документов работника в чат), без дублирования кода.

Ключи — только в переменных окружения Railway. Если boto3 не установлен или ключи не заданы,
функции бросают понятную RuntimeError, старт приложения не роняется.
"""
import os

S3_ENDPOINT = os.environ.get("S3_ENDPOINT", "")        # напр. https://s3.cloud.ru (без имени бакета!)
S3_BUCKET = os.environ.get("S3_BUCKET", "")            # напр. migbot51-scan
S3_REGION = os.environ.get("S3_REGION", "ru-central-1")
S3_ACCESS_KEY = os.environ.get("S3_ACCESS_KEY", "")
S3_SECRET_KEY = os.environ.get("S3_SECRET_KEY", "")

# Персональные сканы (в карточке работника, свои у каждого).
SCAN_TYPES = {
    "passport": "Паспорт работника (все страницы)",
    "migration_card": "Миграционная карта",
    "arrival_notice": "Уведомление о прибытии (отрывная часть с Госуслуг)",
    "payment_registration": "Исполненная платёжка — постановка на учёт (500 ₽)",
    "payment_renewal": "Исполненная платёжка — продление регистрации (1000 ₽)",
    "medical_certificate": "Справка о медосвидетельствовании (медкомиссия)",
}
# Платёжки, для которых при загрузке проверяем фамилию работника в назначении платежа (PDF).
PAYMENT_SCAN_TYPES = {"payment_registration", "payment_renewal"}
# Общие документы (один файл на всех) — паспорт директора, основание на адрес.
COMMON_DOC_TYPES = {
    "director_passport": "Паспорт директора (принимающая сторона)",
    "address_basis": "Документ-основание на адрес подразделения",
}
# Пусто: персональные типы общими не бывают (для _s3_list_for_employee).
SCAN_COMMON_TYPES = set()

# Сканы, привязанные не к работнику, а к производственной сущности (подписанные документы).
# Хранятся по конвенции "{kind}_{entity_id}/scan" — отдельной колонки в БД не требуют.
ENTITY_SCAN_KINDS = {
    "workorder": "Наряд-допуск",
    "instruction": "Журнал инструктажа",
    "order": "Приказ ОТ",
}


_S3_CLIENT_CACHE = None  # переиспользуемый boto3-клиент: создание клиента дорогое, не плодим


def _s3_client():
    """Возвращает boto3 S3-клиент для Cloud.ru (кэшируется на процесс) или бросает RuntimeError."""
    global _S3_CLIENT_CACHE
    if _S3_CLIENT_CACHE is not None:
        return _S3_CLIENT_CACHE
    if not (S3_ENDPOINT and S3_BUCKET and S3_ACCESS_KEY and S3_SECRET_KEY):
        raise RuntimeError(
            "Хранилище сканов не настроено: заданы не все переменные S3_ENDPOINT/S3_BUCKET/"
            "S3_ACCESS_KEY/S3_SECRET_KEY. Проверьте переменные окружения на Railway."
        )
    try:
        import boto3
    except ImportError:
        raise RuntimeError("Библиотека boto3 не установлена (добавьте boto3 в requirements).")
    _S3_CLIENT_CACHE = boto3.client(
        "s3",
        endpoint_url=S3_ENDPOINT,
        region_name=S3_REGION,
        aws_access_key_id=S3_ACCESS_KEY,
        aws_secret_access_key=S3_SECRET_KEY,
    )
    return _S3_CLIENT_CACHE


def _scan_key(scan_type: str, employee_id: str | None) -> str:
    """Ключ объекта в бакете. Общие типы — под common/, персональные — под employee_<id>/."""
    if scan_type in SCAN_COMMON_TYPES:
        return f"common/{scan_type}"
    return f"employee_{employee_id}/{scan_type}"


def _s3_upload(scan_type: str, employee_id: str | None, data: bytes, content_type: str,
               metadata: dict | None = None) -> None:
    """Загружает скан в бакет. metadata — пользовательские метаданные объекта."""
    client = _s3_client()
    key = _scan_key(scan_type, employee_id)
    try:
        kwargs = dict(Bucket=S3_BUCKET, Key=key, Body=data, ContentType=content_type)
        if metadata:
            kwargs["Metadata"] = metadata
        client.put_object(**kwargs)
    except Exception as e:
        raise RuntimeError(f"Не удалось загрузить скан в хранилище: {str(e)[:200]}")


def _s3_list_for_employee(employee_id: str) -> dict:
    """{scan_type: {"present": True, "check": bool}} для сканов работника."""
    present = {}
    try:
        client = _s3_client()
    except RuntimeError:
        return present
    for scan_type in SCAN_TYPES:
        key = _scan_key(scan_type, employee_id)
        try:
            head = client.head_object(Bucket=S3_BUCKET, Key=key)
            meta = head.get("Metadata", {}) or {}
            present[scan_type] = {"present": True, "check": meta.get("check") == "1"}
        except Exception:
            pass
    return present


def _s3_download(scan_type: str, employee_id: str | None):
    """Возвращает (bytes, content_type) скана или бросает RuntimeError, если нет/недоступен."""
    client = _s3_client()
    key = _scan_key(scan_type, employee_id)
    try:
        obj = client.get_object(Bucket=S3_BUCKET, Key=key)
        return obj["Body"].read(), obj.get("ContentType", "application/octet-stream")
    except Exception as e:
        raise RuntimeError(f"Скан не найден или недоступен: {str(e)[:200]}")


def _s3_delete(scan_type: str, employee_id: str | None) -> None:
    client = _s3_client()
    key = _scan_key(scan_type, employee_id)
    try:
        client.delete_object(Bucket=S3_BUCKET, Key=key)
    except Exception as e:
        raise RuntimeError(f"Не удалось удалить скан: {str(e)[:200]}")


# --- Сканы производственных сущностей (наряд / инструктаж / приказ) --------------------
# Конвенция ключа: "{kind}_{entity_id}/scan". Отдельной колонки в БД не требуют — наличие
# проверяется head_object'ом, как у сканов работника. entity_id для инструктажа — тип
# журнала (introductory / primary_workplace), для наряда/приказа — их id.

def _entity_scan_key(kind: str, entity_id: str) -> str:
    return f"{kind}_{entity_id}/scan"


def _s3_upload_entity(kind: str, entity_id: str, data: bytes, content_type: str,
                      metadata: dict | None = None) -> None:
    """Загружает скан подписанного документа (наряд/инструктаж/приказ) в бакет."""
    if kind not in ENTITY_SCAN_KINDS:
        raise RuntimeError(f"Неизвестный тип сущности для скана: {kind}")
    client = _s3_client()
    key = _entity_scan_key(kind, entity_id)
    try:
        kwargs = dict(Bucket=S3_BUCKET, Key=key, Body=data, ContentType=content_type)
        if metadata:
            kwargs["Metadata"] = metadata
        client.put_object(**kwargs)
    except Exception as e:
        raise RuntimeError(f"Не удалось загрузить скан в хранилище: {str(e)[:200]}")


def _s3_entity_present(kind: str, entity_id: str) -> bool:
    """True, если скан сущности уже загружен."""
    try:
        client = _s3_client()
    except RuntimeError:
        return False
    try:
        client.head_object(Bucket=S3_BUCKET, Key=_entity_scan_key(kind, entity_id))
        return True
    except Exception:
        return False


def _s3_download_entity(kind: str, entity_id: str):
    """(bytes, content_type) скана сущности или RuntimeError."""
    client = _s3_client()
    try:
        obj = client.get_object(Bucket=S3_BUCKET, Key=_entity_scan_key(kind, entity_id))
        return obj["Body"].read(), obj.get("ContentType", "application/octet-stream")
    except Exception as e:
        raise RuntimeError(f"Скан не найден или недоступен: {str(e)[:200]}")


def _common_key(doc_type: str) -> str:
    return f"common/{doc_type}"


def _s3_upload_common(doc_type: str, data: bytes, content_type: str) -> None:
    client = _s3_client()
    try:
        client.put_object(Bucket=S3_BUCKET, Key=_common_key(doc_type), Body=data, ContentType=content_type)
    except Exception as e:
        raise RuntimeError(f"Не удалось загрузить документ: {str(e)[:200]}")


def _s3_list_common() -> dict:
    present = {}
    try:
        client = _s3_client()
    except RuntimeError:
        return present
    for dt in COMMON_DOC_TYPES:
        try:
            client.head_object(Bucket=S3_BUCKET, Key=_common_key(dt))
            present[dt] = True
        except Exception:
            pass
    return present


def _s3_download_common(doc_type: str):
    client = _s3_client()
    try:
        obj = client.get_object(Bucket=S3_BUCKET, Key=_common_key(doc_type))
        return obj["Body"].read(), obj.get("ContentType", "application/octet-stream")
    except Exception as e:
        raise RuntimeError(f"Документ не найден или недоступен: {str(e)[:200]}")


def _s3_delete_common(doc_type: str) -> None:
    client = _s3_client()
    try:
        client.delete_object(Bucket=S3_BUCKET, Key=_common_key(doc_type))
    except Exception as e:
        raise RuntimeError(f"Не удалось удалить документ: {str(e)[:200]}")


def _s3_clear_check(scan_type: str, employee_id: str | None) -> None:
    """Снимает метку «требует проверки» со скана: копирует объект сам в себя с пустыми
    метаданными (S3 не умеет менять метаданные на месте, только через copy с REPLACE)."""
    client = _s3_client()
    key = _scan_key(scan_type, employee_id)
    try:
        client.copy_object(
            Bucket=S3_BUCKET, Key=key,
            CopySource={"Bucket": S3_BUCKET, "Key": key},
            Metadata={}, MetadataDirective="REPLACE",
        )
    except Exception as e:
        raise RuntimeError(f"Не удалось снять метку: {str(e)[:200]}")


def _ext_for(content_type: str) -> str:
    """Расширение файла по content-type (для имён скачиваемых файлов)."""
    ct = (content_type or "").lower()
    if "pdf" in ct:
        return "pdf"
    if "jpeg" in ct or "jpg" in ct:
        return "jpg"
    if "png" in ct:
        return "png"
    return "bin"
