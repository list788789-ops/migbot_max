"""
Временный тестовый модуль для проверки OCR (passporteye) на реальных фото удостоверений.
ИЗОЛИРОВАН от основной системы — ничего не трогает, только читает загруженное фото и показывает,
что распознала passporteye из MRZ-зоны + прошли ли контрольные суммы.

ПОДКЛЮЧЕНИЕ (одна строка в webforms.py, в самом конце после создания app):
    import ocr_test; ocr_test.register(app)

REQUIREMENTS (добавить на время теста):
    passporteye>=2.2
    (passporteye тянет за собой: pytesseract, opencv-python-headless, scikit-image — тяжёлые,
     ~200 МБ; сборка на Railway станет дольше. После теста убрать из requirements.)

СИСТЕМНЫЙ ПАКЕТ (Railway env var RAILPACK_DEPLOY_APT_PACKAGES):
    добавить tesseract-ocr и tesseract-ocr-rus к libreoffice, через пробел:
    libreoffice tesseract-ocr tesseract-ocr-rus

УДАЛЕНИЕ ПОСЛЕ ТЕСТА: убрать строку import ocr_test; ocr_test.register(app), удалить файл,
убрать passporteye из requirements и tesseract из apt-переменной.
"""
from fastapi import Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse
import html as _html


_PAGE = """<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OCR-тест</title>
<style>
body{{font-family:-apple-system,Arial,sans-serif;max-width:760px;margin:0 auto;padding:16px;color:#141a1f}}
h1{{font-size:22px}}
input[type=file]{{display:block;width:100%;margin:12px 0;padding:12px;border:1px solid #b8c0cc;border-radius:10px;font-size:16px;background:#fff}}
button{{background:#4a90e2;color:#fff;border:none;border-radius:10px;padding:14px 20px;font-size:16px;width:100%;cursor:pointer}}
.res{{margin-top:16px;padding:14px;border:1px solid #e6e9ee;border-radius:10px;background:#f8fafc}}
.res b{{color:#1a7f37}}
.err{{color:#b00}}
table{{border-collapse:collapse;width:100%;margin-top:8px}}
td{{border:1px solid #e0e4ea;padding:6px 8px;font-size:14px;vertical-align:top}}
td:first-child{{color:#667;width:42%}}
.muted{{color:#889;font-size:13px}}
pre{{white-space:pre-wrap;word-break:break-word;background:#fff;padding:8px;border-radius:6px;font-size:12px}}
</style></head><body>
<h1>OCR-тест удостоверения (passporteye)</h1>
<p class="muted">Загрузите фото удостоверения (лучше стороной с MRZ — тремя строчками латиницей
внизу). Скрипт прогонит passporteye и покажет, что распозналось и сошлись ли контрольные суммы.
Это временный тест, на рабочие данные не влияет.</p>
<form method="post" action="/ocr-test" enctype="multipart/form-data">
<input type="file" name="photo" accept="image/*" required>
<button type="submit">Распознать</button>
</form>
{result}
</body></html>"""


def _translit_to_cyrillic(s: str) -> str:
    """Обратная транслитерация латиница->кириллица для ЧЕРНОВИКА ФИО. ВНИМАНИЕ: неоднозначна —
    BALGABAYEV->Балгабайев (офиц. Балгабаев), ZHUNISSOV->Жуниссов (офиц. Жунисов). Даёт
    правдоподобное, но НЕ гарантированно точное написание. Только подсказка, кадровик ОБЯЗАН
    сверить с кириллицей на лицевой стороне удостоверения."""
    if not s:
        return ""
    s = s.upper()
    for lat, cyr in [("SHCH","Щ"),("KH","Х"),("ZH","Ж"),("CH","Ч"),("SH","Ш"),
                     ("YU","Ю"),("YA","Я"),("YO","Ё"),("TS","Ц")]:
        s = s.replace(lat, cyr)
    single = {"A":"А","B":"Б","V":"В","G":"Г","D":"Д","E":"Е","Z":"З","I":"И","Y":"Й",
              "K":"К","L":"Л","M":"М","N":"Н","O":"О","P":"П","R":"Р","S":"С","T":"Т",
              "U":"У","F":"Ф","H":"Х","C":"К","J":"Ж","Q":"К","W":"В","X":"КС"}
    out = "".join(single.get(ch, ch) for ch in s)
    return out.capitalize()


def _run_ocr(image_bytes: bytes) -> str:
    """Прогоняет passporteye по фото, возвращает HTML с результатом."""
    try:
        import io
        from passporteye import read_mrz
        from PIL import Image, ImageOps
    except ImportError:
        return ('<div class="res err">passporteye не установлен. Добавьте в requirements '
                '<code>passporteye&gt;=2.2</code> и tesseract-ocr в apt-пакеты Railway, '
                'передеплойте.</div>')

    # НОРМАЛИЗАЦИЯ: телефонные фото удостоверения часто повёрнуты (MRZ идёт вертикально).
    # Пробуем автоповорот по EXIF + перебор 4 поворотов (0/90/180/270). Берём тот результат,
    # где у passporteye валидность MRZ выше (контрольные суммы) — не угадываем, а проверяем.
    try:
        base = Image.open(io.BytesIO(image_bytes))
        base = ImageOps.exif_transpose(base)
        if base.mode != "RGB":
            base = base.convert("RGB")
    except Exception as e:
        return f'<div class="res err">Не удалось открыть изображение: {_html.escape(str(e)[:200])}</div>'

    best = None
    best_score = -1
    best_angle = 0
    for angle in (0, 90, 180, 270):
        try:
            rot = base.rotate(angle, expand=True)
            buf = io.BytesIO()
            rot.save(buf, format="PNG")
            buf.seek(0)
            m = read_mrz(buf)
            if m is not None:
                score = m.to_dict().get("valid_score", 0)
                if score > best_score:
                    best_score, best, best_angle = score, m, angle
        except Exception:
            continue

    mrz = best
    if mrz is None:
        return ('<div class="res err">MRZ-зона не найдена ни при одном повороте. Причины: '
                'блик, сильный наклон, обрезан край, низкое качество, не та сторона. '
                'Попробуйте более ровное фото стороны с MRZ при хорошем свете.</div>')

    d = mrz.to_dict()

    # ИИН passporteye не выделяет отдельным полем, но в казахском удостоверении (формат TD1)
    # он лежит в optional-данных ПЕРВОЙ строки MRZ: 12 цифр после номера и его контрольной цифры.
    # Достаём из сырого текста вручную. Проверка: 12 цифр, первые 6 = дата рождения (ГГММДД).
    iin = None
    citizenship_raw = None
    citizenship_corrected = False
    try:
        raw = (d.get("raw_text", "") or "")
        lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
        if lines:
            l1 = lines[0].replace(" ", "")
            # optional first line: символы с 15-й позиции (после IDKAZ+9номер+1контр)
            opt1 = l1[15:].replace("<", "")
            if len(opt1) >= 12 and opt1[:12].isdigit():
                iin = opt1[:12]
        if len(lines) >= 2:
            l2 = lines[1].replace(" ", "")
            # гражданство: позиции 15-18 второй строки. OCR часто искажает Z<->L, Z<->2 —
            # KAZ читается как KAL/KA2/KAI. Раз все работники граждане Казахстана, распознанное
            # «похоже на KAZ» нормализуем в KAZ (с пометкой), точное KAZ оставляем как есть.
            cand = l2[15:18]
            if cand == "KAZ":
                citizenship_raw = "KAZ"
            elif cand[:2] == "KA":  # KAL, KA2, KAI и т.п. — искажённое KAZ
                citizenship_raw = "KAZ"
                citizenship_corrected = True
            elif cand.isalpha():
                citizenship_raw = cand
    except Exception:
        pass

    # ключевые поля + валидность (контрольные суммы)
    fields = [
        ("Тип документа", d.get("type")),
        ("Страна", d.get("country")),
        ("Номер документа", d.get("number")),
        ("Проверка номера (контр. сумма)", "✓ ок" if d.get("valid_number") else "✗ НЕ сошлась"),
        ("Фамилия (латиницей)", d.get("surname")),
        ("Фамилия (транслит, ЧЕРНОВИК)", _translit_to_cyrillic(d.get("surname") or "")),
        ("Имя (латиницей)", d.get("names")),
        ("Имя (транслит, ЧЕРНОВИК)", _translit_to_cyrillic(d.get("names") or "")),
        ("Дата рождения (ГГММДД)", d.get("date_of_birth")),
        ("Проверка даты рожд.", "✓ ок" if d.get("valid_date_of_birth") else "✗ НЕ сошлась"),
        ("Срок действия (ГГММДД)", d.get("expiration_date")),
        ("Проверка срока", "✓ ок" if d.get("valid_expiration_date") else "✗ НЕ сошлась"),
        ("Пол", d.get("sex")),
        ("Гражданство (из MRZ)", (citizenship_raw or d.get("nationality") or "—")
            + (" (скорректировано из искажённого OCR)" if citizenship_corrected else "")),
        ("ИИН (из optional MRZ)", iin or d.get("personal_number") or "не извлечён"),
        ("Общая валидность MRZ", "✓ всё сошлось" if d.get("valid_score", 0) == 100 else f"частично ({d.get('valid_score')}%)"),
    ]
    rows = "".join(
        f"<tr><td>{_html.escape(str(k))}</td><td>{_html.escape(str(v if v is not None else '—'))}</td></tr>"
        for k, v in fields
    )
    raw = _html.escape(str(d.get("raw_text", "") or ""))
    return (f'<div class="res"><b>MRZ распознана (поворот {best_angle}°).</b> Проверьте поля и контрольные суммы — '
            f'«✗ НЕ сошлась» означает ошибку распознавания в этом поле:'
            f'<table>{rows}</table>'
            f'<p class="muted">Сырой текст MRZ (как распозналось):</p><pre>{raw}</pre>'
            f'<p class="muted">Вывод: если контрольные суммы номера/даты «✓ ок» — этим полям '
            f'можно доверять как черновику. «✗» — распозналось с ошибкой, вводить вручную. '
            f'Кириллическое ФИО (как в карточке) passporteye НЕ даёт — только латиницу из MRZ.</p></div>')


def register(app):
    """Подключает тестовые роуты к переданному FastAPI-приложению."""

    @app.get("/ocr-test", response_class=HTMLResponse)
    def ocr_test_page(request: Request):
        if not request.session.get("user_id") and not request.session.get("role"):
            return RedirectResponse("/login", status_code=303)
        return _PAGE.format(result="")

    @app.post("/ocr-test", response_class=HTMLResponse)
    async def ocr_test_run(request: Request, photo: UploadFile = File(...)):
        if not request.session.get("user_id") and not request.session.get("role"):
            return RedirectResponse("/login", status_code=303)
        data = await photo.read()
        if not data:
            return _PAGE.format(result='<div class="res err">Пустой файл.</div>')
        result = _run_ocr(data)
        return _PAGE.format(result=result)
