"""
Модели БД для бота миграционного учёта.

Ключевые решения, зафиксированные в диалоге с заказчиком (не менять без причины):
- category у employee — enum с полным набором значений с самого начала (eaeu/patent/visa/hqs),
  даже если сейчас используется только 'eaeu'. Это дёшево сейчас и дорого добавлять потом.
- obligations.deadline_unit различает calendar/working дни — ЕАЭС считается в календарных днях (30),
  уведомление о договоре — в рабочих (3). Смешение этих единиц — источник ошибок в дедлайнах.
- consent_status блокирует создание obligations, пока не confirmed (см. obligations.py).
- invoices разделяет clinic_id (куда идёт сотрудник) и payer_id (кто платит) — это разные сущности,
  заказчик услуги (ИП) не обязан совпадать с клиникой.

2026-07: добавлены поля под перенос данных из ручной xlsx-таблицы (паспорт, адрес, дата рождения,
медкомиссия, срок регистрации, страна въезда). entry_date, category, citizenship переведены в
nullable=True — в исходной таблице 9 из 70 строк не имеют даты въезда, часть строк без гражданства/
категории. Бизнес-логика, которая от них зависит (расчёт obligations), должна сама проверять None
и не создавать дедлайны для неполных карточек, а не полагаться на NOT NULL в БД.

2026-07 (второе изменение): добавлены address_since (Employee) и is_current (Obligation) —
поддержка версионирования обязательств при смене места пребывания. Эти колонки добавлены
в уже существующие таблицы вручную через ALTER TABLE (см. obligations.py) — Base.metadata.create_all
не добавляет колонки в существующие таблицы, только создаёт отсутствующие таблицы целиком.
Если разворачиваешь БД с нуля — create_all создаст обе колонки автоматически, ALTER TABLE не нужен.
"""

from __future__ import annotations

import enum
import uuid
from datetime import date, datetime

from sqlalchemy import (
    Integer,
    Boolean,
    Date,
    DateTime,
    Enum,
    ForeignKey,
    Numeric,
    String,
    Text,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


def _uuid() -> str:
    return str(uuid.uuid4())


class Category(str, enum.Enum):
    EAEU = "eaeu"          # Казахстан, Армения, Киргизия — 30 календарных дней на учёт, без патента
    BELARUS = "belarus"    # отдельная категория ЕАЭС — 90 календарных дней (не путать с общим EAEU)
    PATENT = "patent"      # безвизовые не-ЕАЭС — патент, 7 (Узбекистан/Таджикистан — 15) дней
    VISA = "visa"          # визовые страны — разрешение на работу, 7 рабочих дней
    HQS = "hqs"            # высококвалифицированные специалисты


class RegistrationStatus(str, enum.Enum):
    """Статус миграционного учёта. Управляет тем, от какой даты считать дедлайны.
    PRIMARY — первичный учёт: регистрация/медосмотр/дактилоскопия от даты въезда (entry_date).
    PRIOR — ранее стоял на учёте в РФ (приехал на вахту из другого региона): регистрация от
    даты прибытия (address_since), медосмотр и дактилоскопия ЗАНОВО НЕ создаются.
    NULL (не задан) — создание обязательств ЗАБЛОКИРОВАНО до заполнения кадровиком; в карточке
    и списке показывается громкая пометка. Пустой статус НЕ трактуется как PRIMARY намеренно —
    кадровик обязан выбрать явно (решение зафиксировано в задачах)."""
    PRIMARY = "primary"
    PRIOR = "prior"


class UserRole(str, enum.Enum):
    """Роли пользователей системы. PRORAB — только чтение и скачивание документов (договор,
    квитанции, отчёты), без записи в БД. KADROVIK — всё по работникам (заключение, статусы,
    правки), но НЕ управление пользователями. ADMIN — всё + одобрение заявок и управление
    пользователями. Роль назначает админ при одобрении заявки (заявитель её не выбирает)."""
    PRORAB = "prorab"
    KADROVIK = "kadrovik"
    ADMIN = "admin"


class UserStatus(str, enum.Enum):
    """PENDING — заявка подана, вход заблокирован до одобрения админом. APPROVED — активен,
    может входить. BLOCKED — доступ отозван админом (вход запрещён, но запись сохранена)."""
    PENDING = "pending"
    APPROVED = "approved"
    BLOCKED = "blocked"


class ConsentStatus(str, enum.Enum):
    DRAFT = "draft"
    CONFIRMED = "confirmed"


class ConsentMethod(str, enum.Enum):
    PAPER_SCAN = "paper_scan"
    BOT_BUTTON = "bot_button"


class ObligationType(str, enum.Enum):
    REGISTRATION = "registration"                # постановка на миграционный учёт
    CONTRACT_NOTICE = "contract_notice"           # уведомление МВД о заключении договора
    CONTRACT_TERMINATION_NOTICE = "contract_termination_notice"
    MEDICAL_EXAM = "medical_exam"
    PATENT_PAYMENT = "patent_payment"             # не используется в MVP, зарезервировано
    EFS1_REPORT = "efs1_report"                   # ЕФС-1 в СФР — не позднее следующего рабочего дня
                                                   # после приказа о приёме / даты договора
    DACTYLOSCOPY = "dactyloscopy"                 # дактилоскопия + фотографирование ("грин карта") —
                                                   # разовая, 30 календарных дней с даты въезда
                                                   # (п.13 ст.5 №115-ФЗ). Закрывается внесением
                                                   # employee.dactyloscopy_date. Фотографирование —
                                                   # та же процедура и карта, отдельным типом НЕ заводится.
    REGISTRATION_RENEWAL = "registration_renewal"  # продление по правилу "90 из 180" — периодическое,
                                                    # создаётся отдельным cron-скриптом, не разовым триггером
    DEPARTURE_NOTICE = "departure_notice"          # снятие с миграционного учёта (уведомление об убытии)
                                                    # при увольнении. Принимающая сторона (вахта) обязана
                                                    # подать в 7 рабочих дней с даты убытия (ст.23 №109-ФЗ,
                                                    # п.45 Правил ПП №9). Работодатель ТСМ — принимающая
                                                    # сторона. Пропуск -> риск обвинения в фиктивной постановке.


class DeadlineUnit(str, enum.Enum):
    CALENDAR_DAY = "calendar_day"
    WORKING_DAY = "working_day"


class ObligationStatus(str, enum.Enum):
    PENDING = "pending"
    DONE = "done"
    OVERDUE = "overdue"


class ExamStatus(str, enum.Enum):
    NOT_STARTED = "not_started"
    REFERRED = "referred"
    IN_PROGRESS = "in_progress"
    COMPLETED = "completed"


class PaymentStatus(str, enum.Enum):
    UNPAID = "unpaid"
    INVOICED = "invoiced"
    PAID = "paid"


class DocumentType(str, enum.Enum):
    PASSPORT_TRANSLATION = "passport_translation"
    MEDICAL_CERTIFICATE = "medical_certificate"
    REGISTRATION_PROOF = "registration_proof"


class Employee(Base):
    __tablename__ = "employees"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    full_name: Mapped[str] = mapped_column(String, nullable=False)
    citizenship: Mapped[str | None] = mapped_column(String, nullable=True)
    category: Mapped[Category | None] = mapped_column(Enum(Category), nullable=True)

    entry_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    contract_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    contract_end_date: Mapped[date | None] = mapped_column(Date, nullable=True)  # дата увольнения
    # (расторжения договора). Ставится при оформлении увольнения; создаёт обязательства
    # CONTRACT_TERMINATION_NOTICE (+3 раб.дня) и DEPARTURE_NOTICE (+7 раб.дней).

    language: Mapped[str] = mapped_column(String, default="ru")  # ISO-код, напр. 'kk', 'ru'
    phone: Mapped[str | None] = mapped_column(String, nullable=True)  # свой номер сотрудника, не корпоративный

    consent_status: Mapped[ConsentStatus] = mapped_column(
        Enum(ConsentStatus), default=ConsentStatus.DRAFT, nullable=False
    )

    # --- поля, перенесённые из ручной xlsx-таблицы (2026-07) ---
    birth_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    passport_series: Mapped[str | None] = mapped_column(String, nullable=True)
    passport_number: Mapped[str | None] = mapped_column(String, nullable=True)
    address: Mapped[str | None] = mapped_column(Text, nullable=True)
    entry_country: Mapped[str | None] = mapped_column(String, nullable=True)  # "откуда въехал"
    # Свободный текст, не дата: в таблице это либо дата+заметка ("25.07.2026 Хостел"), либо пусто.
    # Не приводим к Date намеренно — иначе теряем текстовую часть без явного запроса на это.
    registration_deadline_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    # Тоже свободный текст ("срочно надо делать") — статус медкомиссии ещё не формализован в enum.
    medical_exam_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    employment_status: Mapped[str | None] = mapped_column(String, nullable=True)

    # --- поле для версионирования REGISTRATION при смене места пребывания (2026-07) ---
    # Дата, с которой действует ТЕКУЩЕЕ значение address. Это trigger_field для второго
    # правила REGISTRATION в deadlines.py — по закону смена адреса требует новой постановки
    # на учёт с тем же сроком, что при первичном въезде (см. комментарий в deadlines.py).
    # NULL означает "адрес не менялся с момента первичного въезда" — правило просто
    # пропускается в create_obligations_for_employee, как и любое другое пустое trigger_field.
    address_since: Mapped[date | None] = mapped_column(Date, nullable=True)

    # --- дата прохождения дактилоскопии (2026-07) ---
    # NULL = не пройдена: обязанность DACTYLOSCOPY активна и горит от entry_date+30.
    # Заполнено = пройдена: обработчик в webforms.py переводит текущую DACTYLOSCOPY в DONE.
    # Это НЕ trigger_field (триггер — entry_date), а поле ЗАКРЫТИЯ обязанности, по аналогии
    # с тем, как медосмотр закрывается результатом Referral. Дата разовая: карта на 10 лет,
    # ежегодного повтора нет. Точная дата у уже прошедших может быть неизвестна — для закрытия
    # достаточно любой корректной; юридически разовую карту по дате мы не обязаны хранить.
    # На боевой БД колонку добавить вручную (create_all не меняет существующие таблицы):
    #   ALTER TABLE employees ADD COLUMN dactyloscopy_date DATE;
    dactyloscopy_date: Mapped[date | None] = mapped_column(Date, nullable=True)

    # Табельный номер (2026-07). Присваивается работнику при трудоустройстве, ведётся во
    # внешнем табеле. Идёт в номер трудового договора (формат "БК-ПСМ-{tab_number}").
    # Заполняется разовым SQL-импортом из табеля по ФИО; у новых — вручную по мере оформления.
    # На боевой БД добавить колонку вручную (create_all не меняет существующие таблицы):
    #   ALTER TABLE employees ADD COLUMN tab_number VARCHAR;
    tab_number: Mapped[str | None] = mapped_column(String, nullable=True)

    # Статус миграционного учёта (2026-07). Управляет расчётом дедлайнов (см. RegistrationStatus).
    # NULL = не задан → создание обязательств заблокировано. Кадровик проставляет вручную.
    # На боевой БД добавить колонку вручную (create_all не меняет существующие таблицы):
    #   ALTER TABLE employees ADD COLUMN registration_status VARCHAR;
    registration_status: Mapped[RegistrationStatus | None] = mapped_column(
        Enum(RegistrationStatus), nullable=True
    )

    created_by: Mapped[str | None] = mapped_column(String, nullable=True)  # кто завёл (кадровик/прораб)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    consents: Mapped[list["Consent"]] = relationship(back_populates="employee", cascade="all, delete-orphan")
    obligations: Mapped[list["Obligation"]] = relationship(back_populates="employee", cascade="all, delete-orphan")
    documents: Mapped[list["Document"]] = relationship(back_populates="employee", cascade="all, delete-orphan")
    registration_periods: Mapped[list["RegistrationPeriod"]] = relationship(
        back_populates="employee", cascade="all, delete-orphan"
    )
    referrals: Mapped[list["Referral"]] = relationship(back_populates="employee", cascade="all, delete-orphan")


class RegistrationPeriod(Base):
    """Периоды учёта по правилу '90 дней из скользящих 180' (см. п.9 ст.97 Договора о ЕАЭС).

    Не хранит полную историю пересечений границы — только периоды, заведённые ботом.
    Если сотрудник уже въезжал и выезжал ДО того, как попал в систему, это не учитывается
    автоматически: при первом занесении сотрудника это нужно уточнить у кадровика вручную,
    иначе period_end может быть посчитан оптимистичнее, чем есть на самом деле по закону."""

    __tablename__ = "registration_periods"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)

    period_start: Mapped[date] = mapped_column(Date, nullable=False)
    period_end: Mapped[date] = mapped_column(Date, nullable=False)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)

    employee: Mapped["Employee"] = relationship(back_populates="registration_periods")


class SystemFlag(Base):
    """Разовые системные флаги, не привязанные к конкретному сотруднику.

    TODO удалить после разового прогона бэкфилла дактилоскопии:
    сейчас используется только флаг key="dactyloscopy_backfill_done" — эндпоинт
    /admin/recompute-dactyloscopy в webforms.py ставит его после успешного прогона
    create_obligations_for_employee по всем подтверждённым сотрудникам, чтобы
    правило DACTYLOSCOPY создало обязанности задним числом для уже заведённых людей.
    После прогона эндпоинт и эту таблицу можно убрать отдельной правкой."""

    __tablename__ = "system_flags"

    key: Mapped[str] = mapped_column(String, primary_key=True)
    value: Mapped[str | None] = mapped_column(String, nullable=True)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class NotificationSubscriber(Base):
    """chat_id получателей проактивных напоминаний. Заполняется при /start —
    без этого cron-скрипту некому слать сообщения о горящих дедлайнах."""

    __tablename__ = "notification_subscribers"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    chat_id: Mapped[str] = mapped_column(String, nullable=False, unique=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class User(Base):
    """Пользователь системы (прораб/кадровик/админ). Логин — номер телефона (уникальный).
    Пароль хранится ТОЛЬКО хэшем (bcrypt через passlib), открытый пароль нигде не сохраняется.
    Регистрация открытая: заявитель указывает телефон, пароль, ФИО; роль назначает админ при
    одобрении. До одобрения status=PENDING, вход заблокирован.
    На боевой БД создать таблицу (create_all создаёт новые таблицы — при первом деплое модели
    таблица появится сама; если нет — CREATE TABLE вручную)."""

    __tablename__ = "users"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    phone: Mapped[str] = mapped_column(String, nullable=False, unique=True)  # логин
    password_hash: Mapped[str] = mapped_column(String, nullable=False)       # bcrypt, не открытый
    full_name: Mapped[str] = mapped_column(String, nullable=False)
    role: Mapped[UserRole | None] = mapped_column(Enum(UserRole), nullable=True)  # назначает админ
    status: Mapped[UserStatus] = mapped_column(
        Enum(UserStatus), default=UserStatus.PENDING, nullable=False
    )
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    approved_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    # Защита от подбора пароля: 5 неудачных попыток подряд -> блокировка на час (locked_until).
    # Успешный вход обнуляет счётчик. Админ снимает блокировку раньше кнопкой в админке.
    # Это ОТДЕЛЬНО от status=BLOCKED (ручной отзыв доступа админом навсегда).
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    locked_until: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


class Consent(Base):
    __tablename__ = "consents"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)

    method: Mapped[ConsentMethod] = mapped_column(Enum(ConsentMethod), nullable=False)
    proof: Mapped[str] = mapped_column(String, nullable=False)  # file_id скана ИЛИ id callback-события MAX
    consent_text_version: Mapped[str] = mapped_column(String, nullable=False)
    confirmed_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    employee: Mapped["Employee"] = relationship(back_populates="consents")


class Obligation(Base):
    __tablename__ = "obligations"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)

    type: Mapped[ObligationType] = mapped_column(Enum(ObligationType), nullable=False)
    trigger_date: Mapped[date] = mapped_column(Date, nullable=False)
    deadline_value: Mapped[int] = mapped_column(nullable=False)
    deadline_unit: Mapped[DeadlineUnit] = mapped_column(Enum(DeadlineUnit), nullable=False)
    deadline_date: Mapped[date] = mapped_column(Date, nullable=False)  # вычисляется при создании

    status: Mapped[ObligationStatus] = mapped_column(
        Enum(ObligationStatus), default=ObligationStatus.PENDING, nullable=False
    )

    # Версионирование (2026-07): True — это актуальная запись для (employee_id, type).
    # При создании новой версии (смена адреса, исправление даты и т.п.) старая помечается
    # False, а не удаляется — история остаётся в базе. Все запросы, считающие "активные"
    # или "просроченные" обязательства, ДОЛЖНЫ фильтровать is_current=True, иначе старые
    # версии задваивают списки. См. create_obligations_for_employee в obligations.py.
    is_current: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)

    employee: Mapped["Employee"] = relationship(back_populates="obligations")
    referrals: Mapped[list["Referral"]] = relationship(back_populates="obligation", cascade="all, delete-orphan")


class Referral(Base):
    __tablename__ = "referrals"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    obligation_id: Mapped[str] = mapped_column(ForeignKey("obligations.id"), nullable=False)

    clinic_id: Mapped[str] = mapped_column(String, nullable=False)  # напр. 'pirogova_murmansk'
    referral_date: Mapped[date] = mapped_column(Date, default=date.today)
    exam_status: Mapped[ExamStatus] = mapped_column(Enum(ExamStatus), default=ExamStatus.NOT_STARTED)
    result_date: Mapped[date | None] = mapped_column(Date, nullable=True)

    obligation: Mapped["Obligation"] = relationship(back_populates="referrals")
    employee: Mapped["Employee"] = relationship(back_populates="referrals")
    invoices: Mapped[list["Invoice"]] = relationship(back_populates="referral", cascade="all, delete-orphan")


class Invoice(Base):
    __tablename__ = "invoices"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    referral_id: Mapped[str] = mapped_column(ForeignKey("referrals.id"), nullable=False)

    clinic_id: Mapped[str] = mapped_column(String, nullable=False)
    payer_id: Mapped[str] = mapped_column(String, nullable=False)  # напр. 'ip_buts_sy' — реквизиты отдельно
    amount: Mapped[float] = mapped_column(Numeric(10, 2), nullable=False)
    payment_status: Mapped[PaymentStatus] = mapped_column(
        Enum(PaymentStatus), default=PaymentStatus.UNPAID
    )
    invoice_document: Mapped[str | None] = mapped_column(String, nullable=True)  # file_id сгенерированного счёта

    referral: Mapped["Referral"] = relationship(back_populates="invoices")


class Document(Base):
    __tablename__ = "documents"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)

    type: Mapped[DocumentType] = mapped_column(Enum(DocumentType), nullable=False)
    file_id: Mapped[str] = mapped_column(String, nullable=False)  # ссылка/id вложения в MAX
    uploaded_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    verified: Mapped[bool] = mapped_column(Boolean, default=False)  # ручная верификация кадровиком

    employee: Mapped["Employee"] = relationship(back_populates="documents")
