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
    UniqueConstraint,
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
    """Роли пользователей системы. PRORAB — чтение и скачивание документов (договор,
    квитанции, отчёты) БЕЗ записи в основные таблицы (employees/obligations/consents и
    т.д.), НО с узким исключением: разметка явки в attendance_marks (Утро/Вечер/Причины/
    Межвахта, см. tabel.py) — это единственное, что PRORAB может писать. Решение
    зафиксировано в 2026-07 при слиянии с ботом ТабельБелокаменка: разметка явки
    исторически была основной функцией прораба, вынести её в отдельную роль сочли
    избыточным (это тот же самый человек и та же обязанность, а не новая). KADROVIK —
    всё по работникам (заключение, статусы, правки), но НЕ управление пользователями.
    ADMIN — всё + одобрение заявок и управление пользователями. Роль назначает админ
    при одобрении (заявитель её не выбирает)."""
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
    # Снято при увольнении: обязательство стало неактуальным, т.к. работник уволен и убыл.
    # НЕ то же, что DONE (не исполнено) и не OVERDUE (не должно висеть в просроченных).
    # Все выборки горящих/просроченных обязаны исключать CANCELLED.
    # Требует ALTER TYPE на проде: ALTER TYPE obligationstatus ADD VALUE 'CANCELLED';
    CANCELLED = "cancelled"


class ExamStatus(str, enum.Enum):
    NOT_STARTED = "not_started"
    REFERRED = "referred"
    IN_PROGRESS = "in_progress"
    COMPLETED = "completed"


class PaymentStatus(str, enum.Enum):
    UNPAID = "unpaid"
    INVOICED = "invoiced"
    PAID = "paid"


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
    # ИИН (12 цифр, Казахстан). Из MRZ удостоверения или вручную. Требует ALTER на проде:
    #   ALTER TABLE employees ADD COLUMN iin VARCHAR;
    iin: Mapped[str | None] = mapped_column(String, nullable=True)
    # Вид документа: "id" (удостоверение TD1) или "passport" (загранпаспорт TD3). Из OCR.
    #   ALTER TABLE employees ADD COLUMN doc_type VARCHAR;
    doc_type: Mapped[str | None] = mapped_column(String, nullable=True)
    # Чекбокс: все страницы паспорта загружены (система не считает страницы, подтверждает кадровик).
    #   ALTER TABLE employees ADD COLUMN passport_all_pages BOOLEAN DEFAULT FALSE;
    passport_all_pages: Mapped[bool] = mapped_column(Boolean, default=False)
    # СНИЛС. Нужен для корректного ЕФС-1 (реквизит). ЕФС-1 подаётся в срок независимо от СНИЛС
    #   (срок ЕФС-1 жёсткий — 1 раб. день от договора; СНИЛС долгий), при появлении -> корректировка.
    #   ALTER TABLE employees ADD COLUMN snils VARCHAR;
    snils: Mapped[str | None] = mapped_column(String, nullable=True)
    # Вид процедуры получения СНИЛС: "new" — первичное, "merge" — объединение дублей (было
    #   несколько СНИЛС от прошлых въездов, СФР сливает в один).
    #   ALTER TABLE employees ADD COLUMN snils_procedure VARCHAR;
    snils_procedure: Mapped[str | None] = mapped_column(String, nullable=True)
    # Дата записи в СФР на получение/объединение СНИЛС (если СНИЛС ещё нет).
    #   ALTER TABLE employees ADD COLUMN snils_appointment_date DATE;
    snils_appointment_date: Mapped[date | None] = mapped_column(Date, nullable=True)
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

    # --- дата окончания срока регистрации по адресу (2026-07) ---
    # На Госуслугах в уведомлении о прибытии печатается НЕ дата начала, а «срок пребывания до»
    # (дата окончания). Это поле хранит её — то, что реально стоит в отрывной части уведомления.
    # Справочное: НЕ триггер обязательства (постановку считает address_since), а фактическая
    # дата из документа для сверки и напоминания о продлении. Требует ALTER TABLE на проде:
    #   ALTER TABLE employees ADD COLUMN registration_valid_until DATE;
    registration_valid_until: Mapped[date | None] = mapped_column(Date, nullable=True)

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
    registration_periods: Mapped[list["RegistrationPeriod"]] = relationship(
        back_populates="employee", cascade="all, delete-orphan"
    )
    referrals: Mapped[list["Referral"]] = relationship(back_populates="employee", cascade="all, delete-orphan")
    attendance_marks: Mapped[list["AttendanceMark"]] = relationship(
        back_populates="employee", cascade="all, delete-orphan"
    )


class WorkOrderStatus(str, enum.Enum):
    DRAFT = "draft"
    ACTIVE = "active"
    CLOSED = "closed"
    CANCELLED = "cancelled"


class InstructionType(str, enum.Enum):
    """Виды инструктажей по охране труда — все нужны и учитываются отдельно
    (2026-07, модуль «Производство»)."""
    INTRODUCTORY = "introductory"              # вводный — один раз при приёме
    PRIMARY_WORKPLACE = "primary_workplace"    # первичный на рабочем месте
    REPEATED = "repeated"                      # повторный, регулярный
    UNSCHEDULED = "unscheduled"                # внеплановый
    TARGETED = "targeted"                      # целевой (под конкретную задачу)


class Brigade(Base):
    """
    Бригада (2026-07) — сохраняемый состав, отдельно от конкретного наряда-допуска.
    По договорённости: раньше при каждом новом наряде состав отмечался заново
    чекбоксами; теперь можно один раз завести бригаду и дальше выбирать её целиком.
    Состав можно менять со временем (не история версий — просто текущий список,
    см. BrigadeMember); у самого наряда-допуска состав фиксируется отдельно в
    WorkOrderMember на момент создания — смена состава бригады ПОСЛЕ не меняет
    задним числом уже выданные наряды.
    """

    __tablename__ = "brigades"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    name: Mapped[str] = mapped_column(String, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    members: Mapped[list["BrigadeMember"]] = relationship(
        back_populates="brigade", cascade="all, delete-orphan"
    )


class BrigadeMember(Base):
    __tablename__ = "brigade_members"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    brigade_id: Mapped[str] = mapped_column(ForeignKey("brigades.id"), nullable=False)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)

    brigade: Mapped["Brigade"] = relationship(back_populates="members")
    employee: Mapped["Employee"] = relationship()


class WorkOrder(Base):
    """
    Наряд-допуск (2026-07, модуль «Производство», отдельный от миграционного
    учёта — своя доменная область, охрана труда). Официальный документ на
    производство работ.

    2026-07 (пересмотр по образцу реального бланка ООО «Промстроймонтаж»,
    наряд №25 на работы на высоте — фото прислал пользователь): структура
    оказалась заметно сложнее первой версии. Ключевые поправки:

    - ДВА разных ответственных, не один: responsible_supervisor_id
      ("Ответственный руководитель работ" — уровень выше, например старший
      производитель работ) и responsible_executor_id ("Ответственный
      исполнитель работ" — бригадир). В реальном бланке это разные люди с
      разными группами допуска, путать нельзя.
    - Наряд действует НЕСКОЛЬКО ДНЕЙ (в примере — 9 дней), и каждый день
      регистрируется ОТДЕЛЬНО в разделе "Ежедневный допуск к работе" —
      см. WorkOrderDailyAdmission ниже. Один наряд ≠ одна явка.
    - Материалы/инструменты/приспособления/спецтехника, ссылка на
      технологическую карту, системы обеспечения безопасности (страховочные/
      поддерживающие/эвакуационные) — были в реальном бланке, у меня не было.

    Реализовано в отдельном файле production.py — из bot.py/webforms.py только
    пункты меню/ссылки, минимальная связанность с основной системой (по
    договорённости — "тестовый режим в отдельных файлах").
    """

    __tablename__ = "work_orders"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    number: Mapped[str] = mapped_column(String, nullable=False)  # номер наряда для документа
    subdivision: Mapped[str | None] = mapped_column(String, nullable=True)  # "ОС" и т.п.
    work_description: Mapped[str] = mapped_column(Text, nullable=False)
    location: Mapped[str] = mapped_column(String, nullable=False)

    responsible_supervisor_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    responsible_executor_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    issued_by: Mapped[str] = mapped_column(String, nullable=False)  # User.id кадровика/мастера

    valid_from: Mapped[date] = mapped_column(Date, nullable=False)
    valid_to: Mapped[date] = mapped_column(Date, nullable=False)
    status: Mapped[WorkOrderStatus] = mapped_column(
        Enum(WorkOrderStatus), default=WorkOrderStatus.DRAFT, nullable=False
    )
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    # Материалы/инструменты/приспособления/спецтехника — по образцу реального бланка,
    # там это отдельные строки (см. фото: "Материалы: Арматура, опалубка, доска." и т.д.).
    materials: Mapped[str | None] = mapped_column(Text, nullable=True)
    tools: Mapped[str | None] = mapped_column(Text, nullable=True)
    equipment: Mapped[str | None] = mapped_column(Text, nullable=True)  # приспособления
    special_machinery: Mapped[str | None] = mapped_column(Text, nullable=True)  # спецтехника
    technological_card_ref: Mapped[str | None] = mapped_column(String, nullable=True)  # "Шифр ТК: НУЛ-ПСМ-ТК-07-24.005.3"
    safety_systems: Mapped[str | None] = mapped_column(Text, nullable=True)  # страховочные/поддерживающие/эвакуационные
    special_conditions: Mapped[str | None] = mapped_column(Text, nullable=True)  # погодные ограничения и т.п.

    responsible_supervisor: Mapped["Employee"] = relationship(foreign_keys=[responsible_supervisor_id])
    responsible_executor: Mapped["Employee"] = relationship(foreign_keys=[responsible_executor_id])
    members: Mapped[list["WorkOrderMember"]] = relationship(
        back_populates="work_order", cascade="all, delete-orphan"
    )
    daily_admissions: Mapped[list["WorkOrderDailyAdmission"]] = relationship(
        back_populates="work_order", cascade="all, delete-orphan"
    )


class WorkOrderDailyAdmission(Base):
    """
    Ежедневный допуск к работе (2026-07, по образцу реального бланка — раздел
    "6. Ежедневный допуск к работе"). Наряд-допуск может действовать несколько
    дней (в примере — 9), и КАЖДЫЙ день требует отдельной регистрации: бригада
    получила целевой инструктаж (дата/время начала, ответственный руководитель
    подтверждает), и отдельно — работа на этот день полностью закончена
    (дата/время окончания, ответственный исполнитель подтверждает).

    Это НЕ то же самое, что подпись члена бригады в WorkOrderMember (та —
    разовое ознакомление с условиями при выдаче наряда). Здесь — суточный цикл
    допуск/окончание, повторяется каждый рабочий день в пределах valid_from..valid_to.
    """

    __tablename__ = "work_order_daily_admissions"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    work_order_id: Mapped[str] = mapped_column(ForeignKey("work_orders.id"), nullable=False)
    admission_date: Mapped[date] = mapped_column(Date, nullable=False)

    briefing_time: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)  # выдан целевой инструктаж
    briefing_confirmed_by: Mapped[str | None] = mapped_column(String, nullable=True)  # ответственный руководитель
    completion_time: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)  # работа закончена, место убрано
    completion_confirmed_by: Mapped[str | None] = mapped_column(String, nullable=True)  # ответственный исполнитель

    work_order: Mapped["WorkOrder"] = relationship(back_populates="daily_admissions")


class WorkOrderMember(Base):
    """Член бригады по наряду-допуску + подтверждение ознакомления (подпись).
    signed_at=NULL — ещё не подтвердил, что ознакомлен с условиями работ."""

    __tablename__ = "work_order_members"


    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    work_order_id: Mapped[str] = mapped_column(ForeignKey("work_orders.id"), nullable=False)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    signed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    work_order: Mapped["WorkOrder"] = relationship(back_populates="members")
    employee: Mapped["Employee"] = relationship()


class Instruction(Base):
    """Инструктаж по охране труда (2026-07, модуль «Производство»). Все виды
    учитываются — см. InstructionType. next_due_date — для повторных, когда
    ждать следующий (используется для напоминаний по аналогии с обязательствами
    миграционного учёта, см. production.get_due_instructions).

    2026-07 (второй заход): journal_row_number/printed_at — механизм допечатки
    журнала партиями (см. production.print_new_journal_entries). Номер строки
    присваивается ПОСЛЕДОВАТЕЛЬНО, отдельно по каждому InstructionType (это
    разные физические журналы — вводный, на рабочем месте, целевой — со своей
    нумерацией). NULL = запись ещё не попала ни в одну распечатанную партию."""

    __tablename__ = "instructions"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    type: Mapped[InstructionType] = mapped_column(Enum(InstructionType), nullable=False)
    topic: Mapped[str | None] = mapped_column(String, nullable=True)  # тема — особенно для целевого/внепланового
    conducted_by: Mapped[str] = mapped_column(String, nullable=False)  # кто провёл (ФИО или User.id)
    conducted_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    employee_confirmed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    next_due_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    journal_row_number: Mapped[int | None] = mapped_column(Integer, nullable=True)
    printed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    employee: Mapped["Employee"] = relationship()


class OrderCategory(str, enum.Enum):
    """Раздел, к которому относится приказ — для сортировки/фильтрации реестра."""
    PERSONNEL = "personnel"      # кадровые (приём, увольнение, назначения)
    PRODUCTION = "production"    # производственные (наряды, инструктажи, ОТ)
    OTHER = "other"              # прочие


class InternalOrder(Base):
    """
    Реестр внутренних приказов организации (2026-07, модуль «Производство»).
    Начат с приказа №20-ПСМ/2026 — более ранние приказы (1-19), если есть,
    в реестр не заведены (не было этой системы на момент их издания).

    Формат номера — по договорённости: "{порядковый}-{код заказчика}/{год}",
    например "20-ПСМ/2026". Код заказчика/номер — не автоматизированы жёстко,
    вводятся текстом при создании записи (не все приказы обязательно будут
    иметь этот формат — например, кадровые приказы могут нумероваться иначе).

    Используется как ссылка (INTERNAL_ORDER_REF) в футерах печатных бланков
    наряда-допуска и журналов инструктажа — см. production.get_latest_order_ref.
    Скан самого приказа хранится в S3 (тот же механизм, что и Certificate.scan_key),
    scan_type = f"order_{order.id}".
    """

    __tablename__ = "internal_orders"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    number: Mapped[str] = mapped_column(String, nullable=False)  # "20-ПСМ/2026"
    order_date: Mapped[date] = mapped_column(Date, nullable=False)
    topic: Mapped[str] = mapped_column(String, nullable=False)
    category: Mapped[OrderCategory] = mapped_column(
        Enum(OrderCategory), default=OrderCategory.OTHER, nullable=False
    )
    note: Mapped[str | None] = mapped_column(Text, nullable=True)
    scan_key: Mapped[str | None] = mapped_column(String, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class Certificate(Base):
    """
    Удостоверение по профессии (2026-07, модуль «Производство») — "корочки":
    электробезопасность, стропальщик, верхолазные работы и т.д. Срок действия +
    скан отслеживаются (по договорённости — оба нужны, не только факт наличия).

    scan_key — ключ файла в S3 (тот же механизм, что уже использует s3_storage.py
    для сканов паспорта/документов сотрудника, только другой префикс, см.
    production.py).
    """

    __tablename__ = "certificates"

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    profession: Mapped[str] = mapped_column(String, nullable=False)  # "Электробезопасность IV группа" и т.п.
    issued_by_org: Mapped[str | None] = mapped_column(String, nullable=True)  # орган выдачи
    issue_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    expiry_date: Mapped[date | None] = mapped_column(Date, nullable=True)  # NULL = бессрочное
    scan_key: Mapped[str | None] = mapped_column(String, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    employee: Mapped["Employee"] = relationship()


class RotationReturn(Base):
    """
    Ожидаемая дата возврата с межвахты + флаг "требует внимания кадровика".

    2026-07: слияние с ботом ТабельБелокаменка. Прораб ставит МЖ без задержек —
    его работа не блокируется чужими просрочками. Но если на момент постановки
    у сотрудника были незакрытые обязательства (Obligation, is_current=True,
    status IN PENDING/OVERDUE) — flagged=True, и это попадает кадровику в раздел
    "⚠️ Требует внимания" (видит только KADROVIK/ADMIN, не PRORAB — см. UserRole).

    Одна строка на сотрудника (перезаписывается при новой межвахте) — история
    предыдущих межвахт не нужна, только "когда ждать в этот раз" и "решено ли".
    flagged НЕ снимается автоматически при закрытии обязательств — это
    сознательная задача кадровику ("уточнить причину межвахты"), должна быть
    явно закрыта им самим (reviewed_at/reviewed_by), а не молча исчезнуть,
    когда кто-то другой закрыл дедлайн.

    На боевой БД: если таблица уже была создана ДО добавления departure_type
    (create_all создаёт только отсутствующие таблицы, не добавляет колонки в
    существующие) — нужен ручной ALTER TABLE:
        ALTER TABLE rotation_returns ADD COLUMN departure_type VARCHAR;
    Если разворачиваешь БД с нуля — create_all создаст колонку сама.
    """

    __tablename__ = "rotation_returns"

    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), primary_key=True)
    # NULL = дата возврата ещё не известна ("заглушка нужно уточнить у прораба",
    # см. tabel.get_pending_clarification_rotations). До 2026-07 поле было NOT NULL —
    # ALTER TABLE rotation_returns ALTER COLUMN expected_return_date DROP NOT NULL;
    expected_return_date: Mapped[date | None] = mapped_column(Date, nullable=True)

    # 2026-07: тип отбытия на межвахту — от него зависит, какое юридическое
    # событие срабатывает при ФАКТИЧЕСКОМ возврате (см. tabel.apply_rotation_return):
    #   'abroad'   — пересёк границу РФ (едет домой) -> новая постановка на учёт
    #                (REGISTRATION), но БЕЗ повторной дактилоскопии/медосмотра
    #                (это разовые обязанности, не привязаны к повторному въезду).
    #   'domestic' — остался в РФ, но покидал место пребывания -> address_since
    #                обновляется датой возврата, дальше работает уже существующее
    #                правило REGISTRATION/address_since в deadlines.py.
    #   'none'     — физически не выезжал с площадки -> ничего не создаётся,
    #                регистрация не прерывалась.
    # NULL = не указано (старые записи до этого поля, или упрощённый флоу).
    departure_type: Mapped[str | None] = mapped_column(String, nullable=True)

    flagged: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    flagged_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    reviewed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    reviewed_by: Mapped[str | None] = mapped_column(String, nullable=True)  # User.id кадровика

    employee: Mapped["Employee"] = relationship()


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


class AttendanceMark(Base):
    """
    Отметка присутствия за день (перенесено из бота ТабельБелокаменка,
    Google Sheets). Один слот (день/ночь) на дату на сотрудника.

    2026-07: перенос истории из Sheets одноразовым скриптом
    (migrate_attendance_from_sheets.py) — created_by='migration_script'
    у перенесённых записей, отличает их от тех, что появятся из бота
    (created_by=user_id того, кто поставил).

    Код (day/night slot) — те же буквенные коды, что были в Sheets:
    Д/О/Б/МЖ/Н/МУ/В (дневной слот), НЧ/О (ночной слот). Намеренно не
    заведён отдельный enum под них здесь — это калька старой системы
    кодов, а не новая доменная модель; если понадобится собственная
    типизация статусов явки — заводить отдельно, не путать со старыми
    буквами.
    """

    __tablename__ = "attendance_marks"
    __table_args__ = (
        UniqueConstraint("employee_id", "mark_date", "slot", name="uq_attendance_employee_date_slot"),
    )

    id: Mapped[str] = mapped_column(String, primary_key=True, default=_uuid)
    employee_id: Mapped[str] = mapped_column(ForeignKey("employees.id"), nullable=False)
    # Снимок ФИО на момент отметки — для отчётов без лишнего JOIN.
    # НЕ ключ связи (ключ — employee_id), просто удобство чтения.
    employee_name_snap: Mapped[str] = mapped_column(String, nullable=False)

    mark_date: Mapped[date] = mapped_column(Date, nullable=False)
    slot: Mapped[str] = mapped_column(String, nullable=False)  # 'day' | 'night'
    code: Mapped[str] = mapped_column(String, nullable=False)  # Д/О/Б/МЖ/Н/МУ/В/НЧ

    created_by: Mapped[str] = mapped_column(String, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.utcnow, onupdate=datetime.utcnow
    )

    employee: Mapped["Employee"] = relationship(back_populates="attendance_marks")


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
    # 2026-07: nullable (было NOT NULL) — регистрация теперь возможна и ТОЛЬКО через
    # MAX, без пароля вообще (человек никогда не логинится в веб). Пароль обязателен
    # только тем, кто регистрируется через веб-форму. На боевой БД:
    #   ALTER TABLE users ALTER COLUMN password_hash DROP NOT NULL;
    password_hash: Mapped[str | None] = mapped_column(String, nullable=True)       # bcrypt, не открытый
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

    # Привязка MAX-аккаунта (2026-07, слияние с ботом ТабельБелокаменка).
    # user_id из события maxapi (event.message.sender.user_id) — числовой ID,
    # никак не связанный с телефоном сам по себе. Привязывается один раз через
    # команду /login <телефон> в самом боте: бот ищет User по телефону,
    # проверяет status=APPROVED, и если max_user_id ещё пуст — записывает сюда.
    # Дальше бот узнаёт роль/права человека по этому полю, не спрашивая телефон
    # заново при каждом обращении. NULL = ещё не привязан к MAX вообще.
    # На боевой БД добавить колонку вручную (create_all не меняет существующие
    # таблицы):
    #   ALTER TABLE users ADD COLUMN max_user_id VARCHAR UNIQUE;
    max_user_id: Mapped[str | None] = mapped_column(String, nullable=True, unique=True)

    # Код подтверждения привязки MAX при регистрации ЧЕРЕЗ ВЕБ (2026-07). Веб не
    # знает MAX-аккаунт человека напрямую — форма показывает код, человек шлёт его
    # боту командой /confirm <код>, бот находит User по коду (не истёкшему) и
    # привязывает max_user_id. Код одноразовый — очищается сразу после успешной
    # привязки, expires_at защищает от использования кода, если его кто-то перехватит
    # (например, скриншот регистрации попал не в те руки).
    # На боевой БД:
    #   ALTER TABLE users ADD COLUMN pending_max_code VARCHAR;
    #   ALTER TABLE users ADD COLUMN pending_max_code_expires TIMESTAMP;
    pending_max_code: Mapped[str | None] = mapped_column(String, nullable=True)
    pending_max_code_expires: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)


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

    # Ручная отметка о выполнении (2026-07): для обязательств, которые подаются ВОВНЕ и не имеют
    # своего закрывателя (ЕФС-1 в СФР, уведомление МВД, регистрация). Кадровик отмечает факт
    # подачи кнопкой в карточке; хранится дата отметки и кто отметил — след для проверки.
    # Требует ALTER TABLE на проде:
    #   ALTER TABLE obligations ADD COLUMN done_date DATE;
    #   ALTER TABLE obligations ADD COLUMN done_by VARCHAR;
    done_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    done_by: Mapped[str | None] = mapped_column(String, nullable=True)

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
