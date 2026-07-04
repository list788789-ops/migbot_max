"""
Модели БД для бота миграционного учёта.

Ключевые решения, зафиксированные в диалоге с заказчиком (не менять без причины):
- category у employee — enum с полным набором значений с самого начала (eaeu/patent/visa/hqs),
  даже если сейчас используется только 'eaeu'. Это дёшево сейчас и дорого добавлять потом.
- obligations.deadline_unit различает calendar/working дни — ЕАЭС считается в календарных днях (30),
  уведомление о договоре — в рабочих (3). Смешение этих единиц — источник ошибок в дедлайнах.
- consent_status блокирует создание obligations, пока не confirmed (см. bot.py, on_consent_confirmed).
- invoices разделяет clinic_id (куда идёт сотрудник) и payer_id (кто платит) — это разные сущности,
  заказчик услуги (ИП) не обязан совпадать с клиникой.
"""

from __future__ import annotations

import enum
import uuid
from datetime import date, datetime

from sqlalchemy import (
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
    citizenship: Mapped[str] = mapped_column(String, nullable=False)
    category: Mapped[Category] = mapped_column(Enum(Category), nullable=False, default=Category.EAEU)

    entry_date: Mapped[date] = mapped_column(Date, nullable=False)
    contract_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    contract_end_date: Mapped[date | None] = mapped_column(Date, nullable=True)

    language: Mapped[str] = mapped_column(String, default="ru")  # ISO-код, напр. 'kk', 'ru'
    phone: Mapped[str | None] = mapped_column(String, nullable=True)  # свой номер сотрудника, не корпоративный

    consent_status: Mapped[ConsentStatus] = mapped_column(
        Enum(ConsentStatus), default=ConsentStatus.DRAFT, nullable=False
    )

    created_by: Mapped[str | None] = mapped_column(String, nullable=True)  # кто завёл (кадровик/прораб)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

    consents: Mapped[list["Consent"]] = relationship(back_populates="employee")
    obligations: Mapped[list["Obligation"]] = relationship(back_populates="employee")
    documents: Mapped[list["Document"]] = relationship(back_populates="employee")


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

    employee: Mapped["Employee"] = relationship(back_populates="obligations")
    referrals: Mapped[list["Referral"]] = relationship(back_populates="obligation")


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
    invoices: Mapped[list["Invoice"]] = relationship(back_populates="referral")


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
