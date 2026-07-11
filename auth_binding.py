"""
auth_binding.py — привязка MAX-аккаунта (event.message.sender.user_id) к
записи User (роль/права), и проверка прав перед действиями в боте.

2026-07: слияние с ботом ТабельБелокаменка. У User нет своего понятия
"MAX-чат" — это телефон+пароль для веб-формы кадровика. Команда /login
в самом MAX-боте один раз связывает конкретный MAX user_id с конкретным
User по номеру телефона, дальше бот узнаёт роль по этой связке, не
спрашивая телефон заново при каждом действии.

Требует колонку users.max_user_id (см. models.py, ALTER TABLE на проде).
"""

import logging

from sqlalchemy.orm import Session

from models import User, UserStatus
from common_utils import normalize_phone

log = logging.getLogger("auth_binding")


def find_user_by_phone(session: Session, phone: str) -> User | None:
    phone = normalize_phone(phone)
    return session.query(User).filter_by(phone=phone).first()


def find_user_by_max_id(session: Session, max_user_id: str) -> User | None:
    return session.query(User).filter_by(max_user_id=str(max_user_id)).first()


def bind_max_account(session: Session, phone: str, max_user_id: str) -> tuple[bool, str]:
    """
    Пытается привязать MAX-аккаунт max_user_id к пользователю с телефоном phone.
    Возвращает (успех, текст_для_пользователя).

    Не привязывает молча, если:
      - телефон не найден,
      - заявка не одобрена (PENDING/BLOCKED),
      - этот телефон уже привязан к ДРУГОМУ max_user_id (защита от случайного
        перехвата чужого аккаунта — переподключение делает админ отдельно,
        не эта функция).
    """
    user = find_user_by_phone(session, phone)
    if user is None:
        return False, "Такого номера нет в системе. Обратитесь к кадровику для регистрации."

    if user.status == UserStatus.PENDING:
        return False, "Заявка ещё не одобрена админом. Попробуйте позже."
    if user.status == UserStatus.BLOCKED:
        return False, "Доступ заблокирован. Обратитесь к админу."

    if user.max_user_id is not None and user.max_user_id != str(max_user_id):
        return False, (
            "Этот номер уже привязан к другому MAX-аккаунту. "
            "Если это ошибка — обратитесь к админу для переподключения."
        )

    if user.max_user_id == str(max_user_id):
        return True, (f"Вы уже вошли как {user.full_name}.\n"
                       f"Рабочее место: Автоматизированная система учёта на производстве. Роль: "
                       f"{user.role.value if user.role else '—'}.")

    user.max_user_id = str(max_user_id)
    session.add(user)
    session.commit()
    return True, (f"Готово, вы вошли как {user.full_name}.\n"
                   f"Рабочее место: Автоматизированная система учёта на производстве. Роль: "
                   f"{user.role.value if user.role else '—'}.")


def get_role_label(user: User) -> str:
    if user.role is None:
        return "—"
    return user.role.value


# ================= Код подтверждения MAX при веб-регистрации =================

CONFIRM_CODE_TTL_MINUTES = 30


def generate_max_confirm_code(session: Session, user: User) -> str:
    """Генерирует 6-значный код для привязки MAX после регистрации через веб.
    Веб не знает MAX-аккаунт человека напрямую — человек присылает этот код
    боту командой /confirm <код>, см. confirm_max_code ниже."""
    import random
    from datetime import datetime, timedelta

    code = f"{random.randint(0, 999999):06d}"
    user.pending_max_code = code
    user.pending_max_code_expires = datetime.utcnow() + timedelta(minutes=CONFIRM_CODE_TTL_MINUTES)
    session.add(user)
    session.commit()
    return code


def confirm_max_code(session: Session, code: str, max_user_id: str) -> tuple[bool, str]:
    """Обрабатывает /confirm <код> в боте — находит User по коду, проверяет срок
    действия, привязывает max_user_id. Код одноразовый — очищается сразу после
    использования, успешного или нет (истёкший код нельзя вводить повторно)."""
    from datetime import datetime

    code = code.strip()
    user = session.query(User).filter_by(pending_max_code=code).first()
    if user is None:
        return False, "Код не найден или уже использован. Проверьте, что ввели верно."

    expired = user.pending_max_code_expires is None or user.pending_max_code_expires < datetime.utcnow()
    user.pending_max_code = None
    user.pending_max_code_expires = None
    if expired:
        session.add(user)
        session.commit()
        return False, "Код истёк (действует 30 минут). Зарегистрируйтесь заново на сайте."

    if user.max_user_id is not None and user.max_user_id != str(max_user_id):
        session.add(user)
        session.commit()
        return False, "Этот аккаунт уже привязан к другому MAX. Обратитесь к админу."

    user.max_user_id = str(max_user_id)
    session.add(user)
    session.commit()

    if user.status == UserStatus.PENDING:
        return True, (f"Готово, MAX привязан. Заявка {user.full_name} ещё ожидает "
                       f"одобрения админом — доступ откроется после этого.")
    return True, (f"Готово, вы вошли как {user.full_name}.\n"
                   f"Рабочее место: Автоматизированная система учёта на производстве. Роль: {get_role_label(user)}.")


# ================= Регистрация ТОЛЬКО через MAX (без веба) =================

def register_via_max(session: Session, full_name: str, phone: str,
                      max_user_id: str) -> tuple[bool, str]:
    """
    Регистрация с нуля прямо в боте (человек никогда не заходил в веб).
    Без пароля (password_hash=NULL) — этому человеку он не нужен, пока он
    сам не решит логиниться в веб (тогда сброс пароля сделает админ).
    max_user_id привязывается СРАЗУ — /login потом не нужен.

    Не создаёт дубль, если телефон уже занят существующей записью — вместо
    этого пытается привязать MAX к НЕЙ (тот же путь, что bind_max_account),
    чтобы не плодить две заявки на одного человека.
    """
    phone_norm = normalize_phone(phone)
    if not phone_norm or len(phone_norm) < 10:
        return False, "Не распознал номер телефона. Попробуйте ещё раз."

    existing = find_user_by_phone(session, phone_norm)
    if existing is not None:
        # Телефон уже занят — не плодим вторую заявку, пробуем привязать к этой же.
        return bind_max_account(session, phone_norm, max_user_id)

    user = User(
        phone=phone_norm,
        password_hash=None,
        full_name=full_name.strip(),
        status=UserStatus.PENDING,
        max_user_id=str(max_user_id),
    )
    session.add(user)
    session.commit()
    return True, (f"Заявка отправлена, {user.full_name}. MAX уже привязан — как только "
                   f"админ одобрит и назначит роль, доступ откроется сам, ничего "
                   f"вводить не придётся.")
