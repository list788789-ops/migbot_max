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

log = logging.getLogger("auth_binding")


def find_user_by_phone(session: Session, phone: str) -> User | None:
    phone = phone.strip()
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
        return True, f"Вы уже вошли как {user.full_name} ({user.role.value if user.role else '—'})."

    user.max_user_id = str(max_user_id)
    session.add(user)
    session.commit()
    return True, f"Готово, вы вошли как {user.full_name} ({user.role.value if user.role else '—'})."


def get_role_label(user: User) -> str:
    if user.role is None:
        return "—"
    return user.role.value
