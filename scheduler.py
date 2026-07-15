# -*- coding: utf-8 -*-
"""Единый планировщик фоновых задач системы (2026-07, вынесен из bot.py).

Раньше APScheduler жил внутри bot.py вперемешку с бот-логикой. Вынесен в отдельный
модуль, чтобы:
  - планировщик не был «частью бота» — это задачи системного уровня;
  - задачи регистрировались в одном месте (morning_job бота, погода, будущие);
  - bot.py остался только про бота и вызывал start_scheduler() одной строкой.

ГДЕ ЗАПУСКАЕТСЯ: в бот-процессе (bot.py → main()). Бот — единственный always-on
процесс, поэтому планировщик крутится там (нет дублей, как было бы при запуске в
нескольких uvicorn-воркерах веба). Веб планировщик НЕ поднимает.

САМИ ЗАДАЧИ живут в своих модулях (morning_job — в bot.py, update_weather — в
production.py); здесь только регистрация и запуск. Timezone — МСК (Мурманская обл.).
"""
from datetime import timezone, timedelta
import logging

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger

from sqlalchemy.orm import Session

log = logging.getLogger("migbot.scheduler")

MSK = timezone(timedelta(hours=3))  # Мурманская обл. — МСК, без перехода на летнее время


def _weather_job(engine):
    """Обёртка задачи обновления погоды: открывает сессию и вызывает update_weather.
    Ошибки внутри update_weather не роняют планировщик (там try/except), но на всякий
    случай логируем и здесь."""
    import production
    try:
        with Session(engine) as session:
            n = production.update_weather(session)
            log.info("scheduler: погода обновлена, дат: %s", n)
    except Exception:
        log.exception("scheduler: ошибка задачи обновления погоды")


def create_scheduler(engine, morning_job=None) -> AsyncIOScheduler:
    """Создаёт планировщик и регистрирует все фоновые задачи.
    engine — SQLAlchemy engine (для задач, работающих с БД).
    morning_job — корутина утренней бот-рассылки (передаётся из bot.py, чтобы не
    тянуть бот-зависимости в этот модуль)."""
    scheduler = AsyncIOScheduler(timezone=MSK)

    # Утренняя бот-рассылка (9:00 МСК) — если передана из bot.py
    if morning_job is not None:
        scheduler.add_job(morning_job, CronTrigger(hour=9, minute=0), id="morning_job")

    # Погода из Open-Meteo (6:00 МСК, до утренней рассылки). Синхронная задача —
    # APScheduler выполнит её в пуле потоков, event loop не блокируется.
    scheduler.add_job(
        _weather_job, CronTrigger(hour=6, minute=0),
        args=[engine], id="weather_update",
    )

    return scheduler


def start_scheduler(engine, morning_job=None) -> AsyncIOScheduler:
    """Создаёт, регистрирует задачи и запускает планировщик. Вызывается из bot.main()."""
    scheduler = create_scheduler(engine, morning_job=morning_job)
    scheduler.start()
    log.info("scheduler: запущен, задач: %s", len(scheduler.get_jobs()))
    return scheduler
