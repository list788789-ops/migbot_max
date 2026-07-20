-- Миграция: бумажное согласие на ПДн для передачи в ООО ПСМ
-- (п.33.6 договора СМР № ПСМИПБ1600106, штраф по п.28.1.27 — 10 000 ₽ за каждое
-- непредоставленное согласие).
--
-- ПОРЯДОК ВЫКАТКИ: сначала ЭТОТ скрипт на боевой базе, только потом деплой кода.
-- Наоборот нельзя: SQLAlchemy запросит несуществующие колонки при первом же открытии
-- карточки сотрудника и весь раздел отдаст Internal Server Error.
--
-- Запуск на сервере:
--   psql "$DATABASE_URL" -f /root/2026-07-20_consent_paper_dates.sql
--
-- Обе колонки nullable, без DEFAULT и без backfill: пустая дата = «бумаги нет»,
-- это корректное стартовое состояние для всех 59 работников. Данные не трогаются,
-- откат безопасен (см. блок в конце).

BEGIN;

ALTER TABLE employees ADD COLUMN IF NOT EXISTS consent_signed_date DATE;
ALTER TABLE employees ADD COLUMN IF NOT EXISTS consent_transferred_date DATE;

-- Проверка ДО фиксации: обе колонки должны присутствовать и быть nullable.
SELECT column_name, data_type, is_nullable
FROM information_schema.columns
WHERE table_name = 'employees'
  AND column_name IN ('consent_signed_date', 'consent_transferred_date')
ORDER BY column_name;

COMMIT;

-- Ожидаемый вывод SELECT — ровно две строки:
--   consent_signed_date       | date | YES
--   consent_transferred_date  | date | YES
-- Если строк меньше двух — COMMIT прошёл вхолостую, код деплоить НЕЛЬЗЯ, разбирайся.

-- Откат (выполнять только вручную и осознанно — сотрёт проставленные даты):
-- BEGIN;
-- ALTER TABLE employees DROP COLUMN IF EXISTS consent_signed_date;
-- ALTER TABLE employees DROP COLUMN IF EXISTS consent_transferred_date;
-- COMMIT;
