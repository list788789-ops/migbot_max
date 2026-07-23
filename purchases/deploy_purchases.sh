#!/usr/bin/env bash
# deploy_purchases.sh — выкатка раздела «Покупки».
#
# Запуск:
#   cd ~/migbot_max
#   bash deploy_purchases.sh
#
# Скрипт НЕ трогает базу до перезапуска приложения: индексы вешаются только после
# того, как create_all() создаст таблицы. Порядок здесь принципиален, поэтому шаги
# идут строго последовательно и падают при первой ошибке.

set -euo pipefail

APP_DIR="${APP_DIR:-$HOME/migbot_max}"
ENV_FILE="$APP_DIR/.env"

say()  { printf '\n\033[1m== %s\033[0m\n' "$*"; }
ok()   { printf '   \033[32m✓\033[0m %s\n' "$*"; }
warn() { printf '   \033[33m!\033[0m %s\n' "$*"; }
die()  { printf '\n\033[31m✗ %s\033[0m\n' "$*" >&2; exit 1; }

cd "$APP_DIR" || die "Нет папки $APP_DIR"

# --- 1. Код ------------------------------------------------------------------
say "1/6  Забираю код из GitHub"
git pull --ff-only || die "git pull не прошёл. Разберись с конфликтами и запусти заново."
ok "код обновлён"

for f in models.py webforms.py purchases/parser.py purchases/matching.py \
         purchases/service.py purchases/routes.py purchases/__init__.py; do
    [ -f "$f" ] || die "Не хватает файла: $f — залей его в репозиторий и повтори."
done
ok "все файлы раздела на месте"

# --- 2. Зависимости ----------------------------------------------------------
say "2/6  Зависимости"
if ! grep -qi '^openpyxl' requirements.txt 2>/dev/null; then
    echo 'openpyxl>=3.1' >> requirements.txt
    warn "openpyxl дописан в requirements.txt — не забудь закоммитить"
fi
if ! grep -qi '^rapidfuzz' requirements.txt 2>/dev/null; then
    echo 'rapidfuzz>=3.0' >> requirements.txt
    warn "rapidfuzz дописан в requirements.txt — не забудь закоммитить"
fi

PIP="pip"
[ -x "$APP_DIR/venv/bin/pip" ] && PIP="$APP_DIR/venv/bin/pip"
[ -x "$APP_DIR/.venv/bin/pip" ] && PIP="$APP_DIR/.venv/bin/pip"
$PIP install -q -r requirements.txt || die "pip install упал"
ok "зависимости поставлены ($PIP)"

# --- 3. Проверка импортов ДО рестарта ----------------------------------------
# Смысл: поймать опечатку здесь, а не уронить работающее приложение.
say "3/6  Проверяю, что модули импортируются"
PY="python3"
[ -x "$APP_DIR/venv/bin/python" ] && PY="$APP_DIR/venv/bin/python"
[ -x "$APP_DIR/.venv/bin/python" ] && PY="$APP_DIR/.venv/bin/python"

set -a; [ -f "$ENV_FILE" ] && . "$ENV_FILE"; set +a

$PY - <<'EOF' || die "Импорт не прошёл — приложение НЕ перезапускалось, работает старая версия."
import sys
try:
    import models
    from purchases import parser, matching, service, routes
    assert hasattr(models, "GoodsReceipt"), "в models.py нет GoodsReceipt"
    assert hasattr(models.UserRole, "BUHGALTER"), "в UserRole нет BUHGALTER"
    print("   ✓ models и purchases импортируются")
except Exception as e:
    print(f"   ошибка: {type(e).__name__}: {e}", file=sys.stderr)
    sys.exit(1)
EOF

# --- 4. Рестарт --------------------------------------------------------------
say "4/6  Перезапуск приложения"
SERVICE="${SERVICE:-}"
if [ -z "$SERVICE" ]; then
    SERVICE=$(systemctl list-units --type=service --no-legend 2>/dev/null \
              | awk '{print $1}' | grep -Ei 'migbot|webform|tact' | head -1 || true)
fi

if [ -n "$SERVICE" ]; then
    systemctl restart "$SERVICE" || die "Не удалось перезапустить $SERVICE"
    sleep 4
    systemctl is-active --quiet "$SERVICE" \
        || die "Сервис $SERVICE не поднялся. Логи: journalctl -u $SERVICE -n 50 --no-pager"
    ok "перезапущен: $SERVICE"
else
    warn "Сервис systemd не найден автоматически."
    warn "Перезапусти приложение вручную, потом запусти скрипт снова так:"
    warn "  SKIP_RESTART=1 bash deploy_purchases.sh"
    [ "${SKIP_RESTART:-0}" = "1" ] || exit 1
fi

# --- 5. Проверка таблиц ------------------------------------------------------
say "5/6  Жду появления таблиц"
[ -n "${DATABASE_URL:-}" ] || die "DATABASE_URL не задан (нет $ENV_FILE?)"
command -v psql >/dev/null || die "psql не установлен: apt install postgresql-client -y"

FOUND=0
for i in 1 2 3 4 5 6; do
    CNT=$(psql "$DATABASE_URL" -tAc "
        SELECT count(*) FROM information_schema.tables
        WHERE table_name IN ('suppliers','nomenclature','supplier_nomenclature_map',
                             'supplier_parse_profiles','goods_receipts','goods_receipt_lines');
    " 2>/dev/null || echo 0)
    if [ "$CNT" = "6" ]; then FOUND=1; break; fi
    printf '   ждём... (%s/6 таблиц)\n' "$CNT"
    sleep 5
done
[ "$FOUND" = "1" ] || die "Таблицы не создались. Приложение стартовало? Смотри логи."
ok "все 6 таблиц созданы"

# --- 6. Индексы --------------------------------------------------------------
say "6/6  Индексы"
psql "$DATABASE_URL" -v ON_ERROR_STOP=1 <<'SQL'
CREATE INDEX IF NOT EXISTS ix_snm_lookup
    ON supplier_nomenclature_map (supplier_id, supplier_key);
CREATE INDEX IF NOT EXISTS ix_receipt_supplier_date
    ON goods_receipts (supplier_id, doc_date DESC);
CREATE INDEX IF NOT EXISTS ix_receipt_status
    ON goods_receipts (status);
CREATE INDEX IF NOT EXISTS ix_lines_receipt
    ON goods_receipt_lines (receipt_id, line_no);
CREATE INDEX IF NOT EXISTS ix_lines_nomenclature
    ON goods_receipt_lines (nomenclature_id);
SQL
ok "индексы созданы"

# --- Итог --------------------------------------------------------------------
say "Готово"
cat <<'TXT'
   Дальше руками:
   1) Зайти в приложение под админом → Пользователи → назначить кому-то
      роль «Бухгалтер» (или себе, для проверки).
   2) В меню появится пункт «Покупки». Он виден только бухгалтеру и админу.
   3) Завести поставщика, загрузить тестовую накладную .xlsx
   4) Загрузить ТОТ ЖЕ файл повторно — должно открыться существующее
      поступление с баннером, а не создаться второе.
   5) Подтвердить пару строк вручную, загрузить файл заново — подтверждённые
      должны подтянуться сами. Это проверка регистра соответствий, главный пункт.
TXT
