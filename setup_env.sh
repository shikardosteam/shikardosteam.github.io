#!/bin/bash
# Одноразовый скрипт: добавляет JWT_SECRET (генерит сам) и GOOGLE_CLIENT_ID в .env бота, перезапускает сервис.
# Запуск на сервере:  curl -sL https://shikardos.team/setup_env.sh | bash
set -e
DIR=/opt/shikardos/shikardos-bot
cd "$DIR" || { echo "НЕ НАЙДЕНА папка $DIR"; exit 1; }

cp .env ".env.bak.$(date +%s)" 2>/dev/null || true
# убрать старые строки, если были
sed -i '/^JWT_SECRET=/d;/^GOOGLE_CLIENT_ID=/d' .env
# сгенерировать секрет прямо на сервере (ничего вводить не нужно)
echo "JWT_SECRET=$(openssl rand -hex 48)" >> .env
echo "GOOGLE_CLIENT_ID=654227330582-94gbaus2l20l9l9d3j7260s1h33g1blt.apps.googleusercontent.com" >> .env

echo "=== .env обновлён (значения скрыты) ==="
grep -E '^JWT_SECRET|^GOOGLE_CLIENT_ID' .env | sed 's/=.*/=***** установлено/'

echo "=== перезапуск сервисов ==="
systemctl restart shikardos || echo "ВНИМАНИЕ: сервис shikardos не перезапустился — проверь имя"
# если есть отдельный сервис API — перезапустим и его (если существует)
for s in shikardos-api shikardos-bot shikardosapi uvicorn; do
  systemctl list-units --type=service --all 2>/dev/null | grep -q "$s" && systemctl restart "$s" 2>/dev/null && echo "перезапущен: $s" || true
done

echo "=== статус ==="
systemctl --no-pager status shikardos | head -5
echo "=== все сервисы shikardos ==="
systemctl list-units --type=service | grep -i shikardos || echo "(других не найдено)"
echo "=== ГОТОВО. JWT_SECRET и GOOGLE_CLIENT_ID применены. ==="
