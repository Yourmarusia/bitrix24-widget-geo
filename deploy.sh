#!/bin/bash
# Запускать на сервере из папки /var/www/geo-widget (или где лежит репо)
set -e

echo "=== 1. Проверяем сеть Caddy ==="
CADDY_NET=$(docker network ls --format '{{.Name}}' | grep -i caddy | head -1)
if [ -z "$CADDY_NET" ]; then
  echo "Сеть caddy не найдена, создаём caddy-net..."
  docker network create caddy-net
  CADDY_NET="caddy-net"
fi
echo "Используем сеть: $CADDY_NET"

# Если сеть называется не caddy-net — подставляем в compose
sed -i "s/caddy-net/$CADDY_NET/g" docker-compose.yml

echo "=== 2. Собираем и запускаем контейнер ==="
docker compose up -d --build

echo "=== 3. Проверяем контейнер ==="
sleep 2
docker ps | grep geo-widget

echo "=== 4. Тест внутри контейнера ==="
docker exec geo-widget wget -qO- http://localhost/widget/ | head -5

echo ""
echo "=== Готово! ==="
echo "Осталось обновить Caddyfile:"
echo ""
echo "  handle /geo-widget/* {"
echo "      uri strip_prefix /geo-widget"
echo "      reverse_proxy geo-widget:80"
echo "  }"
echo ""
echo "Затем: sudo systemctl reload caddy"
