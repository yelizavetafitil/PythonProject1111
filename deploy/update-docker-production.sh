#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ENV_FILE="${1:-$PROJECT_DIR/.env.production}"
BACKUP_DIR="${PORTAL_BACKUP_DIR:-$PROJECT_DIR/backups}"

cd "$PROJECT_DIR"

if ! docker compose version >/dev/null 2>&1; then
  echo "Docker Compose v2 не найден" >&2
  exit 1
fi
if [[ ! -f "$ENV_FILE" ]]; then
  echo "Не найден production-конфиг: $ENV_FILE" >&2
  exit 1
fi

read_value() {
  awk -F= -v key="$2" '$1 == key {sub(/^[^=]*=/, ""); print; exit}' "$1"
}

require_secret() {
  local key="$1"
  local value
  value="$(read_value "$ENV_FILE" "$key")"
  if [[ ${#value} -lt 32 || "$value" == CHANGE_ME* ]]; then
    echo "В $ENV_FILE не задан безопасный $key" >&2
    exit 1
  fi
}

require_value() {
  local key="$1"
  local value
  value="$(read_value "$ENV_FILE" "$key")"
  if [[ -z "$value" || "$value" == CHANGE_ME* ]]; then
    echo "В $ENV_FILE не задан $key" >&2
    exit 1
  fi
}

require_secret PORTAL_SECRET_KEY
require_secret AI_SSO_SHARED_SECRET
require_value LDAP_BIND_DN
require_value LDAP_BIND_PASSWORD
require_value AI_ASSISTANT_PUBLIC_URL
require_value AI_ASSISTANT_CALLBACK_URL
chmod 600 "$ENV_FILE"

compose=(docker compose --env-file "$ENV_FILE")
"${compose[@]}" config --quiet

echo "Собираю новый образ портала..."
"${compose[@]}" build web

install -d -m 700 "$BACKUP_DIR"
timestamp="$(date +%Y%m%d-%H%M%S)"

# SQLite копируется только после остановки web-процесса: так бэкап целостный.
"${compose[@]}" stop web
if [[ -f database.db ]]; then
  backup_path="$BACKUP_DIR/database_before_sso_${timestamp}.db"
  cp --preserve=timestamps database.db "$backup_path"
  chmod 600 "$backup_path"
  echo "Резервная копия базы: $backup_path"
fi

echo "Запускаю обновлённый контейнер..."
"${compose[@]}" up -d --remove-orphans web

container_id="$("${compose[@]}" ps -q web)"
for _ in $(seq 1 30); do
  status="$(docker inspect --format '{{if .State.Health}}{{.State.Health.Status}}{{else}}{{.State.Status}}{{end}}' "$container_id" 2>/dev/null || true)"
  if [[ "$status" == healthy ]]; then
    echo "Портал обновлён, контейнер healthy."
    "${compose[@]}" ps
    exit 0
  fi
  if [[ "$status" == unhealthy || "$status" == exited || "$status" == dead ]]; then
    break
  fi
  sleep 2
done

echo "Портал не прошёл healthcheck. Последние логи:" >&2
"${compose[@]}" logs --tail=100 web >&2
exit 1
