#!/usr/bin/env bash
# Deploy to Cloudflare Pages while detecting Wrangler's occasional log hang.
# Usage: ./scripts/deploy-pages.sh <public_dir> <project_name> <branch>

set -euo pipefail

PUBLIC_DIR="${1:?Usage: $0 <public_dir> <project_name> <branch>}"
PROJECT_NAME="${2:?Usage: $0 <public_dir> <project_name> <branch>}"
BRANCH="${3:?Usage: $0 <public_dir> <project_name> <branch>}"

if ! command -v wrangler &>/dev/null; then
  echo "ERROR: wrangler not found. Run bun install first." >&2
  exit 1
fi

workspace_root=$(cd "$(dirname "${BASH_SOURCE[0]}")/../../../.." && pwd)
mkdir -p "$workspace_root/tmp"
log_file=$(mktemp "$workspace_root/tmp/sheets-pages-deploy.XXXXXX.log")
wrangler_pid=""

cleanup() {
  if [ -n "$wrangler_pid" ] && kill -0 "$wrangler_pid" 2>/dev/null; then
    kill "$wrangler_pid" 2>/dev/null || true
    wait "$wrangler_pid" 2>/dev/null || true
  fi
  rm -f "$log_file"
}
trap cleanup EXIT

echo "Deploying $PUBLIC_DIR to Cloudflare Pages project $PROJECT_NAME..."
env -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY -u http_proxy -u https_proxy -u all_proxy \
  CI=true CLOUDFLARE_TELEMETRY_DISABLED=1 NO_UPDATE_NOTIFIER=1 \
  wrangler pages deploy "$PUBLIC_DIR" --project-name "$PROJECT_NAME" --branch "$BRANCH" --commit-dirty=true \
  >"$log_file" 2>&1 &
wrangler_pid=$!

success=false
last_line_count=0

for _ in {1..300}; do
  if [ -f "$log_file" ]; then
    current_lines=$(wc -l <"$log_file")
    if [ "$current_lines" -gt "$last_line_count" ]; then
      tail -n +"$((last_line_count + 1))" "$log_file"
      last_line_count=$current_lines
    fi
  fi

  if grep -q "Deployment complete!" "$log_file" 2>/dev/null; then
    success=true
    break
  fi

  if ! kill -0 "$wrangler_pid" 2>/dev/null; then
    grep -q "Deployment complete!" "$log_file" 2>/dev/null && success=true
    break
  fi
  sleep 0.5
done

if [ -f "$log_file" ]; then
  current_lines=$(wc -l <"$log_file")
  if [ "$current_lines" -gt "$last_line_count" ]; then
    tail -n +"$((last_line_count + 1))" "$log_file"
  fi
fi

if [ "$success" = true ]; then
  echo "Cloudflare Pages deployment finished successfully."
else
  echo "Cloudflare Pages deployment failed or timed out." >&2
  exit 1
fi
