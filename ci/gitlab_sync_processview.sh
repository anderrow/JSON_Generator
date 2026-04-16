#!/usr/bin/env bash
set -euo pipefail

require_var() {
  local name="$1"
  if [[ -z "${!name:-}" ]]; then
    echo "Missing required variable: ${name}" >&2
    exit 1
  fi
}

build_auth_url() {
  local base_url="$1"
  local user="$2"
  local token="$3"
  python - "$base_url" "$user" "$token" <<'PY'
import sys
from urllib.parse import urlsplit, urlunsplit

base_url, user, token = sys.argv[1:]
parts = urlsplit(base_url)
netloc = f"{user}:{token}@{parts.netloc}"
print(urlunsplit((parts.scheme, netloc, parts.path, parts.query, parts.fragment)))
PY
}

clone_source_repo() {
  local repo_dir="$1"

  if [[ -z "${SOURCE_REPO_URL:-}" ]]; then
    if [[ -d "$CI_PROJECT_DIR/SOURCE FILES" ]]; then
      printf '%s\n' "$CI_PROJECT_DIR/SOURCE FILES"
      return
    fi

    echo "SOURCE_REPO_URL is not set and this repository does not contain 'SOURCE FILES/'." >&2
    echo "Put the CSV files in this repo or configure SOURCE_REPO_URL." >&2
    exit 1
  fi

  local source_ref="${SOURCE_REPO_REF:-main}"
  local source_subdir="${SOURCE_REPO_SUBDIR:-SOURCE FILES}"
  local clone_url="$SOURCE_REPO_URL"

  if [[ -n "${SOURCE_REPO_TOKEN:-}" ]]; then
    clone_url="$(build_auth_url "$SOURCE_REPO_URL" "oauth2" "$SOURCE_REPO_TOKEN")"
  elif [[ -n "${CI_JOB_TOKEN:-}" ]]; then
    clone_url="$(build_auth_url "$SOURCE_REPO_URL" "gitlab-ci-token" "$CI_JOB_TOKEN")"
  fi

  git clone --depth 1 --branch "$source_ref" "$clone_url" "$repo_dir"
  printf '%s\n' "$repo_dir/$source_subdir"
}

require_var PROCESSVIEW_WRITE_TOKEN

processview_clone="$CI_PROJECT_DIR/processview"
source_repo_clone="$CI_PROJECT_DIR/source-repo"
generated_json_dir="${GENERATED_JSON_DIR:-$CI_PROJECT_DIR/generated-json}"

source_dir="$(clone_source_repo "$source_repo_clone")"

if [[ ! -d "$source_dir" ]]; then
  echo "Source directory not found: $source_dir" >&2
  exit 1
fi

processview_url="$(build_auth_url "$PROCESSVIEW_REPO_URL" "oauth2" "$PROCESSVIEW_WRITE_TOKEN")"
git clone --depth 1 --branch "${PROCESSVIEW_TARGET_BRANCH:-develop_ibo}" "$processview_url" "$processview_clone"

git -C "$processview_clone" config user.name "${GIT_BOT_NAME:-ProcessView CI Bot}"
git -C "$processview_clone" config user.email "${GIT_BOT_EMAIL:-processview-ci@example.com}"

python sync_processview.py \
  --source-dir "$source_dir" \
  --json-dir "$generated_json_dir" \
  --processview-repo "$processview_clone" \
  --skip-pull \
  --push \
  --commit-message "${PROCESSVIEW_COMMIT_MESSAGE:-Update ProcessView translations from ${CI_PROJECT_PATH:-manual}@${CI_COMMIT_SHORT_SHA:-manual}}"
