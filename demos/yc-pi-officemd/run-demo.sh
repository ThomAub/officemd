#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
REPO_ROOT=$(cd -- "$SCRIPT_DIR/../.." && pwd)
INPUTS_DIR="$SCRIPT_DIR/fixtures/inputs"
SKILL_DIR="$REPO_ROOT/.pi/skills/officemd"
MODE="both"
PROVIDER="${PI_PROVIDER:-openai-codex}"
MODEL="${PI_MODEL:-gpt-5.5}"
THINKING="${PI_THINKING:-high}"
DRY_RUN=0
RUN_ID=$(date -u +%Y%m%dT%H%M%SZ)
ANY_FAILURE=0

usage() {
  printf '%s\n' \
    'Usage: ./run-demo.sh [--mode baseline|skill|both] [--provider NAME] [--model ID]' \
    '                     [--thinking LEVEL] [--run-id ID] [--dry-run]'
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --mode) MODE=$2; shift 2 ;;
    --provider) PROVIDER=$2; shift 2 ;;
    --model) MODEL=$2; shift 2 ;;
    --thinking) THINKING=$2; shift 2 ;;
    --run-id) RUN_ID=$2; shift 2 ;;
    --dry-run) DRY_RUN=1; shift ;;
    --help|-h) usage; exit 0 ;;
    *) printf 'unknown argument: %s\n' "$1" >&2; usage >&2; exit 2 ;;
  esac
done

if [[ "$MODE" != "baseline" && "$MODE" != "skill" && "$MODE" != "both" ]]; then
  printf 'invalid mode: %s\n' "$MODE" >&2
  exit 2
fi

command -v pi >/dev/null || { printf 'pi is not installed\n' >&2; exit 1; }
"$SCRIPT_DIR/setup.sh"

RUN_ROOT="$SCRIPT_DIR/runs/$RUN_ID"
mkdir -p "$RUN_ROOT"

run_variant() {
  local variant=$1
  local variant_dir="$RUN_ROOT/$variant"
  local work_dir="$variant_dir/work"
  local agent_work_dir
  rm -rf "$variant_dir"
  mkdir -p "$variant_dir"
  if [[ "$DRY_RUN" -eq 1 ]]; then
    agent_work_dir="$work_dir"
    mkdir -p "$agent_work_dir"
  else
    agent_work_dir=$(mktemp -d "/tmp/officemd-yc-${RUN_ID}-${variant}.XXXXXX")
  fi
  cp "$INPUTS_DIR"/*.xlsx "$agent_work_dir/"

  local -a command=(
    pi
    --provider "$PROVIDER"
    --model "$MODEL"
    --thinking "$THINKING"
    --mode json
    --print
    --no-session
    --no-skills
    --no-extensions
    --no-context-files
    --no-prompt-templates
  )
  if [[ "$variant" == "skill" ]]; then
    command+=(--skill "$SKILL_DIR")
  fi
  command+=("$(<"$SCRIPT_DIR/PROMPT.md")")

  printf '%s\n' "variant=$variant" "provider=$PROVIDER" "model=$MODEL" "thinking=$THINKING" > "$variant_dir/config.txt"
  printf '%q ' "${command[@]}" > "$variant_dir/command.txt"
  printf '\n' >> "$variant_dir/command.txt"

  if [[ "$DRY_RUN" -eq 1 ]]; then
    printf '[dry-run] %s: ' "$variant"
    printf '%q ' "${command[@]}"
    printf '\n'
    return
  fi

  printf 'running %s with %s/%s\n' "$variant" "$PROVIDER" "$MODEL"
  local started ended exit_code
  started=$(date -u +%Y-%m-%dT%H:%M:%SZ)
  set +e
  (
    cd "$agent_work_dir"
    PATH="$REPO_ROOT/target/debug:$PATH" "${command[@]}"
  ) > "$variant_dir/events.jsonl" 2> "$variant_dir/stderr.log"
  exit_code=$?
  set -e
  ended=$(date -u +%Y-%m-%dT%H:%M:%SZ)
  if [[ "$exit_code" -ne 0 ]]; then
    ANY_FAILURE=1
  fi
  mkdir -p "$work_dir"
  cp -R "$agent_work_dir/." "$work_dir/"
  rm -rf "$agent_work_dir"

  jq -n \
    --arg variant "$variant" \
    --arg provider "$PROVIDER" \
    --arg model "$MODEL" \
    --arg thinking "$THINKING" \
    --arg started "$started" \
    --arg ended "$ended" \
    --argjson exit_code "$exit_code" \
    '{variant:$variant,provider:$provider,model:$model,thinking:$thinking,started:$started,ended:$ended,exit_code:$exit_code}' \
    > "$variant_dir/run.json"

  PATH="$REPO_ROOT/target/debug:$PATH" uv run "$SCRIPT_DIR/score.py" \
    --officemd "$REPO_ROOT/target/debug/officemd" \
    --work-dir "$work_dir" \
    --fixtures-dir "$INPUTS_DIR" \
    --evidence-dir "$variant_dir/evidence" \
    --output "$variant_dir/score.json"
  printf '%s score: %s/100\n' "$variant" "$(jq -r '.score' "$variant_dir/score.json")"
}

if [[ "$MODE" == "baseline" || "$MODE" == "both" ]]; then
  run_variant baseline
fi
if [[ "$MODE" == "skill" || "$MODE" == "both" ]]; then
  run_variant skill
fi

if [[ "$DRY_RUN" -eq 0 ]]; then
  uv run "$SCRIPT_DIR/report.py" --run-dir "$RUN_ROOT"
  printf 'report: %s\n' "$RUN_ROOT/report.html"
else
  printf 'dry run prepared at %s\n' "$RUN_ROOT"
fi

if [[ "$ANY_FAILURE" -ne 0 ]]; then
  printf 'one or more Pi variants failed; inspect stderr.log under %s\n' "$RUN_ROOT" >&2
  exit 1
fi
