#!/usr/bin/env bash

set -uo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

# Use a repo-local cargo home by default to avoid lock contention on ~/.cargo
# when many agents build concurrently. Preserve any user/CI override.
#
# Note: some agent runners pre-set `CARGO_HOME=$HOME/.cargo`. Treat that value as
# "unset" for our purposes so we still get per-repo isolation by default.
# In CI we respect `CARGO_HOME` even if it points at `$HOME/.cargo` so CI can use
# shared caching.
# To explicitly keep `CARGO_HOME=$HOME/.cargo` in local runs, set
# `FORMULA_ALLOW_GLOBAL_CARGO_HOME=1`.
DEFAULT_GLOBAL_CARGO_HOME="${HOME:-/root}/.cargo"
CARGO_HOME_NORM="${CARGO_HOME:-}"
CARGO_HOME_NORM="${CARGO_HOME_NORM%/}"
DEFAULT_GLOBAL_CARGO_HOME_NORM="${DEFAULT_GLOBAL_CARGO_HOME%/}"
if [ -z "${CARGO_HOME:-}" ] || {
  [ -z "${CI:-}" ] &&
    [ -z "${FORMULA_ALLOW_GLOBAL_CARGO_HOME:-}" ] &&
    [ "${CARGO_HOME_NORM}" = "${DEFAULT_GLOBAL_CARGO_HOME_NORM}" ];
}; then
  export CARGO_HOME="$ROOT_DIR/target/cargo-home"
fi
mkdir -p "$CARGO_HOME"
mkdir -p "$CARGO_HOME/bin"
case ":$PATH:" in
  *":$CARGO_HOME/bin:"*) ;;
  *) export PATH="$CARGO_HOME/bin:$PATH" ;;
esac

# Some environments configure Cargo globally with `build.rustc-wrapper` (often `sccache`).
# When the wrapper is unavailable/misconfigured, builds can fail even for `cargo metadata`.
# Default to disabling any configured wrapper unless the user explicitly overrides it in the env.
export RUSTC_WRAPPER="${RUSTC_WRAPPER:-}"
export RUSTC_WORKSPACE_WRAPPER="${RUSTC_WORKSPACE_WRAPPER:-}"

REPORT_DIR="${REPORT_DIR:-security-report}"
ALLOWLIST_CARGO="security/allowlist/cargo-audit.txt"
ALLOWLIST_NODE="security/allowlist/node-audit.txt"

rm -rf "$REPORT_DIR"
mkdir -p "$REPORT_DIR"

SUMMARY="${REPORT_DIR}/summary.md"

commit_sha="${GITHUB_SHA:-}"
if [ -z "$commit_sha" ] && command -v git >/dev/null 2>&1; then
  commit_sha="$(git rev-parse HEAD 2>/dev/null || true)"
fi

{
  echo "# Security report"
  echo
  echo "- Generated: $(date -u +"%Y-%m-%dT%H:%M:%SZ")"
  if [ -n "$commit_sha" ]; then
    echo "- Commit: ${commit_sha}"
  fi
} >"$SUMMARY"

fail=0

section() {
  echo >>"$SUMMARY"
  echo "## $1" >>"$SUMMARY"
}

note() {
  echo "- $1" >>"$SUMMARY"
}

run_gitleaks() {
  section "Secret scanning (gitleaks)"

  if ! command -v gitleaks >/dev/null 2>&1; then
    note "Skipped: gitleaks not installed"
    return 0
  fi

  local log_file="${REPORT_DIR}/gitleaks.log"
  local report_file="${REPORT_DIR}/gitleaks.json"

  set +e
  gitleaks detect \
    --source . \
    --no-git \
    --redact \
    --report-format json \
    --report-path "$report_file" \
    --no-banner >"$log_file" 2>&1
  local rc=$?
  set -e

  if [ $rc -eq 0 ]; then
    note "✅ No secrets detected"
    return 0
  fi

  note "❌ Potential secrets detected (see ${report_file})"
  return 1
}

find_files() {
  local name="$1"
  find . -type f -name "$name" \
    -not -path "./.git/*" \
    -not -path "*/node_modules/*" \
    -not -path "*/target/*" \
    -not -path "./${REPORT_DIR}/*" \
    2>/dev/null
}

slugify_path() {
  local p="$1"
  p="${p#./}"
  p="${p//\//_}"
  if [ -z "$p" ] || [ "$p" = "." ]; then
    echo "root"
  else
    echo "$p"
  fi
}

run_cargo_audit() {
  section "Dependency vulnerabilities (Rust: cargo audit)"

  if ! command -v cargo >/dev/null 2>&1 || ! cargo audit --version >/dev/null 2>&1; then
    note "Skipped: cargo-audit not installed"
    return 0
  fi

  mapfile -t locks < <(find_files "Cargo.lock")
  if [ ${#locks[@]} -eq 0 ]; then
    note "Skipped: no Cargo.lock found"
    return 0
  fi

  for lock in "${locks[@]}"; do
    local dir
    dir="$(dirname "$lock")"
    local slug
    slug="$(slugify_path "$dir")"

    local out_json="${REPORT_DIR}/cargo-audit_${slug}.json"
    local out_err="${REPORT_DIR}/cargo-audit_${slug}.stderr"
    local eval_json="${REPORT_DIR}/cargo-audit_${slug}.policy.json"
    local eval_txt="${REPORT_DIR}/cargo-audit_${slug}.policy.txt"

    set +e
    (cd "$dir" && cargo audit --json) >"$out_json" 2>"$out_err"
    local rc=$?
    set -e

    # cargo audit exits non-zero for any vulnerability; policy evaluation controls CI failure.
    set +e
    python3 scripts/security/evaluate_cargo_audit.py \
      --input "$out_json" \
      --allowlist "$ALLOWLIST_CARGO" \
      --output "$eval_json" >"$eval_txt" 2>&1
    local policy_rc=$?
    set -e

    local label="✅"
    if [ $policy_rc -ne 0 ]; then
      label="❌"
      fail=1
    fi

    note "${label} ${dir}: $(head -n 1 "$eval_txt") (details: cargo-audit_${slug}.policy.json)"

    # Preserve the raw cargo-audit exit code for debugging.
    echo "$rc" >"${REPORT_DIR}/cargo-audit_${slug}.exitcode"
  done
}

detect_node_root() {
  # Prefer repo root; otherwise pick the first directory containing a lockfile and package.json.
  if [ -f "pnpm-lock.yaml" ] && [ -f "package.json" ]; then
    echo "."
    return 0
  fi
  if [ -f "package-lock.json" ] && [ -f "package.json" ]; then
    echo "."
    return 0
  fi

  local lock
  lock="$(find_files "pnpm-lock.yaml" | head -n 1 || true)"
  if [ -n "$lock" ] && [ -f "$(dirname "$lock")/package.json" ]; then
    echo "$(dirname "$lock")"
    return 0
  fi

  lock="$(find_files "package-lock.json" | head -n 1 || true)"
  if [ -n "$lock" ] && [ -f "$(dirname "$lock")/package.json" ]; then
    echo "$(dirname "$lock")"
    return 0
  fi

  echo ""
}

run_node_audit() {
  section "Dependency vulnerabilities (Node: pnpm/npm audit)"

  local node_root
  node_root="$(detect_node_root)"
  if [ -z "$node_root" ]; then
    note "Skipped: no Node lockfile + package.json found"
    return 0
  fi

  local slug
  slug="$(slugify_path "$node_root")"

  local out_json="${REPORT_DIR}/node-audit_${slug}.json"
  local out_err="${REPORT_DIR}/node-audit_${slug}.stderr"
  local eval_json="${REPORT_DIR}/node-audit_${slug}.policy.json"
  local eval_txt="${REPORT_DIR}/node-audit_${slug}.policy.txt"

  set +e
  if [ -f "${node_root}/pnpm-lock.yaml" ] && command -v pnpm >/dev/null 2>&1; then
    (cd "$node_root" && pnpm audit --json) >"$out_json" 2>"$out_err"
  else
    (cd "$node_root" && npm audit --json) >"$out_json" 2>"$out_err"
  fi
  local rc=$?
  set -e

  set +e
  python3 scripts/security/evaluate_node_audit.py \
    --input "$out_json" \
    --allowlist "$ALLOWLIST_NODE" \
    --output "$eval_json" >"$eval_txt" 2>&1
  local policy_rc=$?
  set -e

  local label="✅"
  if [ $policy_rc -ne 0 ]; then
    label="❌"
    fail=1
  fi

  note "${label} ${node_root}: $(head -n 1 "$eval_txt") (details: node-audit_${slug}.policy.json)"

  echo "$rc" >"${REPORT_DIR}/node-audit_${slug}.exitcode"
}

run_clippy() {
  section "SAST (Rust: cargo clippy)"

  if ! command -v cargo >/dev/null 2>&1; then
    note "Skipped: cargo not installed"
    return 0
  fi

  mapfile -t manifests < <(find_files "Cargo.toml")
  if [ ${#manifests[@]} -eq 0 ]; then
    note "Skipped: no Cargo.toml found"
    return 0
  fi

  # Prefer the workspace root (Cargo.lock directory) when available.
  mapfile -t locks < <(find_files "Cargo.lock")
  if [ ${#locks[@]} -gt 0 ]; then
    local dir
    dir="$(dirname "${locks[0]}")"
    local slug
    slug="$(slugify_path "$dir")"
    local out="${REPORT_DIR}/clippy_${slug}.log"

    set +e
    (
      cd "$dir"
      bash "${ROOT_DIR}/scripts/cargo_agent.sh" clippy --workspace --all-targets --all-features -- \
        -D clippy::unwrap_used \
        -D clippy::expect_used \
        -D clippy::panic \
        -D clippy::panic_in_result_fn \
        -D clippy::todo \
        -D clippy::unimplemented
    ) >"$out" 2>&1
    local rc=$?
    set -e

    if [ $rc -eq 0 ]; then
      note "✅ ${dir}: clippy passed"
    else
      note "❌ ${dir}: clippy failed (see ${out})"
      fail=1
    fi
    return 0
  fi

  # Fallback: run clippy in the directory of each Cargo.toml.
  for manifest in "${manifests[@]}"; do
    local dir
    dir="$(dirname "$manifest")"
    local slug
    slug="$(slugify_path "$dir")"
    local out="${REPORT_DIR}/clippy_${slug}.log"

    set +e
    (
      cd "$dir"
      bash "${ROOT_DIR}/scripts/cargo_agent.sh" clippy --all-targets --all-features -- \
        -D clippy::unwrap_used \
        -D clippy::expect_used \
        -D clippy::panic \
        -D clippy::panic_in_result_fn \
        -D clippy::todo \
        -D clippy::unimplemented
    ) >"$out" 2>&1
    local rc=$?
    set -e

    if [ $rc -eq 0 ]; then
      note "✅ ${dir}: clippy passed"
    else
      note "❌ ${dir}: clippy failed (see ${out})"
      fail=1
    fi
  done
}

run_node_sast() {
  section "SAST (JS/TS: eslint + TypeScript strict mode)"

  local node_root
  node_root="$(detect_node_root)"
  if [ -z "$node_root" ]; then
    note "Skipped: no Node project detected"
    return 0
  fi

  local slug
  slug="$(slugify_path "$node_root")"
  local out="${REPORT_DIR}/node-sast_${slug}.log"
  local rc=0

  set +e
  if [ -f "${node_root}/pnpm-lock.yaml" ] && command -v pnpm >/dev/null 2>&1; then
    (
      cd "$node_root"
      pnpm install --frozen-lockfile
      pnpm -r --if-present lint
      pnpm -r --if-present typecheck
    ) >"$out" 2>&1
    rc=$?
  else
    (
      cd "$node_root"
      npm ci
      npm run --if-present lint
      npm run --if-present typecheck
    ) >"$out" 2>&1
    rc=$?
  fi
  set -e

  if [ $rc -eq 0 ]; then
    note "✅ ${node_root}: lint/typecheck passed"
  else
    note "❌ ${node_root}: lint/typecheck failed (see ${out})"
    fail=1
  fi

  # Lightweight policy checks for security-focused configuration.
  local policy_out="${REPORT_DIR}/node-sast-policy_${slug}.txt"
  set +e
  python3 - >"$policy_out" 2>&1 <<'PY'
import json
import os
import sys
from pathlib import Path

root = Path(".")

def find_files(name: str):
    for p in root.rglob(name):
        if ".git" in p.parts or "node_modules" in p.parts or "target" in p.parts:
            continue
        yield p

def read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None

issues = []

# TypeScript strict mode: if any TS/TSX exists, require *some* tsconfig.json to enable strict mode.
has_ts = any(
    p.suffix in {".ts", ".tsx"} and "node_modules" not in p.parts and ".git" not in p.parts
    for p in root.rglob("*")
)
if has_ts:
    tsconfigs = list(find_files("tsconfig.json"))
    strict_enabled = False
    for cfg in tsconfigs:
        text = cfg.read_text(encoding="utf-8", errors="replace")
        normalized = "".join(text.split())
        if '"strict":true' in normalized:
            strict_enabled = True
            break
    if not strict_enabled:
        issues.append("TypeScript is present but no tsconfig.json enables compilerOptions.strict = true")

# ESLint security plugin expectation: if eslint is present in dependencies, require eslint-plugin-security.
eslint_projects = 0
eslint_security = 0
for pkg_json in find_files("package.json"):
    pkg = read_json(pkg_json)
    if not isinstance(pkg, dict):
        continue
    deps = {}
    for k in ("dependencies", "devDependencies", "optionalDependencies"):
        if isinstance(pkg.get(k), dict):
            deps.update(pkg[k])
    if "eslint" in deps:
        eslint_projects += 1
        if "eslint-plugin-security" in deps:
            eslint_security += 1

if eslint_projects > 0 and eslint_security == 0:
    issues.append("eslint is present but eslint-plugin-security was not found in dependencies/devDependencies")

if issues:
    print("❌ Node SAST configuration policy failed:")
    for i in issues:
        print(f"  - {i}")
    raise SystemExit(1)

print("✅ Node SAST configuration policy OK")
PY
  local policy_rc=$?
  set -e

  if [ $policy_rc -eq 0 ]; then
    note "✅ ${node_root}: Node SAST configuration policy OK"
  else
    note "❌ ${node_root}: Node SAST configuration policy failed (see ${policy_out})"
    fail=1
  fi
}

run_config_hardening() {
  section "Configuration hardening checks"

  local out="${REPORT_DIR}/config-hardening.txt"

  set +e
  REPORT_DIR="$REPORT_DIR" OUT_FILE="$out" ./scripts/security/config_hardening.sh
  local rc=$?
  set -e

  if [ $rc -eq 0 ]; then
    note "✅ Config hardening checks passed"
  else
    note "❌ Config hardening checks failed (see ${out})"
    fail=1
  fi
}

set -e

if ! run_gitleaks; then
  fail=1
fi

run_cargo_audit
run_node_audit
run_clippy
run_node_sast
run_config_hardening

echo >>"$SUMMARY"
echo "---" >>"$SUMMARY"
if [ $fail -eq 0 ]; then
  echo "Overall: ✅ PASS" >>"$SUMMARY"
else
  echo "Overall: ❌ FAIL" >>"$SUMMARY"
fi

echo "Wrote consolidated report to ${REPORT_DIR}/" >&2
echo "Summary: ${SUMMARY}" >&2

exit "$fail"
