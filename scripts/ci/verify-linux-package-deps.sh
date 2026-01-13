#!/usr/bin/env bash

set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

fail() {
  echo "::error::verify-linux-package-deps: $*" >&2
  exit 1
}

require_cmd() {
  local cmd="$1"
  if ! command -v "$cmd" >/dev/null 2>&1; then
    fail "missing required command '$cmd' (did you forget to install it in CI?)"
  fi
}

require_cmd find
require_cmd grep
require_cmd dpkg
require_cmd dpkg-deb
require_cmd rpm

target_dirs=()

# Respect `CARGO_TARGET_DIR` if set (common in CI caching setups). Cargo interprets relative paths
# relative to the working directory used for the build (repo root in CI).
if [[ -n "${CARGO_TARGET_DIR:-}" ]]; then
  cargo_target_dir="${CARGO_TARGET_DIR}"
  if [[ "${cargo_target_dir}" != /* ]]; then
    cargo_target_dir="${repo_root}/${cargo_target_dir}"
  fi
  if [[ -d "${cargo_target_dir}" ]]; then
    target_dirs+=("${cargo_target_dir}")
  fi
fi

# Common locations:
# - workspace builds: target/
# - standalone Tauri app builds: apps/desktop/src-tauri/target
# - some setups build from apps/desktop, producing apps/desktop/target
for d in "apps/desktop/src-tauri/target" "apps/desktop/target" "target"; do
  if [[ -d "${d}" ]]; then
    target_dirs+=("${d}")
  fi
done

if [[ "${#target_dirs[@]}" -eq 0 ]]; then
  fail "no Cargo target directories found (expected CARGO_TARGET_DIR, apps/desktop/src-tauri/target, apps/desktop/target, or target)"
fi

# Prefer predictable paths rather than traversing the entire Cargo build tree (which can be large).
bundle_dirs=()
shopt -s nullglob
for target_dir in "${target_dirs[@]}"; do
  if [[ -d "${target_dir}/release/bundle" ]]; then
    bundle_dirs+=("${target_dir}/release/bundle")
  fi
  for dir in "${target_dir}"/*/release/bundle; do
    if [[ -d "${dir}" ]]; then
      bundle_dirs+=("${dir}")
    fi
  done
done
shopt -u nullglob

# Fallback (slower): scan for bundle dirs via find if the expected layout isn't present.
if [[ "${#bundle_dirs[@]}" -eq 0 ]]; then
  for target_dir in "${target_dirs[@]}"; do
    if [[ ! -d "${target_dir}" ]]; then
      continue
    fi
    while IFS= read -r -d '' dir; do
      bundle_dirs+=("$dir")
    done < <(find "$target_dir" -type d -path "*/release/bundle" -print0)
  done
fi

# De-dupe (preserve order).
if [[ "${#bundle_dirs[@]}" -gt 0 ]]; then
  declare -A seen_bundle_dirs=()
  uniq_bundle_dirs=()
  for dir in "${bundle_dirs[@]}"; do
    if [[ -n "${seen_bundle_dirs[${dir}]:-}" ]]; then
      continue
    fi
    seen_bundle_dirs["${dir}"]=1
    uniq_bundle_dirs+=("${dir}")
  done
  bundle_dirs=("${uniq_bundle_dirs[@]}")
fi

if [ "${#bundle_dirs[@]}" -eq 0 ]; then
  fail "no Tauri bundle directories found (expected something like target/**/release/bundle)"
fi

debs=()
rpms=()
shopt -s nullglob
for bundle_dir in "${bundle_dirs[@]}"; do
  debs+=("${bundle_dir}/deb/"*.deb)
  rpms+=("${bundle_dir}/rpm/"*.rpm)
done
shopt -u nullglob

# Sort/de-dupe for determinism.
if [[ "${#debs[@]}" -gt 0 ]]; then
  mapfile -t debs < <(printf '%s\n' "${debs[@]}" | sort -u)
fi
if [[ "${#rpms[@]}" -gt 0 ]]; then
  mapfile -t rpms < <(printf '%s\n' "${rpms[@]}" | sort -u)
fi

if [ "${#debs[@]}" -eq 0 ]; then
  fail "no .deb artifacts found under: ${bundle_dirs[*]}"
fi

if [ "${#rpms[@]}" -eq 0 ]; then
  fail "no .rpm artifacts found under: ${bundle_dirs[*]}"
fi

assert_contains_any() {
  local haystack="$1"
  local artifact="$2"
  local label="$3"
  shift 3
  local matched=0

  for needle in "$@"; do
    if printf '%s\n' "$haystack" | grep -Eqi "$needle"; then
      matched=1
      break
    fi
  done

  if [ "$matched" -ne 1 ]; then
    local needles_joined
    needles_joined="$(printf "%s, " "$@")"
    needles_joined="${needles_joined%, }"
    fail "$artifact: missing required dependency (${label}); expected one of: ${needles_joined}"
  fi
}

echo "verify-linux-package-deps: found ${#debs[@]} .deb and ${#rpms[@]} .rpm artifact(s)"

for deb in "${debs[@]}"; do
  echo "::group::verify-linux-package-deps: dpkg -I $(basename "$deb")"
  dpkg -I "$deb"
  echo "::endgroup::"

  depends="$(dpkg-deb -f "$deb" Depends 2>/dev/null || true)"
  if [ -z "$depends" ]; then
    fail "could not read Depends field from .deb: $deb"
  fi

  echo "verify-linux-package-deps: .deb Depends: $depends"

  # Core runtime deps (WebView + GTK + tray + SSL). Keep these checks intentionally fuzzy:
  # - Ubuntu/Debian may rename packages (e.g. *t64 suffix in Ubuntu 24.04)
  # - We accept either `libappindicator*` or `libayatana-appindicator*`
  assert_contains_any "$depends" "$deb" "WebKitGTK (webview)" "libwebkit2gtk"
  assert_contains_any "$depends" "$deb" "GTK3" "libgtk-3"
  assert_contains_any "$depends" "$deb" "AppIndicator (tray)" "appindicator"
  assert_contains_any "$depends" "$deb" "librsvg2 (icons)" "librsvg2"
  assert_contains_any "$depends" "$deb" "OpenSSL (libssl)" "libssl"
done

for rpm_path in "${rpms[@]}"; do
  echo "::group::verify-linux-package-deps: rpm -qpR $(basename "$rpm_path")"
  requires="$(rpm -qpR "$rpm_path")"
  echo "$requires"
  echo "::endgroup::"

  # `rpm -qpR` lists "capabilities" which may be package names (when explicitly declared)
  # or shared-library requirements (auto-generated). Match either so the check is robust
  # across rpm-based distros and packaging strategies.
  assert_contains_any "$requires" "$rpm_path" "WebKitGTK (webview)" "(^|[^a-z0-9])webkit2gtk" "libwebkit2gtk"
  assert_contains_any "$requires" "$rpm_path" "GTK3" "(^|[^a-z0-9])gtk3" "libgtk-3"
  assert_contains_any "$requires" "$rpm_path" "AppIndicator (tray)" "appindicator"
  assert_contains_any "$requires" "$rpm_path" "librsvg2 (icons)" "librsvg"
  assert_contains_any "$requires" "$rpm_path" "OpenSSL (libssl)" "(^|[^a-z0-9])openssl" "libssl\\.so" "libssl"
done

echo "verify-linux-package-deps: OK (core runtime dependencies present in .deb and .rpm metadata)"
