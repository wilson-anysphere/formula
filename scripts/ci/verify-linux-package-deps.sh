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
require_cmd file
require_cmd readelf
require_cmd dpkg
require_cmd dpkg-deb
require_cmd rpm
require_cmd rpm2cpio
require_cmd cpio
require_cmd file
require_cmd readelf
require_cmd rpm2cpio
require_cmd cpio

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

# Normalize to absolute real paths (and de-dupe). This avoids scanning the same directory twice
# (e.g. when CARGO_TARGET_DIR=target and we also add "target" as a default candidate).
declare -A seen_target_dirs=()
normalized_target_dirs=()
for dir in "${target_dirs[@]}"; do
  abs="${dir}"
  if [[ "${abs}" != /* ]]; then
    abs="${repo_root}/${abs}"
  fi
  if [[ ! -d "${abs}" ]]; then
    continue
  fi
  abs="$(cd "${abs}" && pwd -P)"
  if [[ -n "${seen_target_dirs[${abs}]:-}" ]]; then
    continue
  fi
  seen_target_dirs["${abs}"]=1
  normalized_target_dirs+=("${abs}")
done
target_dirs=("${normalized_target_dirs[@]}")

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

# Fallback: if the bundle layout changes, locate packages anywhere under the bundle dirs.
if [[ "${#debs[@]}" -eq 0 ]]; then
  while IFS= read -r -d '' f; do
    debs+=("$f")
  done < <(find "${bundle_dirs[@]}" -type f -name "*.deb" -print0)
fi

if [[ "${#rpms[@]}" -eq 0 ]]; then
  while IFS= read -r -d '' f; do
    rpms+=("$f")
  done < <(find "${bundle_dirs[@]}" -type f -name "*.rpm" -print0)
fi

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
echo "::group::verify-linux-package-deps: bundle directories"
printf ' - %s\n' "${bundle_dirs[@]}"
echo "::endgroup::"

echo "::group::verify-linux-package-deps: discovered .deb artifacts"
printf ' - %s\n' "${debs[@]}"
echo "::endgroup::"

echo "::group::verify-linux-package-deps: discovered .rpm artifacts"
printf ' - %s\n' "${rpms[@]}"
echo "::endgroup::"

assert_stripped_elf() {
  local elf_path="$1"
  local artifact="$2"

  if [[ ! -f "$elf_path" ]]; then
    fail "$artifact: expected ELF binary not found at: $elf_path"
  fi

  local out
  out="$(file -b "$elf_path" || true)"
  echo "verify-linux-package-deps: file $elf_path -> $out"
  if echo "$out" | grep -q "not stripped"; then
    fail "$artifact: binary is not stripped (expected stripped): $out"
  fi

  if readelf -S --wide "$elf_path" 2>/dev/null | grep -q '\.debug_'; then
    fail "$artifact: ELF contains .debug_* sections (expected stripped)"
  fi
}

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
  #
  # But do enforce WebKitGTK **4.1** specifically (Tauri v2.9 + wry expects WebKitGTK 4.1
  # in this repo); accidentally drifting to 4.0 would break runtime compatibility.
  assert_contains_any "$depends" "$deb" "WebKitGTK 4.1 (webview)" "libwebkit2gtk-4\\.1"
  assert_contains_any "$depends" "$deb" "GTK3" "libgtk-3"
  assert_contains_any "$depends" "$deb" "AppIndicator (tray)" "appindicator"
  assert_contains_any "$depends" "$deb" "librsvg2 (icons)" "librsvg2"
  assert_contains_any "$depends" "$deb" "OpenSSL (libssl)" "libssl"

  # Ensure the packaged binary itself is stripped (no accidental debug/symbol sections shipped).
  echo "::group::verify-linux-package-deps: stripped binary check (deb) $(basename "$deb")"
  tmpdir="$(mktemp -d)"
  dpkg-deb -x "$deb" "$tmpdir"
  assert_stripped_elf "$tmpdir/usr/bin/formula-desktop" "$(basename "$deb")"
  rm -rf "$tmpdir"
  echo "::endgroup::"
done

for rpm_path in "${rpms[@]}"; do
  echo "::group::verify-linux-package-deps: rpm -qpR $(basename "$rpm_path")"
  requires="$(rpm -qpR "$rpm_path")"
  echo "$requires"
  echo "::endgroup::"

  # `rpm -qpR` lists "capabilities" which may be package names (when explicitly declared)
  # or shared-library requirements (auto-generated).
  #
  # We validate **both**:
  # 1) explicit package requirements (driven by `bundle.linux.rpm.depends` in `tauri.conf.json`),
  #    so the RPM declares distro package names (Fedora/RHEL + openSUSE) instead of relying only
  #    on auto-generated ELF soname requirements.
  # 2) presence of the expected shared-library requirements (defense-in-depth).

  # 1) Explicit package requirements (Fedora/RHEL + openSUSE naming via RPM rich deps).
  assert_contains_any "$requires" "$rpm_path" "WebKitGTK 4.1 package (webview)" "webkit2gtk4\\.1" "libwebkit2gtk-4_1"
  assert_contains_any "$requires" "$rpm_path" "GTK3 package" "(^|[^a-z0-9])gtk3([^a-z0-9]|$)" "libgtk-3-0"
  assert_contains_any "$requires" "$rpm_path" "AppIndicator/Ayatana package (tray)" \
    "libayatana-appindicator-gtk3" \
    "libappindicator-gtk3" \
    "libayatana-appindicator3-1" \
    "libappindicator3-1"
  assert_contains_any "$requires" "$rpm_path" "librsvg package (icons)" "librsvg2" "librsvg-2-2"
  assert_contains_any "$requires" "$rpm_path" "OpenSSL package" "openssl-libs" "libopenssl3"

  # 2) Shared-library requirements (auto-generated).
  #
  # Keep these checks narrow: some runtime deps (notably AppIndicator + OpenSSL) may be loaded
  # dynamically or pulled in indirectly, so they may not appear as direct ELF NEEDED entries.
  # We enforce those via the explicit package Requires above.
  assert_contains_any "$requires" "$rpm_path" "WebKitGTK 4.1 (webview) soname" "libwebkit2gtk-4\\.1"
  assert_contains_any "$requires" "$rpm_path" "GTK3 soname" "libgtk-3"

  # Ensure the packaged binary itself is stripped (no accidental debug/symbol sections shipped).
  echo "::group::verify-linux-package-deps: stripped binary check (rpm) $(basename "$rpm_path")"
  tmpdir="$(mktemp -d)"
  (
    cd "$tmpdir"
    rpm2cpio "$rpm_path" | cpio -idm --quiet --no-absolute-filenames
  )
  assert_stripped_elf "$tmpdir/usr/bin/formula-desktop" "$(basename "$rpm_path")"
  rm -rf "$tmpdir"
  echo "::endgroup::"
done

echo "verify-linux-package-deps: OK (core runtime dependencies present in .deb and .rpm metadata)"
