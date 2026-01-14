#!/usr/bin/env bash

set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

tmpdirs=()
cleanup() {
  for d in "${tmpdirs[@]}"; do
    rm -rf "$d" >/dev/null 2>&1 || true
  done
}
trap cleanup EXIT

fail() {
  echo "::error::verify-linux-package-deps: $*" >&2
  exit 1
}

debug_list_bundle_roots() {
  echo "::group::verify-linux-package-deps: debug bundle root listing"
  for root in "${target_dirs[@]:-}"; do
    echo "==> $root"
    ls -lah "$root/release/bundle" 2>/dev/null || true
    ls -lah "$root"/*/release/bundle 2>/dev/null || true
  done
  echo "::endgroup::"
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

TAURI_CONF_PATH="${FORMULA_TAURI_CONF_PATH:-${repo_root}/apps/desktop/src-tauri/tauri.conf.json}"
EXPECTED_DESKTOP_VERSION=""
EXPECTED_PACKAGE_NAME=""
EXPECTED_IDENTIFIER=""
PARQUET_ASSOCIATION=""
if [[ -f "${TAURI_CONF_PATH}" ]]; then
  if command -v python3 >/dev/null 2>&1; then
    EXPECTED_DESKTOP_VERSION="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print((conf.get("version") or "").strip())
PY
    )"
    EXPECTED_PACKAGE_NAME="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print((conf.get("mainBinaryName") or "").strip())
PY
    )"
    EXPECTED_IDENTIFIER="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
print((conf.get("identifier") or "").strip())
PY
    )"
    PARQUET_ASSOCIATION="$(
      python3 - "$TAURI_CONF_PATH" <<'PY' 2>/dev/null || true
import json
import sys
with open(sys.argv[1], "r", encoding="utf-8") as f:
    conf = json.load(f)
assocs = conf.get("bundle", {}).get("fileAssociations", [])
def has_parquet(a):
    if not isinstance(a, dict):
        return False
    mt = a.get("mimeType")
    if isinstance(mt, str) and mt.strip().lower() == "application/vnd.apache.parquet":
        return True
    exts = a.get("ext")
    if isinstance(exts, str):
        exts = [exts]
    if isinstance(exts, list):
        for e in exts:
            if isinstance(e, str) and e.strip().lower().lstrip(".") == "parquet":
                return True
    return False
print("1" if any(has_parquet(a) for a in assocs) else "0")
PY
    )"
  fi

  if [[ -z "${EXPECTED_DESKTOP_VERSION}" ]]; then
    # Best-effort fallback when python/json parsing isn't available.
    EXPECTED_DESKTOP_VERSION="$(
      sed -nE 's/^[[:space:]]*"version"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "${TAURI_CONF_PATH}" | head -n 1
    )"
  fi

  if [[ -z "${EXPECTED_PACKAGE_NAME}" ]]; then
    EXPECTED_PACKAGE_NAME="$(
      sed -nE 's/^[[:space:]]*"mainBinaryName"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "${TAURI_CONF_PATH}" | head -n 1
    )"
  fi

  if [[ -z "${EXPECTED_IDENTIFIER}" ]]; then
    EXPECTED_IDENTIFIER="$(
      sed -nE 's/^[[:space:]]*"identifier"[[:space:]]*:[[:space:]]*"([^"]+)".*$/\1/p' "${TAURI_CONF_PATH}" | head -n 1
    )"
  fi

  if [[ -z "${PARQUET_ASSOCIATION}" ]]; then
    if grep -q 'application/vnd.apache.parquet' "${TAURI_CONF_PATH}"; then
      PARQUET_ASSOCIATION=1
    else
      PARQUET_ASSOCIATION=0
    fi
  fi
fi
if [[ -z "${EXPECTED_DESKTOP_VERSION}" ]]; then
  fail "Unable to determine expected desktop version from ${TAURI_CONF_PATH}"
fi
: "${EXPECTED_PACKAGE_NAME:=formula-desktop}"
: "${PARQUET_ASSOCIATION:=0}"

# Parquet file associations rely on a packaged shared-mime-info definition at:
#   /usr/share/mime/packages/<identifier>.xml
# When Parquet is configured in tauri.conf.json, `identifier` is required so we can
# validate the expected filename inside the built .deb/.rpm bundles.
if [[ "${PARQUET_ASSOCIATION}" -eq 1 && -z "${EXPECTED_IDENTIFIER}" ]]; then
  fail "Parquet file association configured in ${TAURI_CONF_PATH}, but tauri \"identifier\" is missing/empty (required for /usr/share/mime/packages/<identifier>.xml)."
fi
if [[ "${PARQUET_ASSOCIATION}" -eq 1 ]]; then
  if [[ "${EXPECTED_IDENTIFIER}" == */* || "${EXPECTED_IDENTIFIER}" == *\\* ]]; then
    fail "Parquet association configured but tauri identifier is not a valid filename (contains path separators): ${EXPECTED_IDENTIFIER}"
  fi
fi

echo "verify-linux-package-deps: expected desktop version: ${EXPECTED_DESKTOP_VERSION}"
echo "verify-linux-package-deps: expected package/binary name: ${EXPECTED_PACKAGE_NAME}"
echo "verify-linux-package-deps: expected tauri identifier: ${EXPECTED_IDENTIFIER}"
echo "verify-linux-package-deps: parquet association configured: ${PARQUET_ASSOCIATION}"

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
    # Avoid a full traversal of the Cargo target directory (which can be large in CI) by
    # bounding the search depth. The bundle layout is expected to be shallow:
    #   <target_dir>/release/bundle
    #   <target_dir>/<triple>/release/bundle
    done < <(find "$target_dir" -maxdepth 6 -type d -path "*/release/bundle" -print0)
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
  echo "::error::verify-linux-package-deps: no Tauri bundle directories found (expected something like <target>/release/bundle or <target>/<triple>/release/bundle)" >&2
  echo "Searched target dirs:" >&2
  printf '  - %s\n' "${target_dirs[@]}" >&2
  debug_list_bundle_roots
  exit 1
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
  # Keep fallback discovery bounded: bundle dirs should be shallow and scanning deeply can be
  # surprisingly expensive if unexpected extracted artifacts exist.
  done < <(find "${bundle_dirs[@]}" -maxdepth 6 -type f -name "*.deb" -print0)
fi

if [[ "${#rpms[@]}" -eq 0 ]]; then
  while IFS= read -r -d '' f; do
    rpms+=("$f")
  # Keep fallback discovery bounded: bundle dirs should be shallow and scanning deeply can be
  # surprisingly expensive if unexpected extracted artifacts exist.
  done < <(find "${bundle_dirs[@]}" -maxdepth 6 -type f -name "*.rpm" -print0)
fi

# Sort/de-dupe for determinism.
if [[ "${#debs[@]}" -gt 0 ]]; then
  mapfile -t debs < <(printf '%s\n' "${debs[@]}" | sort -u)
fi
if [[ "${#rpms[@]}" -gt 0 ]]; then
  mapfile -t rpms < <(printf '%s\n' "${rpms[@]}" | sort -u)
fi

if [ "${#debs[@]}" -eq 0 ]; then
  echo "::error::verify-linux-package-deps: no .deb artifacts found under bundle dirs:" >&2
  printf '  - %s\n' "${bundle_dirs[@]}" >&2
  debug_list_bundle_roots
  exit 1
fi

if [ "${#rpms[@]}" -eq 0 ]; then
  echo "::error::verify-linux-package-deps: no .rpm artifacts found under bundle dirs:" >&2
  printf '  - %s\n' "${bundle_dirs[@]}" >&2
  debug_list_bundle_roots
  exit 1
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

assert_contains_rich_or() {
  local haystack="$1"
  local artifact="$2"
  local label="$3"
  local left_re="$4"
  local right_re="$5"

  # RPM rich dependencies are expressed as boolean expressions, e.g.:
  #   (webkit2gtk4.1 or libwebkit2gtk-4_1-0)
  #
  # We require distro-specific alternatives to be declared as a *single* rich dependency
  # entry (instead of listing both packages separately as two Requires, which would make
  # the RPM uninstallable on at least one distro family).
  local pattern
  pattern="(${left_re}).*\\bor\\b.*(${right_re})|(${right_re}).*\\bor\\b.*(${left_re})"

  if ! printf '%s\n' "$haystack" | grep -Eqi "$pattern"; then
    fail "$artifact: missing RPM rich dependency OR expression (${label}); expected a line containing both '${left_re}' and '${right_re}' joined by 'or'"
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

  deb_pkg="$(dpkg-deb -f "$deb" Package 2>/dev/null || true)"
  if [ -z "$deb_pkg" ]; then
    fail "could not read Package field from .deb: $deb"
  fi
  deb_pkg="$(printf '%s' "$deb_pkg" | head -n 1 | tr -d '\r')"
  deb_pkg="${deb_pkg#"${deb_pkg%%[![:space:]]*}"}"
  deb_pkg="${deb_pkg%"${deb_pkg##*[![:space:]]}"}"
  if [[ "$deb_pkg" != "$EXPECTED_PACKAGE_NAME" ]]; then
    fail "$deb: package name mismatch (dpkg Package). Expected ${EXPECTED_PACKAGE_NAME}, found ${deb_pkg}"
  fi

  deb_version="$(dpkg-deb -f "$deb" Version 2>/dev/null || true)"
  if [ -z "$deb_version" ]; then
    fail "could not read Version field from .deb: $deb"
  fi
  deb_version="$(printf '%s' "$deb_version" | head -n 1 | tr -d '\r')"
  # Trim whitespace without introducing extra dependencies (keep this script lightweight).
  deb_version="${deb_version#"${deb_version%%[![:space:]]*}"}"
  deb_version="${deb_version%"${deb_version##*[![:space:]]}"}"
  deb_version_no_epoch="${deb_version}"
  if [[ "$deb_version_no_epoch" == *:* ]]; then
    # Debian version format: [epoch:]upstream[-revision]
    deb_version_no_epoch="${deb_version_no_epoch#*:}"
  fi
  if [[ "$deb_version_no_epoch" != "$EXPECTED_DESKTOP_VERSION" ]]; then
    # Allow Debian revision suffixes (e.g. 1.2.3-1, 1.2.3-beta.1-1), but avoid accepting
    # non-numeric suffixes like 1.2.3-beta.1 when EXPECTED_DESKTOP_VERSION is 1.2.3.
    if [[ "$deb_version_no_epoch" == "${EXPECTED_DESKTOP_VERSION}-"* ]]; then
      deb_revision="${deb_version_no_epoch#${EXPECTED_DESKTOP_VERSION}-}"
      if [[ -z "$deb_revision" || ! "$deb_revision" =~ ^[0-9][0-9A-Za-z.+~]*$ ]]; then
        fail "$deb: version mismatch (dpkg Version). Expected ${EXPECTED_DESKTOP_VERSION} (or ${EXPECTED_DESKTOP_VERSION}-<debian-revision>), found ${deb_version}"
      fi
    else
      fail "$deb: version mismatch (dpkg Version). Expected ${EXPECTED_DESKTOP_VERSION} (or ${EXPECTED_DESKTOP_VERSION}-<debian-revision>), found ${deb_version}"
    fi
  fi

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
  assert_contains_any "$depends" "$deb" "shared-mime-info (MIME database integration)" "shared-mime-info"
  assert_contains_any "$depends" "$deb" "WebKitGTK 4.1 (webview)" "libwebkit2gtk-4\\.1"
  assert_contains_any "$depends" "$deb" "GTK3" "libgtk-3"
  assert_contains_any "$depends" "$deb" "AppIndicator (tray)" "appindicator"
  assert_contains_any "$depends" "$deb" "librsvg2 (icons)" "librsvg2"
  assert_contains_any "$depends" "$deb" "OpenSSL (libssl)" "libssl"
  # We ship a shared-mime-info definition for Parquet so `*.parquet` maps to
  # application/vnd.apache.parquet (many distros don't include this by default).
  # Require shared-mime-info so install triggers can run update-mime-database.
  assert_contains_any "$depends" "$deb" "shared-mime-info (MIME database)" "shared-mime-info"

  # Ensure the packaged binary itself is stripped (no accidental debug/symbol sections shipped).
  echo "::group::verify-linux-package-deps: stripped binary check (deb) $(basename "$deb")"
  tmpdir="$(mktemp -d)"
  tmpdirs+=("$tmpdir")
  dpkg-deb -x "$deb" "$tmpdir"
  assert_stripped_elf "$tmpdir/usr/bin/${EXPECTED_PACKAGE_NAME}" "$(basename "$deb")"
  if [[ "${PARQUET_ASSOCIATION}" -eq 1 ]]; then
    mime_xml="$tmpdir/usr/share/mime/packages/${EXPECTED_IDENTIFIER}.xml"
    if [[ ! -f "${mime_xml}" ]]; then
      fail "$deb: missing Parquet shared-mime-info definition file at: ${mime_xml}"
    fi
    if ! grep -Fq 'application/vnd.apache.parquet' "${mime_xml}" || ! grep -Fq '*.parquet' "${mime_xml}"; then
      fail "$deb: Parquet shared-mime-info definition file is missing expected content: ${mime_xml}"
    fi
  fi
  rm -rf "$tmpdir"
  echo "::endgroup::"
done

for rpm_path in "${rpms[@]}"; do
  rpm_version="$(rpm -qp --queryformat '%{VERSION}\n' "$rpm_path" 2>/dev/null | head -n 1 | tr -d '\r' || true)"
  if [[ -z "$rpm_version" ]]; then
    fail "$rpm_path: could not read RPM %{VERSION} via rpm -qp --queryformat"
  fi
  if [[ "$rpm_version" != "$EXPECTED_DESKTOP_VERSION" ]]; then
    fail "$rpm_path: version mismatch (RPM %{VERSION}). Expected ${EXPECTED_DESKTOP_VERSION}, found ${rpm_version}"
  fi

  rpm_name="$(rpm -qp --queryformat '%{NAME}\n' "$rpm_path" 2>/dev/null | head -n 1 | tr -d '\r' || true)"
  if [[ -z "$rpm_name" ]]; then
    fail "$rpm_path: could not read RPM %{NAME} via rpm -qp --queryformat"
  fi
  if [[ "$rpm_name" != "$EXPECTED_PACKAGE_NAME" ]]; then
    fail "$rpm_path: name mismatch (RPM %{NAME}). Expected ${EXPECTED_PACKAGE_NAME}, found ${rpm_name}"
  fi

  echo "::group::verify-linux-package-deps: rpm -qpR $(basename "$rpm_path")"
  requires="$(rpm -qpR "$rpm_path")"
  echo "$requires"
  echo "::endgroup::"

  # `rpm -qpR` lists RPM Requires "capabilities". For the RPMs produced by Tauri (tauri-bundler),
  # these are primarily driven by the explicit `bundle.linux.rpm.depends` list in `tauri.conf.json`.
  #
  # We intentionally validate these explicit package requirements so the RPM declares distro package
  # names (Fedora/RHEL + openSUSE) rather than relying on automatic ELF dependency scanning.

  # 1) Explicit package requirements (Fedora/RHEL + openSUSE naming via RPM rich deps).
  assert_contains_any "$requires" "$rpm_path" "shared-mime-info (MIME database integration)" "shared-mime-info"
  assert_contains_rich_or "$requires" "$rpm_path" "WebKitGTK 4.1 package (Fedora/RHEL vs openSUSE)" "webkit2gtk4\\.1" "libwebkit2gtk-4_1"
  assert_contains_rich_or "$requires" "$rpm_path" "GTK3 package (Fedora/RHEL vs openSUSE)" "gtk3" "libgtk-3-0"
  assert_contains_rich_or "$requires" "$rpm_path" "AppIndicator/Ayatana package (Fedora/RHEL vs openSUSE)" "(libayatana-appindicator-gtk3|libappindicator-gtk3)" "(libayatana-appindicator3-1|libappindicator3-1)"
  assert_contains_rich_or "$requires" "$rpm_path" "librsvg package (Fedora/RHEL vs openSUSE)" "librsvg2" "librsvg-2-2"
  assert_contains_rich_or "$requires" "$rpm_path" "OpenSSL package (Fedora/RHEL vs openSUSE)" "openssl-libs" "libopenssl3"

  assert_contains_any "$requires" "$rpm_path" "WebKitGTK 4.1 package (webview)" "webkit2gtk4\\.1" "libwebkit2gtk-4_1"
  assert_contains_any "$requires" "$rpm_path" "GTK3 package" "(^|[^a-z0-9])gtk3([^a-z0-9]|$)" "libgtk-3-0"
  assert_contains_any "$requires" "$rpm_path" "AppIndicator/Ayatana package (tray)" \
    "libayatana-appindicator-gtk3" \
    "libappindicator-gtk3" \
    "libayatana-appindicator3-1" \
    "libappindicator3-1"
  assert_contains_any "$requires" "$rpm_path" "librsvg package (icons)" "librsvg2" "librsvg-2-2"
  assert_contains_any "$requires" "$rpm_path" "OpenSSL package" "openssl-libs" "libopenssl3"
  # Ensure update-mime-database is available so packaged MIME definitions (e.g. Parquet)
  # are registered on install.
  assert_contains_any "$requires" "$rpm_path" "shared-mime-info (MIME database)" "shared-mime-info"

  # Ensure the packaged binary itself is stripped (no accidental debug/symbol sections shipped).
  echo "::group::verify-linux-package-deps: stripped binary check (rpm) $(basename "$rpm_path")"
  tmpdir="$(mktemp -d)"
  tmpdirs+=("$tmpdir")
  (
    cd "$tmpdir"
    rpm2cpio "$rpm_path" | cpio -idm --quiet --no-absolute-filenames
  )
  assert_stripped_elf "$tmpdir/usr/bin/${EXPECTED_PACKAGE_NAME}" "$(basename "$rpm_path")"
  if [[ "${PARQUET_ASSOCIATION}" -eq 1 ]]; then
    mime_xml="$tmpdir/usr/share/mime/packages/${EXPECTED_IDENTIFIER}.xml"
    if [[ ! -f "${mime_xml}" ]]; then
      fail "$rpm_path: missing Parquet shared-mime-info definition file at: ${mime_xml}"
    fi
    if ! grep -Fq 'application/vnd.apache.parquet' "${mime_xml}" || ! grep -Fq '*.parquet' "${mime_xml}"; then
      fail "$rpm_path: Parquet shared-mime-info definition file is missing expected content: ${mime_xml}"
    fi
  fi
  rm -rf "$tmpdir"
  echo "::endgroup::"
done

echo "verify-linux-package-deps: OK (core runtime dependencies present in .deb and .rpm metadata)"
