#!/usr/bin/env bash
#
# Validate a Linux DEB bundle produced by the Tauri desktop build.
#
# This script is intended for CI (Ubuntu host + optional Docker), but can also be used
# locally. It performs:
#   1) Host "static" validation via dpkg-deb queries + payload inspection.
#   2) Optional installability validation inside an Ubuntu container.
#
# Usage:
#   ./scripts/validate-linux-deb.sh
#   ./scripts/validate-linux-deb.sh --deb path/to/formula-desktop.deb
#   ./scripts/validate-linux-deb.sh --no-container
#
set -euo pipefail

SCRIPT_NAME="$(basename "$0")"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
ORIG_PWD="$(pwd)"

# Ensure relative-path discovery works regardless of the caller's cwd.
cd "$REPO_ROOT"

usage() {
  cat <<EOF
${SCRIPT_NAME} - validate a Linux DEB bundle (Tauri desktop build)

Usage:
  ${SCRIPT_NAME} [--deb <path>] [--no-container] [--image <ubuntu-image>]

Options:
  --deb <path>        Validate a specific .deb (or a directory containing .deb files).
                      If omitted, the script searches common Tauri output locations:
                        - \$CARGO_TARGET_DIR/release/bundle/deb/*.deb (if set)
                        - \$CARGO_TARGET_DIR/*/release/bundle/deb/*.deb (if set)
                        - apps/desktop/src-tauri/target/release/bundle/deb/*.deb
                        - apps/desktop/src-tauri/target/*/release/bundle/deb/*.deb
                        - apps/desktop/target/release/bundle/deb/*.deb
                        - apps/desktop/target/*/release/bundle/deb/*.deb
                        - target/release/bundle/deb/*.deb
                        - target/*/release/bundle/deb/*.deb
  --no-container      Skip the Ubuntu container installability check.
  --image <image>     Ubuntu image to use for the container step (default: ubuntu:24.04).
  -h, --help          Show this help text.

Environment:
  DOCKER_PLATFORM
                           Optional docker --platform override (default: host architecture).
  FORMULA_TAURI_CONF_PATH
                           Optional path override for apps/desktop/src-tauri/tauri.conf.json (useful for local testing).
                           If the path is relative, it is resolved relative to the repo root.
  FORMULA_DEB_NAME_OVERRIDE
                           Override the expected Debian package name (dpkg-deb Package field) for validation purposes.
                           This affects the expected /usr/share/doc/<package>/... doc dir, but does NOT affect the
                           expected /usr/bin/<mainBinaryName> path inside the package.
EOF
}

err() {
  if [[ "${GITHUB_ACTIONS:-}" == "true" ]]; then
    echo "::error::${SCRIPT_NAME}: $*" >&2
  else
    echo "${SCRIPT_NAME}: ERROR: $*" >&2
  fi
}

note() {
  echo "${SCRIPT_NAME}: $*"
}

die() {
  err "$@"
  exit 1
}

DEB_OVERRIDE=""
NO_CONTAINER=0
UBUNTU_IMAGE="ubuntu:24.04"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --deb)
      DEB_OVERRIDE="${2:-}"
      [[ -n "${DEB_OVERRIDE}" ]] || die "--deb requires a path argument"
      shift 2
      ;;
    --no-container)
      NO_CONTAINER=1
      shift
      ;;
    --image)
      UBUNTU_IMAGE="${2:-}"
      [[ -n "${UBUNTU_IMAGE}" ]] || die "--image requires an image argument"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      die "Unknown argument: $1 (use --help)"
      ;;
  esac
done

require_cmd() {
  command -v "$1" >/dev/null 2>&1 || die "Required command not found in PATH: $1"
}

require_docker() {
  command -v docker >/dev/null 2>&1 || die "docker is required for container validation (install docker or rerun with --no-container)"

  set +e
  docker info >/dev/null 2>&1
  local info_status=$?
  set -e
  if [[ "${info_status}" -ne 0 ]]; then
    die "docker is installed but the daemon is not available (docker info failed). Start Docker or rerun with --no-container."
  fi
}

detect_docker_platform() {
  local arch
  arch="$(uname -m)"
  case "${arch}" in
    x86_64) echo "linux/amd64" ;;
    aarch64|arm64) echo "linux/arm64" ;;
    *) echo "" ;;
  esac
}

# Force Docker to use the host architecture image variant by default. This avoids confusing
# `exec format error` failures when a mismatched (e.g. ARM) image tag is present locally.
DOCKER_PLATFORM="${DOCKER_PLATFORM:-$(detect_docker_platform)}"

require_cmd dpkg-deb

# Allow overriding the Tauri config path for tests or custom build layouts.
# If a relative path is provided, interpret it relative to the repo root.
TAURI_CONF="${FORMULA_TAURI_CONF_PATH:-$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json}"
if [[ "$TAURI_CONF" != /* ]]; then
  TAURI_CONF="$REPO_ROOT/$TAURI_CONF"
fi
if [[ ! -f "$TAURI_CONF" ]]; then
  die "Missing Tauri config: $TAURI_CONF"
fi

read_tauri_conf_value() {
  local key="$1"
  if command -v python3 >/dev/null 2>&1; then
    python3 - "$TAURI_CONF" "$key" <<'PY'
import json
import sys

path = sys.argv[1]
key = sys.argv[2]
with open(path, "r", encoding="utf-8") as f:
    conf = json.load(f)
val = conf.get(key, "")
if isinstance(val, str):
    print(val.strip())
PY
    return 0
  fi

  if command -v node >/dev/null 2>&1; then
    node -e '
      const fs = require("fs");
      const path = process.argv[1];
      const key = process.argv[2];
      const conf = JSON.parse(fs.readFileSync(path, "utf8"));
      const val = conf?.[key];
      if (typeof val === "string") process.stdout.write(val.trim());
    ' "$TAURI_CONF" "$key"
    return 0
  fi

  die "Neither python3 nor node is available to parse ${TAURI_CONF} (required for version/name checks)."
}

EXPECTED_VERSION="$(read_tauri_conf_value version)"
if [[ -z "$EXPECTED_VERSION" ]]; then
  die "Expected $TAURI_CONF to contain a non-empty \"version\" field."
fi

# The main installed binary name should match tauri.conf.json mainBinaryName.
EXPECTED_MAIN_BINARY="$(read_tauri_conf_value mainBinaryName)"
if [[ -z "$EXPECTED_MAIN_BINARY" ]]; then
  EXPECTED_MAIN_BINARY="formula-desktop"
fi

# The app identifier (reverse-DNS). We use it for MIME definition filenames under:
#   /usr/share/mime/packages/<identifier>.xml
EXPECTED_IDENTIFIER="$(read_tauri_conf_value identifier)"
if [[ -z "$EXPECTED_IDENTIFIER" ]]; then
  die "Expected $TAURI_CONF to contain a non-empty \"identifier\" field."
fi
# The app identifier is used as a filename on Linux:
#   /usr/share/mime/packages/<identifier>.xml
if [[ "${EXPECTED_IDENTIFIER}" == */* || "${EXPECTED_IDENTIFIER}" == *\\* ]]; then
  die "Expected $TAURI_CONF identifier to be a valid filename (no '/' or '\\' path separators). Found: ${EXPECTED_IDENTIFIER}"
fi
EXPECTED_MIME_DEFINITION_BASENAME="${EXPECTED_IDENTIFIER}.xml"
EXPECTED_MIME_DEFINITION_PATH="/usr/share/mime/packages/${EXPECTED_MIME_DEFINITION_BASENAME}"

# Expected deep-link URL scheme handlers (x-scheme-handler/<scheme>) that should be
# advertised by the installed desktop entry.
#
# Source of truth: apps/desktop/src-tauri/tauri.conf.json → plugins.deep-link.desktop.schemes.
# Default to "formula" when missing/unparseable.
declare -a EXPECTED_DEEP_LINK_SCHEMES=()
if command -v python3 >/dev/null 2>&1; then
  mapfile -t EXPECTED_DEEP_LINK_SCHEMES < <(
    python3 - "$TAURI_CONF" <<'PY' 2>/dev/null || true
import json
import re
import sys

path = sys.argv[1]
try:
    with open(path, "r", encoding="utf-8") as f:
        conf = json.load(f)
except Exception:
    sys.exit(0)

plugins = conf.get("plugins") or {}
deep_link = plugins.get("deep-link") or {}
desktop = deep_link.get("desktop")
schemes = set()

def normalize(value: str) -> str:
    v = value.strip().lower()
    v = re.sub(r"[:/]+$", "", v)
    return v

def add_from_protocol(protocol):
    if not isinstance(protocol, dict):
        return
    raw = protocol.get("schemes")
    values = []
    if isinstance(raw, str):
        values = [raw]
    elif isinstance(raw, list):
        values = [v for v in raw if isinstance(v, str)]
    for v in values:
        norm = normalize(v)
        if norm:
            schemes.add(norm)

if isinstance(desktop, list):
    for protocol in desktop:
        add_from_protocol(protocol)
else:
    add_from_protocol(desktop)

for scheme in sorted(schemes):
    print(scheme)
PY
  )
fi
if [[ ${#EXPECTED_DEEP_LINK_SCHEMES[@]} -eq 0 ]] && command -v node >/dev/null 2>&1; then
  mapfile -t EXPECTED_DEEP_LINK_SCHEMES < <(
    node - "$TAURI_CONF" <<'NODE' 2>/dev/null || true
const fs = require("node:fs");
const configPath = process.argv[2];
let conf;
try {
  conf = JSON.parse(fs.readFileSync(configPath, "utf8"));
} catch {
  process.exit(0);
}

const deepLink = conf?.plugins?.["deep-link"];
const desktop = deepLink?.desktop;
const schemes = new Set();

const normalize = (value) => String(value).trim().toLowerCase().replace(/[:/]+$/, "");

const addFromProtocol = (protocol) => {
  if (!protocol || typeof protocol !== "object") return;
  const raw = protocol.schemes;
  const values = typeof raw === "string" ? [raw] : Array.isArray(raw) ? raw : [];
  for (const v of values) {
    if (typeof v !== "string") continue;
    const norm = normalize(v);
    if (norm) schemes.add(norm);
  }
};

if (Array.isArray(desktop)) {
  for (const protocol of desktop) addFromProtocol(protocol);
} else {
  addFromProtocol(desktop);
}

if (schemes.size === 0) process.exit(0);
for (const scheme of Array.from(schemes).sort()) console.log(scheme);
NODE
  )
fi
if [[ ${#EXPECTED_DEEP_LINK_SCHEMES[@]} -eq 0 ]]; then
  EXPECTED_DEEP_LINK_SCHEMES=("formula")
fi

for scheme in "${EXPECTED_DEEP_LINK_SCHEMES[@]}"; do
  if [[ "$scheme" == *:* || "$scheme" == */* ]]; then
    die "Invalid deep-link scheme configured in ${TAURI_CONF} (plugins[\"deep-link\"].desktop.schemes): '${scheme}'. Expected scheme names only (no ':' or '/' characters)."
  fi
done
declare -a EXPECTED_SCHEME_MIMES=()
for scheme in "${EXPECTED_DEEP_LINK_SCHEMES[@]}"; do
  EXPECTED_SCHEME_MIMES+=("x-scheme-handler/${scheme}")
done
EXPECTED_SCHEME_MIMES_ENV="$(
  IFS=';'
  echo "${EXPECTED_SCHEME_MIMES[*]}"
)"

# Debian package name should typically match the main binary name, but allow overriding
# just the Package field/doc dir name for validation purposes.
EXPECTED_DEB_NAME="${FORMULA_DEB_NAME_OVERRIDE:-$EXPECTED_MAIN_BINARY}"
if [[ -z "$EXPECTED_DEB_NAME" ]]; then
  EXPECTED_DEB_NAME="$EXPECTED_MAIN_BINARY"
fi

abs_path() {
  local p="$1"
  if [[ "$p" != /* ]]; then
    p="${REPO_ROOT}/${p}"
  fi
  local dir
  dir="$(dirname "$p")"
  if [[ -d "$dir" ]]; then
    dir="$(cd "$dir" && pwd -P)"
    echo "${dir}/$(basename "$p")"
  else
    echo "$p"
  fi
}

find_debs() {
  local -a debs=()

  if [[ -n "${DEB_OVERRIDE}" ]]; then
    # Resolve relative paths against the invocation directory, not the repo root.
    if [[ "${DEB_OVERRIDE}" != /* ]]; then
      DEB_OVERRIDE="${ORIG_PWD}/${DEB_OVERRIDE}"
    fi
    if [[ -d "${DEB_OVERRIDE}" ]]; then
      while IFS= read -r -d '' f; do debs+=("$(abs_path "$f")"); done < <(find "${DEB_OVERRIDE}" -maxdepth 1 -type f -name '*.deb' -print0)
    else
      debs+=("$(abs_path "${DEB_OVERRIDE}")")
    fi
  else
    local -a roots=()
    if [[ -n "${CARGO_TARGET_DIR:-}" ]]; then
      cargo_target_dir="${CARGO_TARGET_DIR}"
      if [[ "${cargo_target_dir}" != /* ]]; then
        cargo_target_dir="${REPO_ROOT}/${cargo_target_dir}"
      fi
      if [[ -d "${cargo_target_dir}" ]]; then
        roots+=("${cargo_target_dir}")
      fi
    fi
    for root in "apps/desktop/src-tauri/target" "apps/desktop/target" "target"; do
      if [[ -d "${root}" ]]; then
        roots+=("${root}")
      fi
    done

    # Canonicalize + de-dupe.
    local -A seen_roots=()
    local -a uniq_roots=()
    for root in "${roots[@]}"; do
      if [[ "${root}" != /* ]]; then
        root="${REPO_ROOT}/${root}"
      fi
      [[ -d "${root}" ]] || continue
      root="$(cd "${root}" && pwd -P)"
      if [[ -n "${seen_roots[${root}]:-}" ]]; then
        continue
      fi
      seen_roots["${root}"]=1
      uniq_roots+=("${root}")
    done
    roots=("${uniq_roots[@]}")

    local nullglob_was_set=0
    if shopt -q nullglob; then
      nullglob_was_set=1
    fi
    shopt -s nullglob
    for root in "${roots[@]}"; do
      debs+=("${root}/release/bundle/deb/"*.deb)
      debs+=("${root}/"*/release/bundle/deb/*.deb)
    done
    if [[ "${nullglob_was_set}" -eq 0 ]]; then
      shopt -u nullglob
    fi

    if [[ ${#debs[@]} -eq 0 ]]; then
      for root in "${roots[@]}"; do
        # Avoid an unbounded scan of the Cargo target directory: the bundle layout we care about is
        # shallow (<target>/<triple>/release/bundle/deb/*.deb). Keep a conservative maxdepth so
        # unexpected layouts still work, but we don't traverse the full build tree on failures.
        while IFS= read -r -d '' f; do debs+=("$(abs_path "$f")"); done < <(find "${root}" -maxdepth 6 -type f -path '*/release/bundle/deb/*.deb' -print0 2>/dev/null || true)
      done
    fi
  fi

  if [[ ${#debs[@]} -eq 0 ]]; then
    return 0
  fi

  local -A seen=()
  local -a unique=()
  local deb_path
  for deb_path in "${debs[@]}"; do
    seen["$deb_path"]=1
  done
  for deb_path in "${!seen[@]}"; do
    unique+=("$deb_path")
  done
  IFS=$'\n' unique=($(printf '%s\n' "${unique[@]}" | sort))
  unset IFS
  printf '%s\n' "${unique[@]}"
}

assert_contains_any() {
  local haystack="$1"
  local label="$2"
  shift 2
  local matched=0
  local needle
  for needle in "$@"; do
    if printf '%s\n' "$haystack" | grep -Eqi "$needle"; then
      matched=1
      break
    fi
  done
  if [[ "$matched" -ne 1 ]]; then
    die "DEB metadata missing required dependency (${label}); expected one of: $*"
  fi
}

validate_desktop_integration_extracted() {
  local package_root="$1"

  local mime_xml="$package_root${EXPECTED_MIME_DEFINITION_PATH}"
  if [[ ! -f "$mime_xml" ]]; then
    err "Extracted payload missing shared-mime-info definition file: ${EXPECTED_MIME_DEFINITION_PATH#/}"
    return 1
  fi
  if ! grep -Fq 'application/vnd.apache.parquet' "$mime_xml" || ! grep -Fq '*.parquet' "$mime_xml"; then
    err "Extracted Parquet MIME definition file is missing expected content: ${mime_xml#${package_root}/}"
    return 1
  fi

  local applications_dir="$package_root/usr/share/applications"
  if [[ ! -d "$applications_dir" ]]; then
    err "Extracted payload missing /usr/share/applications (expected at: ${applications_dir})"
    return 1
  fi

  local -a desktop_files=()
  while IFS= read -r -d '' desktop_file; do
    desktop_files+=("$desktop_file")
  done < <(find "$applications_dir" -maxdepth 2 -type f -name '*.desktop' -print0 2>/dev/null || true)
  if [[ ${#desktop_files[@]} -eq 0 ]]; then
    err "No .desktop files found under extracted payload: ${applications_dir}"
    return 1
  fi

  # Filter to the desktop entry (or entries) that actually launch our app. This avoids
  # accidentally validating unrelated .desktop files from dependency packages.
  local expected_binary_escaped
  expected_binary_escaped="$(printf '%s' "$EXPECTED_MAIN_BINARY" | sed -e 's/[][(){}.^$*+?|\\]/\\&/g')"
  local exec_token_re
  exec_token_re="(^|[[:space:]])[\"']?([^[:space:]]*/)?${expected_binary_escaped}[\"']?([[:space:]]|$)"

  local -a matched_desktop_files=()
  local desktop_file
  for desktop_file in "${desktop_files[@]}"; do
    local exec_line
    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    local exec_value
    exec_value="$(printf '%s' "$exec_line" | sed -E "s/^[[:space:]]*Exec[[:space:]]*=[[:space:]]*//I")"
    if [[ -n "$exec_value" ]] && printf '%s' "$exec_value" | grep -Eq "${exec_token_re}"; then
      matched_desktop_files+=("$desktop_file")
    fi
  done

  if [[ ${#matched_desktop_files[@]} -eq 0 ]]; then
    err "No extracted .desktop files appear to target the expected executable '${EXPECTED_MAIN_BINARY}' in their Exec= entry."
    err "Extracted .desktop files inspected:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${package_root}/}"
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [[ -z "$exec_line" ]]; then
        exec_line="(no Exec= entry)"
      fi
      echo "  - ${rel}: ${exec_line}" >&2
    done
    return 1
  fi

  desktop_files=("${matched_desktop_files[@]}")

  local required_xlsx_mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  local required_parquet_mime="application/vnd.apache.parquet"
  local spreadsheet_mime_regex
  spreadsheet_mime_regex='xlsx|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|application/vnd\.ms-excel|application/vnd\.ms-excel\.sheet\.macroEnabled\.12|application/vnd\.ms-excel\.sheet\.binary\.macroEnabled\.12|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.template|application/vnd\.ms-excel\.template\.macroEnabled\.12|application/vnd\.ms-excel\.addin\.macroEnabled\.12|text/csv'

  local has_any_mimetype=0
  local has_spreadsheet_mime=0
  local has_xlsx_integration=0
  local has_parquet_mime=0
  local bad_exec_count=0
  declare -A found_scheme_mimes=()

  for desktop_file in "${desktop_files[@]}"; do
    local mime_line
    mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    if [[ -z "$mime_line" ]]; then
      continue
    fi
    has_any_mimetype=1
    local mime_value
    mime_value="$(printf '%s' "$mime_line" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//")"

    if printf '%s' "$mime_value" | grep -Fqi "$required_parquet_mime"; then
      has_parquet_mime=1
    fi

    # Tokenize MimeType= (semicolon-delimited) so scheme handler checks are exact matches,
    # not substring matches (e.g. avoid treating x-scheme-handler/<scheme>-extra as matching
    # expected x-scheme-handler/<scheme>).
    local -a mime_tokens=()
    IFS=';' read -r -a mime_tokens <<<"$mime_value"
    local -A mime_set=()
    local token
    for token in "${mime_tokens[@]}"; do
      token="${token#"${token%%[![:space:]]*}"}"
      token="${token%"${token##*[![:space:]]}"}"
      token="$(printf '%s' "$token" | tr '[:upper:]' '[:lower:]')"
      if [[ -n "$token" ]]; then
        mime_set["$token"]=1
      fi
    done

    local has_any_expected_scheme_in_file=0
    local scheme_mime
    for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
      if [[ -n "${mime_set["$scheme_mime"]+x}" ]]; then
        found_scheme_mimes["$scheme_mime"]=1
        has_any_expected_scheme_in_file=1
      fi
    done

    if [[ "$has_any_expected_scheme_in_file" -eq 1 ]]; then
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [[ -z "$exec_line" ]]; then
        err "Extracted desktop entry ${desktop_file#${package_root}/} is missing an Exec= entry (required for URL scheme handlers)"
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        err "Extracted desktop entry ${desktop_file#${package_root}/} Exec= does not include a URL placeholder (%U/%u/%F/%f): ${exec_line}"
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi

    if printf '%s' "$mime_value" | grep -Fqi "xlsx"; then
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Fqi "$required_xlsx_mime"; then
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Eqi "$spreadsheet_mime_regex"; then
      has_spreadsheet_mime=1
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [[ -z "$exec_line" ]]; then
        err "Extracted desktop entry ${desktop_file#${package_root}/} is missing an Exec= entry (required for file associations)"
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        err "Extracted desktop entry ${desktop_file#${package_root}/} Exec= does not include a file/URL placeholder (%U/%u/%F/%f): ${exec_line}"
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi
  done

  if [[ "$has_any_mimetype" -ne 1 ]]; then
    err "No extracted .desktop file contained a MimeType= entry (file associations missing)."
    return 1
  fi
  if [[ "$has_xlsx_integration" -ne 1 ]]; then
    err "No extracted .desktop MimeType= entry advertised xlsx support (expected substring 'xlsx' or MIME '${required_xlsx_mime}')."
    return 1
  fi
  if [[ "$has_parquet_mime" -ne 1 ]]; then
    err "No extracted .desktop MimeType= entry advertised Parquet support (expected MIME '${required_parquet_mime}')."
    return 1
  fi
  if [[ "$has_spreadsheet_mime" -ne 1 ]]; then
    err "No extracted .desktop MimeType= entry advertised spreadsheet MIME types (expected xlsx/csv/etc)."
    return 1
  fi
  local -a missing_scheme_mimes=()
  local scheme_mime
  for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
    if [[ -z "${found_scheme_mimes["$scheme_mime"]+x}" ]]; then
      missing_scheme_mimes+=("$scheme_mime")
    fi
  done
  if [[ ${#missing_scheme_mimes[@]} -ne 0 ]]; then
    err "No extracted .desktop MimeType= entry advertised the expected URL scheme handler(s): ${missing_scheme_mimes[*]}"
    return 1
  fi
  if [[ "$bad_exec_count" -ne 0 ]]; then
    err "One or more extracted .desktop entries had invalid Exec= lines for file association / URL scheme handling."
    return 1
  fi
}

validate_static() {
  local deb_path="$1"
  [[ -f "${deb_path}" ]] || die "DEB not found: ${deb_path}"
  [[ -s "${deb_path}" ]] || die "DEB is empty: ${deb_path}"

  note "Static validation: ${deb_path}"

  local deb_version
  deb_version="$(dpkg-deb -f "${deb_path}" Version 2>/dev/null | tr -d '\r' | head -n 1 || true)"
  [[ -n "$deb_version" ]] || die "Failed to read DEB Version field (dpkg-deb -f) for: ${deb_path}"

  # Debian version format: [epoch:]upstream[-revision]
  local deb_version_no_epoch
  deb_version_no_epoch="${deb_version}"
  if [[ "$deb_version_no_epoch" == *:* ]]; then
    deb_version_no_epoch="${deb_version_no_epoch#*:}"
  fi

  if [[ "$deb_version_no_epoch" != "${EXPECTED_VERSION}" ]]; then
    # Allow a Debian revision suffix after the expected upstream version, e.g.:
    #   0.1.0-1
    #   0.1.0-beta.1-1
    #
    # NOTE: We intentionally do not accept arbitrary `EXPECTED_VERSION-*` prefixes here because
    # a version like "0.1.0-beta.1" would otherwise be treated as a "revision" of 0.1.0 and
    # could mask stale/mis-versioned artifacts.
    if [[ "$deb_version_no_epoch" == "${EXPECTED_VERSION}-"* ]]; then
      local deb_revision
      deb_revision="${deb_version_no_epoch#${EXPECTED_VERSION}-}"
      # Debian revision strings are typically numeric (e.g. -1, -0ubuntu1). Require the
      # suffix to begin with a digit to avoid accepting pre-release mismatches like
      # "0.1.0-beta.1" when EXPECTED_VERSION is "0.1.0".
      if [[ -z "$deb_revision" || ! "$deb_revision" =~ ^[0-9][0-9A-Za-z.+~]*$ ]]; then
        die "DEB version mismatch for ${deb_path}: expected ${EXPECTED_VERSION} (or ${EXPECTED_VERSION}-<debian-revision>), found ${deb_version}"
      fi
    else
      die "DEB version mismatch for ${deb_path}: expected ${EXPECTED_VERSION} (or ${EXPECTED_VERSION}-<debian-revision>), found ${deb_version}"
    fi
  fi

  local deb_pkg
  deb_pkg="$(dpkg-deb -f "${deb_path}" Package 2>/dev/null | tr -d '\r' | head -n 1 || true)"
  [[ -n "$deb_pkg" ]] || die "Failed to read DEB Package field (dpkg-deb -f) for: ${deb_path}"
  if [[ "$deb_pkg" != "$EXPECTED_DEB_NAME" ]]; then
    die "DEB package name mismatch for ${deb_path}: expected ${EXPECTED_DEB_NAME}, found ${deb_pkg}"
  fi

  local depends
  depends="$(dpkg-deb -f "${deb_path}" Depends 2>/dev/null | tr -d '\r' | head -n 1 || true)"
  [[ -n "$depends" ]] || die "Failed to read DEB Depends field (dpkg-deb -f) for: ${deb_path}"

  # Runtime deps: keep checks fuzzy to allow t64 transitions (Ubuntu 24.04) and AppIndicator variants.
  assert_contains_any "$depends" "shared-mime-info (MIME database integration)" "shared-mime-info"
  assert_contains_any "$depends" "WebKitGTK 4.1 (webview)" "libwebkit2gtk-4\\.1"
  assert_contains_any "$depends" "GTK3" "libgtk-3"
  assert_contains_any "$depends" "AppIndicator/Ayatana (tray)" "appindicator"
  assert_contains_any "$depends" "librsvg2 (icons)" "librsvg2"
  assert_contains_any "$depends" "OpenSSL (libssl)" "libssl"

  local contents
  contents="$(dpkg-deb -c "${deb_path}")" || die "dpkg-deb -c failed for: ${deb_path}"
  local file_list
  file_list="$(printf '%s\n' "$contents" | awk '{print $NF}')"

  if ! grep -qx "./usr/bin/${EXPECTED_MAIN_BINARY}" <<<"${file_list}"; then
    die "DEB payload missing expected desktop binary path: ./usr/bin/${EXPECTED_MAIN_BINARY}"
  fi
  if ! grep -Eq '^\.?/usr/share/applications/[^/]+\.desktop$' <<<"${file_list}"; then
    die "DEB payload missing expected .desktop file under: ./usr/share/applications/"
  fi
  for filename in LICENSE NOTICE; do
    if ! grep -qx "./usr/share/doc/${EXPECTED_DEB_NAME}/${filename}" <<<"${file_list}"; then
      die "DEB payload missing compliance file: ./usr/share/doc/${EXPECTED_DEB_NAME}/${filename}"
    fi
  done
  if ! grep -qx "./${EXPECTED_MIME_DEFINITION_PATH#/}" <<<"${file_list}"; then
    die "DEB payload missing Parquet shared-mime-info definition: ./${EXPECTED_MIME_DEFINITION_PATH#/}"
  fi

  # Extract payload and validate desktop integration metadata.
  local tmpdir
  tmpdir="$(mktemp -d)"
  cleanup_tmpdir() {
    rm -rf "${tmpdir}" >/dev/null 2>&1 || true
  }
  trap cleanup_tmpdir EXIT
  dpkg-deb -x "${deb_path}" "${tmpdir}" || die "dpkg-deb -x failed for: ${deb_path}"

  [[ -x "${tmpdir}/usr/bin/${EXPECTED_MAIN_BINARY}" ]] || die "Extracted payload missing executable: usr/bin/${EXPECTED_MAIN_BINARY}"
  for filename in LICENSE NOTICE; do
    [[ -f "${tmpdir}/usr/share/doc/${EXPECTED_DEB_NAME}/${filename}" ]] || die "Extracted payload missing compliance file: usr/share/doc/${EXPECTED_DEB_NAME}/${filename}"
  done

  # Always run bash-based validation of the extracted .desktop metadata so we have a
  # python-free fallback *and* avoid false positives from substring matching in MimeType=
  # checks (e.g. x-scheme-handler/<scheme>-extra should not satisfy x-scheme-handler/<scheme>).
  if ! validate_desktop_integration_extracted "${tmpdir}"; then
    die "Desktop integration validation failed (inspect .desktop MimeType/Exec entries)"
  fi

  # Additionally, prefer strict validation against tauri.conf.json (when python is available)
  # so we catch missing MIME types, scheme handlers, compliance artifacts, and Parquet
  # shared-mime-info wiring in the built artifact (not just in config).
  if command -v python3 >/dev/null 2>&1; then
    note "Static desktop integration validation (verify extracted DEB payload)"
    if ! python3 "$REPO_ROOT/scripts/ci/verify_linux_desktop_integration.py" \
      --package-root "$tmpdir" \
      --tauri-config "$TAURI_CONF" \
      --expected-main-binary "$EXPECTED_MAIN_BINARY" \
      --doc-package-name "$EXPECTED_DEB_NAME"; then
      die "Desktop integration validation failed (inspect .desktop MimeType/Exec entries)"
    fi
  fi

  cleanup_tmpdir
  trap - EXIT
}

validate_container() {
  local deb_path="$1"
  require_docker

  deb_path="$(abs_path "${deb_path}")"
  local deb_basename
  deb_basename="$(basename "${deb_path}")"

  local -a docker_platform_args=()
  if [[ -n "${DOCKER_PLATFORM}" ]]; then
    docker_platform_args=(--platform "${DOCKER_PLATFORM}")
  fi

  local mount_dir
  mount_dir="$(mktemp -d)"
  if ! ln "${deb_path}" "${mount_dir}/${deb_basename}" 2>/dev/null; then
    cp "${deb_path}" "${mount_dir}/${deb_basename}"
  fi

  note "Container validation (Ubuntu): ${deb_path}"
  note "Using image: ${UBUNTU_IMAGE}"

  docker run --rm "${docker_platform_args[@]}" \
    -e "FORMULA_EXPECTED_SCHEME_MIMES=${EXPECTED_SCHEME_MIMES_ENV}" \
    -v "${mount_dir}:/deb:ro" \
    "${UBUNTU_IMAGE}" \
    bash -euxo pipefail -c '
    export DEBIAN_FRONTEND=noninteractive
    apt-get update
    apt-get install -y --no-install-recommends /deb/*.deb
    bin="'"${EXPECTED_MAIN_BINARY}"'"
    doc_pkg="'"${EXPECTED_DEB_NAME}"'"
    binary_path="/usr/bin/${bin}"
    test -x "${binary_path}"
    test -f "/usr/share/doc/${doc_pkg}/LICENSE"
    test -f "/usr/share/doc/${doc_pkg}/NOTICE"

    # Validate that installer-time MIME integration ran successfully.
    #
    # Many distros do not ship a Parquet glob by default, so the package ships a
    # shared-mime-info definition under /usr/share/mime/packages and relies on
    # shared-mime-info triggers to rebuild /usr/share/mime/globs2.
    mime_xml="'"${EXPECTED_MIME_DEFINITION_PATH}"'"
    test -f "${mime_xml}"
    grep -F "application/vnd.apache.parquet" "${mime_xml}"
    grep -F "*.parquet" "${mime_xml}"
    grep -Eq "application/vnd\\.apache\\.parquet:.*\\*\\.parquet" /usr/share/mime/globs2

    # Validate desktop integration metadata is present in the installed .desktop entry.
    desktop_files=(/usr/share/applications/*.desktop)
    if [ ! -e "${desktop_files[0]}" ]; then
      echo "No .desktop files found under /usr/share/applications after DEB install." >&2
      ls -lah /usr/share/applications || true
      exit 1
    fi

    # Filter to the .desktop entries that actually launch our app (Exec= references the expected binary).
    bin_re="$(printf "%s" "${bin}" | sed -e "s/[][(){}.^$*+?|\\\\]/\\\\&/g")"
    exec_token_re="(^|[[:space:]])([^[:space:]]*/)?${bin_re}([[:space:]]|$)"
    matching_desktop_files=()
    for f in "${desktop_files[@]}"; do
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"
      exec_value="$(printf "%s" "${exec_line}" | sed -E "s/^[[:space:]]*Exec[[:space:]]*=[[:space:]]*//I")"
      # Normalize quoted Exec tokens (e.g. Exec="/usr/bin/formula-desktop" %U)
      exec_value="${exec_value//\"/}"
      sq="$(printf "%b" "\\x27")"
      exec_value="${exec_value//${sq}/}"
      if [ -n "${exec_value}" ] && printf "%s" "${exec_value}" | grep -Eq "${exec_token_re}"; then
        matching_desktop_files+=("$f")
      fi
    done

    if [ "${#matching_desktop_files[@]}" -eq 0 ]; then
      echo "No installed .desktop files appear to target the expected executable ${bin} in their Exec= entry." >&2
      echo "Installed .desktop files inspected:" >&2
      for f in "${desktop_files[@]}"; do
        exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"
        if [ -z "${exec_line}" ]; then
          exec_line="(no Exec= entry)"
        fi
        echo "  - ${f}: ${exec_line}" >&2
      done
      exit 1
    fi

    desktop_file="${matching_desktop_files[0]}"
    echo "Installed desktop entry: ${desktop_file}"
    grep -E "^[[:space:]]*(Exec|MimeType)=" "${desktop_file}" || true
    grep -Eq "^[[:space:]]*Exec=.*%[uUfF]" "${desktop_file}"
    mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "${desktop_file}" | head -n 1 || true)"
    mime_value="$(printf "%s" "${mime_line}" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//I")"
    if [ -z "${mime_value}" ]; then
      echo "Installed desktop entry is missing MimeType= (required for file associations + URL scheme handlers): ${desktop_file}" >&2
      exit 1
    fi
    IFS=";" read -r -a mime_tokens <<< "${mime_value}"
    declare -A mime_set=()
    for token in "${mime_tokens[@]}"; do
      token="${token#"${token%%[![:space:]]*}"}"
      token="${token%"${token##*[![:space:]]}"}"
      token="$(printf "%s" "${token}" | tr "[:upper:]" "[:lower:]")"
      if [ -n "${token}" ]; then
        mime_set["$token"]=1
      fi
    done
    expected_scheme_mimes_raw="${FORMULA_EXPECTED_SCHEME_MIMES:-}"
    if [ -z "${expected_scheme_mimes_raw}" ]; then
      default_scheme="formula"
      expected_scheme_mimes_raw="x-scheme-handler/${default_scheme}"
    fi
    IFS=";" read -r -a expected_scheme_mimes <<< "${expected_scheme_mimes_raw}"
    for i in "${!expected_scheme_mimes[@]}"; do
      token="${expected_scheme_mimes[$i]}"
      token="${token#"${token%%[![:space:]]*}"}"
      token="${token%"${token##*[![:space:]]}"}"
      token="$(printf "%s" "${token}" | tr "[:upper:]" "[:lower:]")"
      expected_scheme_mimes[$i]="${token}"
    done
    for scheme_mime in "${expected_scheme_mimes[@]}"; do
      if [ -n "${scheme_mime}" ] && [ -z "${mime_set["$scheme_mime"]+x}" ]; then
        echo "Missing URL scheme handler in desktop entry MimeType=: ${scheme_mime}" >&2
        echo "Observed MimeType= value: ${mime_value}" >&2
        exit 1
      fi
    done
    grep -qi "application/vnd\\.openxmlformats-officedocument\\.spreadsheetml\\.sheet" "${desktop_file}"
    grep -qi "application/vnd\\.apache\\.parquet" "${desktop_file}"

    # Ensure shared library dependencies are present.
    set +e
    ldd_out="$(ldd "${binary_path}" 2>&1)"
    ldd_status=$?
    set -e
    echo "${ldd_out}"
    if echo "${ldd_out}" | grep -q "not found"; then
      echo "Missing shared libraries detected:" >&2
      echo "${ldd_out}" | grep "not found" >&2 || true
      exit 1
    fi
    # ldd returns non-zero for static binaries ("not a dynamic executable"). Treat that as OK
    # as long as we did not detect missing shared libraries.
    if [ "${ldd_status}" -ne 0 ] && ! echo "${ldd_out}" | grep -q "not a dynamic executable" && ! echo "${ldd_out}" | grep -q "statically linked"; then
      echo "ldd exited with status ${ldd_status}" >&2
      exit 1
    fi
  '

  rm -rf "${mount_dir}" >/dev/null 2>&1 || true
}

main() {
  mapfile -t debs < <(find_debs)
  if [[ ${#debs[@]} -eq 0 ]]; then
    die "No .deb artifacts found. Build one with Tauri, or pass --deb <path>."
  fi

  note "Found ${#debs[@]} DEB artifact(s)."
  local deb_path
  for deb_path in "${debs[@]}"; do
    validate_static "${deb_path}"
    if [[ "${NO_CONTAINER}" -eq 0 ]]; then
      validate_container "${deb_path}"
    fi
  done

  note "OK"
}

main
