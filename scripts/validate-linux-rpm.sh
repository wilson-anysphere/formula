#!/usr/bin/env bash
#
# Validate a Linux RPM bundle produced by the Tauri desktop build.
#
# This script is intended for CI (Ubuntu host + Docker), but can also be used
# locally. It performs:
#   1) Host "static" validation via rpm queries.
#   2) Optional installability validation inside a Fedora container.
#
# Usage:
#   ./scripts/validate-linux-rpm.sh
#   ./scripts/validate-linux-rpm.sh --rpm path/to/formula-desktop.rpm
#   ./scripts/validate-linux-rpm.sh --no-container
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
${SCRIPT_NAME} - validate a Linux RPM bundle (Tauri desktop build)

Usage:
  ${SCRIPT_NAME} [--rpm <path>] [--no-container] [--image <fedora-image>]

Options:
  --rpm <path>        Validate a specific .rpm (or a directory containing .rpm files).
                      If omitted, the script searches common Tauri output locations:
                         - \$CARGO_TARGET_DIR/release/bundle/rpm/*.rpm (if set)
                         - \$CARGO_TARGET_DIR/*/release/bundle/rpm/*.rpm (if set)
                         - apps/desktop/src-tauri/target/release/bundle/rpm/*.rpm
                         - apps/desktop/src-tauri/target/*/release/bundle/rpm/*.rpm
                         - apps/desktop/target/release/bundle/rpm/*.rpm
                         - apps/desktop/target/*/release/bundle/rpm/*.rpm
                         - target/release/bundle/rpm/*.rpm
                         - target/*/release/bundle/rpm/*.rpm
  --no-container      Skip the Fedora container installability check.
                        Note: we still extract the RPM payload to validate desktop integration
                        metadata (MimeType= in *.desktop). This requires rpm2cpio + cpio on the host.
  --image <image>     Fedora image to use for the container step (default: fedora:40).
  -h, --help          Show this help text.

Environment:
  DOCKER_PLATFORM
                       Optional docker --platform override (default: host architecture).
  FORMULA_TAURI_CONF_PATH
                       Optional path override for apps/desktop/src-tauri/tauri.conf.json (useful for local testing).
                       If the path is relative, it is resolved relative to the repo root.
  FORMULA_RPM_NAME_OVERRIDE
                       Override the expected RPM %{NAME} package name for validation purposes.
                       (Does NOT affect the expected /usr/bin/<mainBinaryName> path inside the RPM.)
EOF
}

err() {
  if [[ "${GITHUB_ACTIONS:-}" == "true" ]]; then
    # GitHub Actions error annotation (still readable locally).
    echo "::error::${SCRIPT_NAME}: $*" >&2
  else
    echo "${SCRIPT_NAME}: ERROR: $*" >&2
  fi
}

note() {
  echo "${SCRIPT_NAME}: $*"
}

warn() {
  echo "${SCRIPT_NAME}: WARN: $*" >&2
}

die() {
  err "$@"
  exit 1
}

RPM_OVERRIDE=""
NO_CONTAINER=0
FEDORA_IMAGE="fedora:40"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --rpm)
      RPM_OVERRIDE="${2:-}"
      [[ -n "${RPM_OVERRIDE}" ]] || die "--rpm requires a path argument"
      shift 2
      ;;
    --no-container)
      NO_CONTAINER=1
      shift
      ;;
    --image)
      FEDORA_IMAGE="${2:-}"
      [[ -n "${FEDORA_IMAGE}" ]] || die "--image requires an image argument"
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

  # `docker` can exist on PATH even when the daemon isn't running (common on dev machines).
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
    aarch64 | arm64) echo "linux/arm64" ;;
    *) echo "" ;;
  esac
}

# Force Docker to use the host architecture image variant by default. This avoids confusing
# `exec format error` failures when a mismatched (e.g. ARM) image tag is present locally.
#
# Override for debugging by exporting DOCKER_PLATFORM explicitly.
DOCKER_PLATFORM="${DOCKER_PLATFORM:-$(detect_docker_platform)}"

require_cmd rpm

TAURI_CONF="${FORMULA_TAURI_CONF_PATH:-$REPO_ROOT/apps/desktop/src-tauri/tauri.conf.json}"
if [[ "${TAURI_CONF}" != /* ]]; then
  TAURI_CONF="${REPO_ROOT}/${TAURI_CONF}"
fi
if [[ ! -f "$TAURI_CONF" ]]; then
  die "Missing Tauri config: $TAURI_CONF"
fi

read_tauri_conf_value() {
  local key="$1"
  if command -v python3 >/dev/null 2>&1; then
    python3 - "$TAURI_CONF" "$key" <<'PY' 2>/dev/null || true
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
    # Best-effort fallback when python is unavailable.
    node - "$TAURI_CONF" "$key" <<'NODE' 2>/dev/null || true
const fs = require("node:fs");
const path = process.argv[2];
const key = process.argv[3];
try {
  const conf = JSON.parse(fs.readFileSync(path, "utf8"));
  const val = conf?.[key];
  if (typeof val === "string") process.stdout.write(val.trim());
} catch {
  // ignore
}
NODE
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

# RPM %{NAME} (package name) should match our decided package name. By default we keep this
# in sync with tauri.conf.json mainBinaryName, but allow overriding just the *package name*
# for validation purposes (some distros may prefer different naming conventions).
EXPECTED_RPM_NAME="${FORMULA_RPM_NAME_OVERRIDE:-$EXPECTED_MAIN_BINARY}"
if [[ -z "$EXPECTED_RPM_NAME" ]]; then
  EXPECTED_RPM_NAME="$EXPECTED_MAIN_BINARY"
fi

rel_path() {
  local p="$1"
  if [[ "$p" == "${REPO_ROOT}/"* ]]; then
    echo "${p#${REPO_ROOT}/}"
  else
    echo "$p"
  fi
}

abs_path() {
  local p="$1"
  if [[ "$p" != /* ]]; then
    p="${REPO_ROOT}/${p}"
  fi
  # Canonicalize the directory component so we avoid duplicate (abs vs rel) paths.
  # If the directory does not exist (e.g. user typo), return the original string
  # so downstream checks can emit a clearer error.
  local dir
  dir="$(dirname "$p")"
  if [[ -d "$dir" ]]; then
    dir="$(cd "$dir" && pwd -P)"
    echo "${dir}/$(basename "$p")"
  else
    echo "$p"
  fi
}

# Spreadsheet file association metadata we expect the Linux desktop entry to advertise.
# `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` is the canonical
# xlsx MIME type; we allow a small set of other spreadsheet-ish MIME types as a
# fallback to avoid false negatives across distros.
REQUIRED_XLSX_MIME="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SPREADSHEET_MIME_REGEX='xlsx|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|application/vnd\.ms-excel|application/vnd\.ms-excel\.sheet\.macroEnabled\.12|application/vnd\.ms-excel\.sheet\.binary\.macroEnabled\.12|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.template|application/vnd\.ms-excel\.template\.macroEnabled\.12|application/vnd\.ms-excel\.addin\.macroEnabled\.12|text/csv'

validate_desktop_mime_associations_extracted() {
  local rpm_path="$1"

  if ! command -v rpm2cpio >/dev/null 2>&1; then
    die "rpm2cpio not found. It is required to validate .desktop MimeType entries when running with --no-container. Install rpm2cpio (and cpio) or re-run without --no-container."
  fi
  if ! command -v cpio >/dev/null 2>&1; then
    die "cpio not found. It is required to validate .desktop MimeType entries when running with --no-container. Install cpio (and rpm2cpio) or re-run without --no-container."
  fi

  local tmpdir
  tmpdir="$(mktemp -d)"
  # RETURN traps apply to the entire call stack (they fire again when the caller returns).
  # Expand the tmpdir value at trap installation time, and clear the trap after the first
  # invocation so we don't re-run cleanup after the function returns.
  trap "rm -rf \"${tmpdir}\" >/dev/null 2>&1 || true; trap - RETURN" RETURN

  note "Static desktop integration validation (extract RPM payload): ${rpm_path}"

  (
    cd "$tmpdir"
    rpm2cpio "$rpm_path" | cpio -idm --quiet --no-absolute-filenames
  ) || {
    err "Failed to extract RPM payload with rpm2cpio/cpio: ${rpm_path}"
    return 1
  }

  # Ensure we ship a shared-mime-info definition for Parquet so `*.parquet` resolves to the
  # advertised MIME type (`application/vnd.apache.parquet`) on distros whose shared-mime-info DB
  # does not include a Parquet glob by default.
  local mime_xml="$tmpdir${EXPECTED_MIME_DEFINITION_PATH}"
  if [ ! -f "$mime_xml" ]; then
    err "RPM payload missing Parquet shared-mime-info definition after extraction: ${EXPECTED_MIME_DEFINITION_PATH#/}"
    return 1
  fi
  if ! grep -Fq 'application/vnd.apache.parquet' "$mime_xml" || ! grep -Fq '*.parquet' "$mime_xml"; then
    err "RPM Parquet shared-mime-info definition file is missing expected content: ${mime_xml#${tmpdir}/}"
    return 1
  fi

  declare -a desktop_files=()
  local applications_dir="$tmpdir/usr/share/applications"
  if [ -d "$applications_dir" ]; then
    while IFS= read -r -d '' desktop_file; do
      desktop_files+=("$desktop_file")
  done < <(find "$applications_dir" -maxdepth 2 -type f -name '*.desktop' -print0 2>/dev/null || true)
  fi

  if [ "${#desktop_files[@]}" -eq 0 ]; then
    # Some packaging layouts place desktop entries under `share/applications` (or
    # `usr/local/share/applications`). Check those predictable locations before
    # resorting to a broader filesystem scan.
    local alt_dir
    for alt_dir in "$tmpdir/usr/local/share/applications" "$tmpdir/share/applications"; do
      if [ -d "$alt_dir" ]; then
        while IFS= read -r -d '' desktop_file; do
          desktop_files+=("$desktop_file")
        done < <(find "$alt_dir" -maxdepth 2 -type f -name '*.desktop' -print0 2>/dev/null || true)
      fi
    done
  fi

  if [ "${#desktop_files[@]}" -eq 0 ]; then
    # Fallback: accept any desktop file in the payload to aid debugging.
    #
    # Keep this scan bounded: extracted packages can contain large `usr/lib` trees, and an
    # unbounded recursive `find "$tmpdir"` can be slow in CI.
    while IFS= read -r -d '' desktop_file; do
      desktop_files+=("$desktop_file")
    done < <(
      find "$tmpdir" \
        -maxdepth 8 \
        \( \
          -path "$tmpdir/usr/lib*" -o \
          -path "$tmpdir/usr/bin*" -o \
          -path "$tmpdir/usr/sbin*" -o \
          -path "$tmpdir/bin*" -o \
          -path "$tmpdir/lib*" -o \
          -path "$tmpdir/lib64*" -o \
          -path "$tmpdir/usr/share/doc*" -o \
          -path "$tmpdir/usr/share/icons*" -o \
          -path "$tmpdir/usr/share/locale*" -o \
          -path "$tmpdir/usr/share/man*" -o \
          -path "$tmpdir/usr/share/mime*" \
        \) -prune -o \
        -type f \
        -name '*.desktop' \
        -print0 2>/dev/null || true
    )
  fi

  if [ "${#desktop_files[@]}" -eq 0 ]; then
    err "RPM payload did not contain any .desktop files after extraction. Expected at least one under /usr/share/applications/."
    return 1
  fi

  local desktop_file

  # Filter to the desktop entry (or entries) that actually launch our app. It's
  # possible for dependency packages to install unrelated .desktop files; we
  # should validate *our* desktop integration.
  local expected_binary_escaped
  expected_binary_escaped="$(printf '%s' "$EXPECTED_MAIN_BINARY" | sed -e 's/[][(){}.^$*+?|\\]/\\&/g')"
  local exec_token_re
  exec_token_re="(^|[[:space:]])[\"']?([^[:space:]]*/)?${expected_binary_escaped}[\"']?([[:space:]]|$)"

  declare -a matched_desktop_files=()
  for desktop_file in "${desktop_files[@]}"; do
    local exec_line
    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    local exec_value
    exec_value="$(printf '%s' "$exec_line" | sed -E "s/^[[:space:]]*Exec[[:space:]]*=[[:space:]]*//I")"
    if [[ -n "$exec_value" ]] && printf '%s' "$exec_value" | grep -Eq "${exec_token_re}"; then
      matched_desktop_files+=("$desktop_file")
    fi
  done

  if [ "${#matched_desktop_files[@]}" -eq 0 ]; then
    err "No extracted .desktop files appear to target the expected executable '${EXPECTED_MAIN_BINARY}' in their Exec= entry."
    err "Extracted .desktop files inspected:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${tmpdir}/}"
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

  local has_any_mimetype=0
  local has_spreadsheet_mime=0
  local has_xlsx_mime=0
  local has_xlsx_integration=0
  local has_parquet_mime=0
  local bad_exec_count=0
  local required_parquet_mime="application/vnd.apache.parquet"
  declare -A found_scheme_mimes=()

  for desktop_file in "${desktop_files[@]}"; do
    local mime_line
    mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
    if [ -z "$mime_line" ]; then
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
      if [ -n "$token" ]; then
        mime_set["$token"]=1
      fi
    done

    local has_any_expected_scheme_in_file=0
    local scheme_mime
    for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
      if [ -n "${mime_set["$scheme_mime"]+x}" ]; then
        found_scheme_mimes["$scheme_mime"]=1
        has_any_expected_scheme_in_file=1
      fi
    done

    if printf '%s' "$mime_value" | grep -Fqi "$REQUIRED_XLSX_MIME"; then
      has_xlsx_mime=1
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Fqi "xlsx"; then
      has_xlsx_integration=1
    fi
    if printf '%s' "$mime_value" | grep -Eqi "$SPREADSHEET_MIME_REGEX"; then
      has_spreadsheet_mime=1

      # File associations require a placeholder token (%U/%u/%F/%f) in Exec= so the
      # OS passes the opened file path/URL into the app.
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [ -z "$exec_line" ]; then
        err "Extracted desktop entry ${desktop_file#${tmpdir}/} is missing an Exec= entry (required for file associations)"
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        err "Extracted desktop entry ${desktop_file#${tmpdir}/} Exec= does not include a file/URL placeholder (%U/%u/%F/%f): ${exec_line}"
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi

    if [ "$has_any_expected_scheme_in_file" -eq 1 ]; then
      # URL scheme handlers also require a placeholder token (%U/%u/%F/%f) in Exec= so the
      # OS passes the opened URL into the app.
      local exec_line
      exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$desktop_file" | head -n 1 || true)"
      if [ -z "$exec_line" ]; then
        err "Extracted desktop entry ${desktop_file#${tmpdir}/} is missing an Exec= entry (required for URL scheme handlers)"
        bad_exec_count=$((bad_exec_count + 1))
      elif ! printf '%s' "$exec_line" | grep -Eq '%[uUfF]'; then
        err "Extracted desktop entry ${desktop_file#${tmpdir}/} Exec= does not include a URL placeholder (%U/%u/%F/%f): ${exec_line}"
        bad_exec_count=$((bad_exec_count + 1))
      fi
    fi
  done

  if [ "$has_any_mimetype" -ne 1 ]; then
    err "No extracted .desktop file contained a MimeType= entry (file associations missing)."
    err "Expected MimeType= to advertise spreadsheet MIME types based on $(rel_path "$TAURI_CONF") bundle.fileAssociations."
    err "Extracted .desktop files inspected:"
    for desktop_file in "${desktop_files[@]}"; do
      echo "  - ${desktop_file#${tmpdir}/}" >&2
    done
    return 1
  fi

  if [ "$has_xlsx_integration" -ne 1 ]; then
    err "No extracted .desktop MimeType= entry advertised xlsx support (file associations missing)."
    err "Expected MimeType= to include substring 'xlsx' or MIME '${REQUIRED_XLSX_MIME}'."
    err "MimeType entries found:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${tmpdir}/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    return 1
  fi

  if [ "$has_parquet_mime" -ne 1 ]; then
    err "No extracted .desktop MimeType= entry advertised Parquet support (file associations missing)."
    err "Expected MimeType= to include '${required_parquet_mime}'."
    err "MimeType entries found:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${tmpdir}/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    return 1
  fi

  if [ "$has_spreadsheet_mime" -ne 1 ]; then
    err "No extracted .desktop MimeType= entry advertised spreadsheet/xlsx support (file associations missing)."
    err "Expected MimeType= to include '${REQUIRED_XLSX_MIME}' (xlsx) or another spreadsheet MIME type."
    err "MimeType entries found:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${tmpdir}/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    return 1
  fi

  if [ "$bad_exec_count" -ne 0 ]; then
    err "One or more extracted .desktop entries had invalid Exec= lines for file association handling."
    return 1
  fi

  local -a missing_scheme_mimes=()
  local scheme_mime
  for scheme_mime in "${EXPECTED_SCHEME_MIMES[@]}"; do
    if [ -z "${found_scheme_mimes["$scheme_mime"]+x}" ]; then
      missing_scheme_mimes+=("$scheme_mime")
    fi
  done
  if [ "${#missing_scheme_mimes[@]}" -ne 0 ]; then
    err "No extracted .desktop MimeType= entry advertised the expected URL scheme handler(s): ${missing_scheme_mimes[*]}."
    err "Expected MimeType= to include: ${EXPECTED_SCHEME_MIMES[*]}"
    err "MimeType entries found:"
    for desktop_file in "${desktop_files[@]}"; do
      local rel
      rel="${desktop_file#${tmpdir}/}"
      local lines
      lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$desktop_file" || true)"
      if [ -n "$lines" ]; then
        while IFS= read -r l; do
          echo "  - ${rel}: ${l}" >&2
        done <<<"$lines"
      else
        echo "  - ${rel}: (no MimeType= entry)" >&2
      fi
    done
    return 1
  fi

  if [ "$has_xlsx_mime" -ne 1 ]; then
    warn "No extracted .desktop file explicitly listed xlsx MIME '${REQUIRED_XLSX_MIME}'. Spreadsheet MIME types were present, but .xlsx double-click integration may be incomplete."
  fi

  # Prefer strict validation against tauri.conf.json so we catch missing MIME types,
  # scheme handlers, compliance artifacts, and Parquet shared-mime-info wiring in the
  # built artifact (not just in config).
  if command -v python3 >/dev/null 2>&1; then
    note "Static desktop integration validation (verify extracted RPM payload)"
    if ! python3 "$REPO_ROOT/scripts/ci/verify_linux_desktop_integration.py" \
      --package-root "$tmpdir" \
      --tauri-config "$TAURI_CONF" \
      --expected-main-binary "$EXPECTED_MAIN_BINARY" \
      --doc-package-name "$EXPECTED_RPM_NAME"; then
      err "Linux desktop integration verification failed for extracted RPM payload. See output above for expected vs observed MIME types."
      return 1
    fi
  fi
}

find_rpms() {
  local -a rpms=()

  if [[ -n "${RPM_OVERRIDE}" ]]; then
    # Resolve relative paths against the invocation directory, not the repo root.
    if [[ "${RPM_OVERRIDE}" != /* ]]; then
      RPM_OVERRIDE="${ORIG_PWD}/${RPM_OVERRIDE}"
    fi
    if [[ -d "${RPM_OVERRIDE}" ]]; then
      # Accept a directory to make local usage convenient.
      while IFS= read -r -d '' f; do rpms+=("$(abs_path "$f")"); done < <(find "${RPM_OVERRIDE}" -maxdepth 1 -type f -name '*.rpm' -print0)
    else
      rpms+=("$(abs_path "${RPM_OVERRIDE}")")
    fi
  else
    # Prefer predictable bundle globs (fast), but fall back to `find` for odd layouts.
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

    # Canonicalize and de-dupe roots (avoid duplicate scanning when CARGO_TARGET_DIR overlaps defaults).
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
      rpms+=("${root}/release/bundle/rpm/"*.rpm)
      rpms+=("${root}/"*/release/bundle/rpm/*.rpm)
    done
    if [[ "${nullglob_was_set}" -eq 0 ]]; then
      shopt -u nullglob
    fi

    if [[ ${#rpms[@]} -eq 0 ]]; then
      # Fallback: traverse the expected roots to locate RPM bundles.
      for root in "${roots[@]}"; do
        # Avoid an unbounded scan of the Cargo target directory: the bundle layout we care about is
        # shallow (<target>/<triple>/release/bundle/rpm/*.rpm). Keep a conservative maxdepth so
        # unexpected layouts still work, but we don't traverse the full build tree on failures.
        while IFS= read -r -d '' f; do rpms+=("$(abs_path "$f")"); done < <(find "${root}" -maxdepth 6 -type f -path '*/release/bundle/rpm/*.rpm' -print0 2>/dev/null || true)
      done
    fi
  fi

  # Deduplicate and sort for stable output.
  if [[ ${#rpms[@]} -eq 0 ]]; then
    return 0
  fi

  local -A seen=()
  local -a unique=()
  for rpm_path in "${rpms[@]}"; do
    seen["$rpm_path"]=1
  done
  for rpm_path in "${!seen[@]}"; do
    unique+=("$rpm_path")
  done

  IFS=$'\n' unique=($(printf '%s\n' "${unique[@]}" | sort))
  unset IFS

  printf '%s\n' "${unique[@]}"
}

validate_static() {
  local rpm_path="$1"

  [[ -f "${rpm_path}" ]] || die "RPM not found: ${rpm_path}"
  [[ -s "${rpm_path}" ]] || die "RPM is empty: ${rpm_path}"

  note "Static validation: ${rpm_path}"

  rpm -qp --info "${rpm_path}" >/dev/null || die "rpm --info query failed for: ${rpm_path}"

  local requires
  requires="$(rpm -qpR "${rpm_path}")" || die "rpm -qpR query failed for: ${rpm_path}"

  assert_requires_any() {
    local label="$1"
    shift
    local matched=0
    local needle
    for needle in "$@"; do
      # Avoid `printf ... | grep -q ...` under `set -o pipefail`: `grep -q` may exit early,
      # causing the upstream `printf` to receive SIGPIPE and return a non-zero status. That
      # would make the pipeline look like "no match" and can flake under load.
      if grep -Eqi "${needle}" <<<"${requires}"; then
        matched=1
        break
      fi
    done
    if [[ "${matched}" -ne 1 ]]; then
      die "RPM metadata missing required dependency (${label}); expected one of: $*"
    fi
  }

  assert_requires_rich_or() {
    local label="$1"
    local left_re="$2"
    local right_re="$3"
    # The RPM metadata should express these dependencies as a single "rich dependency"
    # OR expression (e.g. `(gtk3 or libgtk-3-0)`), not as two independent requirements.
    #
    # Avoid `\b` word-boundary regex constructs here: they're not portable across grep
    # implementations (BSD vs GNU). Instead, treat any non-word character as a boundary.
    local word_boundary_or_re
    word_boundary_or_re='(^|[^[:alnum:]_])or([^[:alnum:]_]|$)'

    local matched=0
    local line
    while IFS= read -r line; do
      # Use here-strings (not pipes) for the same SIGPIPE/pipefail reason as `assert_requires_any`.
      if grep -Eqi "${left_re}" <<<"$line" &&
        grep -Eqi "${right_re}" <<<"$line" &&
        grep -Eqi "${word_boundary_or_re}" <<<"$line"
      then
        matched=1
        break
      fi
    done <<<"${requires}"

    if [[ "${matched}" -ne 1 ]]; then
      die "RPM metadata missing rich dependency OR expression (${label}); expected a line containing both '${left_re}' and '${right_re}' joined by 'or'"
    fi
  }

  # Runtime dependencies (match `scripts/ci/verify-linux-package-deps.sh`).
  assert_requires_any "shared-mime-info (MIME database integration)" "shared-mime-info"
  assert_requires_rich_or "WebKitGTK 4.1 (Fedora/RHEL vs openSUSE)" "webkit2gtk4\\.1" "libwebkit2gtk-4_1"
  assert_requires_rich_or "GTK3 (Fedora/RHEL vs openSUSE)" "gtk3" "libgtk-3-0"
  assert_requires_rich_or \
    "AppIndicator/Ayatana (Fedora/RHEL vs openSUSE)" \
    "(libayatana-appindicator-gtk3|libappindicator-gtk3)" \
    "(libayatana-appindicator3-1|libappindicator3-1)"
  assert_requires_rich_or "librsvg (Fedora/RHEL vs openSUSE)" "librsvg2" "librsvg-2-2"
  assert_requires_rich_or "OpenSSL (Fedora/RHEL vs openSUSE)" "openssl-libs" "libopenssl3"

  local rpm_version
  local rpm_version_out
  set +e
  rpm_version_out="$(rpm -qp --queryformat '%{VERSION}\n' "${rpm_path}" 2>&1)"
  local rpm_version_status=$?
  set -e
  if [[ "$rpm_version_status" -ne 0 ]]; then
    die "rpm query failed for %{VERSION} on ${rpm_path}: ${rpm_version_out}"
  fi
  rpm_version="$(printf '%s' "$rpm_version_out" | head -n 1 | tr -d '\r')"
  if [[ -z "$rpm_version" ]]; then
    die "Failed to read RPM %{VERSION} from: ${rpm_path}"
  fi
  if [[ "$rpm_version" != "$EXPECTED_VERSION" ]]; then
    die "RPM version mismatch for ${rpm_path}: expected ${EXPECTED_VERSION}, found ${rpm_version}"
  fi

  local rpm_name
  local rpm_name_out
  set +e
  rpm_name_out="$(rpm -qp --queryformat '%{NAME}\n' "${rpm_path}" 2>&1)"
  local rpm_name_status=$?
  set -e
  if [[ "$rpm_name_status" -ne 0 ]]; then
    die "rpm query failed for %{NAME} on ${rpm_path}: ${rpm_name_out}"
  fi
  rpm_name="$(printf '%s' "$rpm_name_out" | head -n 1 | tr -d '\r')"
  if [[ -z "$rpm_name" ]]; then
    die "Failed to read RPM %{NAME} from: ${rpm_path}"
  fi
  if [[ "$rpm_name" != "$EXPECTED_RPM_NAME" ]]; then
    die "RPM name mismatch for ${rpm_path}: expected ${EXPECTED_RPM_NAME}, found ${rpm_name}"
  fi

  local file_list
  file_list="$(rpm -qp --list "${rpm_path}")" || die "rpm --list query failed for: ${rpm_path}"
  local static_validation_failed=0

  local expected_binary_path="/usr/bin/${EXPECTED_MAIN_BINARY}"
  if ! grep -qx "${expected_binary_path}" <<<"${file_list}"; then
    err "RPM payload missing expected desktop binary path: ${expected_binary_path}"
    err "First 200 lines of rpm file list:"
    echo "${file_list}" | head -n 200 >&2
    exit 1
  fi

  if ! grep -E -q '^/usr/share/applications/[^/]+\.desktop$' <<<"${file_list}"; then
    err "RPM payload missing expected .desktop file under: /usr/share/applications/"
    err "First 200 lines of rpm file list:"
    echo "${file_list}" | head -n 200 >&2
    exit 1
  fi

  # OSS/compliance artifacts should ship with the installed app.
  for filename in LICENSE NOTICE; do
    if ! grep -qx "/usr/share/doc/${EXPECTED_RPM_NAME}/${filename}" <<<"${file_list}"; then
      err "RPM payload missing compliance file: /usr/share/doc/${EXPECTED_RPM_NAME}/${filename}"
      err "First 200 lines of rpm file list:"
      echo "${file_list}" | head -n 200 >&2
      exit 1
    fi
  done

  # We ship a shared-mime-info definition for Parquet so `*.parquet` resolves to our
  # advertised MIME type (`application/vnd.apache.parquet`) on distros that don't
  # include it by default.
  if ! grep -qx "${EXPECTED_MIME_DEFINITION_PATH}" <<<"${file_list}"; then
    err "RPM payload missing Parquet shared-mime-info definition: ${EXPECTED_MIME_DEFINITION_PATH}"
    err "First 200 lines of rpm file list:"
    echo "${file_list}" | head -n 200 >&2
    static_validation_failed=1
  fi

  # If we are skipping the container install step, still validate the `.desktop` file
  # advertises configured file association + deep link MIME types by extracting the payload.
  if [[ "${NO_CONTAINER}" -eq 1 ]]; then
    if ! validate_desktop_mime_associations_extracted "${rpm_path}"; then
      static_validation_failed=1
    fi
  fi

  if [[ "${static_validation_failed}" -ne 0 ]]; then
    exit 1
  fi
}

validate_container() {
  local rpm_path="$1"

  require_docker

  rpm_path="$(abs_path "${rpm_path}")"
  local rpm_basename
  rpm_basename="$(basename "${rpm_path}")"

  local -a docker_platform_args=()
  if [[ -n "${DOCKER_PLATFORM}" ]]; then
    docker_platform_args=(--platform "${DOCKER_PLATFORM}")
  fi

  # Mount a temp directory that contains only the RPM under test so we don't
  # accidentally install multiple unrelated RPMs if the output directory has
  # leftovers from previous builds.
  local mount_dir
  mount_dir="$(mktemp -d)"
  # Note: do NOT symlink here. A symlink to a host path outside the bind mount
  # will be broken inside the container. Prefer a hardlink when possible to
  # avoid copying large RPM payloads; fall back to a normal copy otherwise.
  if ! ln "${rpm_path}" "${mount_dir}/${rpm_basename}" 2>/dev/null; then
    cp "${rpm_path}" "${mount_dir}/${rpm_basename}"
  fi

  note "Container validation (Fedora): ${rpm_path}"
  note "Using image: ${FEDORA_IMAGE}"

  local container_cmd
  container_cmd=$'set -euo pipefail\n'
  container_cmd+=$'echo "Fedora: $(cat /etc/fedora-release 2>/dev/null || true)"\n'
  # We generally do not GPG-sign the built RPM in CI; use --nogpgcheck so this
  # validates runtime deps/installability rather than signature policy.
  container_cmd+=$'dnf -y install --nogpgcheck --setopt=install_weak_deps=False /rpms/*.rpm\n'
  container_cmd+=$'expected_binary="'"${EXPECTED_MAIN_BINARY}"$'"\n'
  container_cmd+=$'binary_path="/usr/bin/${expected_binary}"\n'
  container_cmd+=$'test -x "${binary_path}"\n'
  container_cmd+=$'test -f /usr/share/doc/'"${EXPECTED_RPM_NAME}"$'/LICENSE\n'
  container_cmd+=$'test -f /usr/share/doc/'"${EXPECTED_RPM_NAME}"$'/NOTICE\n'
  container_cmd+=$'test -f '"${EXPECTED_MIME_DEFINITION_PATH}"$'\n'
  container_cmd+=$'grep -F "application/vnd.apache.parquet" '"${EXPECTED_MIME_DEFINITION_PATH}"$'\n'
  container_cmd+=$'grep -F "*.parquet" '"${EXPECTED_MIME_DEFINITION_PATH}"$'\n'
  container_cmd+=$'grep -Eq "application/vnd\\.apache\\.parquet:.*\\*\\.parquet" /usr/share/mime/globs2\n'
  container_cmd+=$'\n'
  container_cmd+=$'# Validate file association metadata is present in the installed .desktop entry.\n'
  container_cmd+=$'required_xlsx_mime="'"${REQUIRED_XLSX_MIME}"$'"\n'
  container_cmd+=$'spreadsheet_mime_regex="'"${SPREADSHEET_MIME_REGEX}"$'"\n'
  container_cmd+=$'required_parquet_mime="application/vnd.apache.parquet"\n'
  container_cmd+=$'expected_scheme_mimes_raw="${FORMULA_EXPECTED_SCHEME_MIMES:-}"\n'
  container_cmd+=$'if [ -z "${expected_scheme_mimes_raw}" ]; then\n'
  container_cmd+=$'  default_scheme="formula"\n'
  container_cmd+=$'  expected_scheme_mimes_raw="x-scheme-handler/${default_scheme}"\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'IFS=";" read -r -a required_scheme_mimes <<< "${expected_scheme_mimes_raw}"\n'
  container_cmd+=$'# Normalize scheme MIME tokens (trim + lowercase) so comparisons are token-exact.\n'
  container_cmd+=$'for i in "${!required_scheme_mimes[@]}"; do\n'
  container_cmd+=$'  t="${required_scheme_mimes[$i]}"\n'
  container_cmd+=$'  t="${t#"${t%%[![:space:]]*}"}"\n'
  container_cmd+=$'  t="${t%"${t##*[![:space:]]}"}"\n'
  container_cmd+=$'  t="$(printf "%s" "${t}" | tr "[:upper:]" "[:lower:]")"\n'
  container_cmd+=$'  required_scheme_mimes[$i]="${t}"\n'
  container_cmd+=$'done\n'
  container_cmd+=$'desktop_files=(/usr/share/applications/*.desktop)\n'
  container_cmd+=$'if [ ! -e "${desktop_files[0]}" ]; then\n'
  container_cmd+=$'  echo "No .desktop files found under /usr/share/applications after RPM install." >&2\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'\n'
  container_cmd+=$'# Filter to the .desktop entries that actually launch our app (Exec= references the expected binary).\n'
  container_cmd+=$'matching_desktop_files=()\n'
  container_cmd+=$'for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'  exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'  exec_value="$(printf "%s" "${exec_line}" | sed -E "s/^[[:space:]]*Exec[[:space:]]*=[[:space:]]*//")"\n'
  container_cmd+=$'  # Normalize quoted Exec tokens (e.g. Exec="/usr/bin/formula-desktop" %U)\n'
  container_cmd+=$'  exec_value="${exec_value//\\\"/}"\n'
  container_cmd+=$'  exec_value="${exec_value//\\\'/}"\n'
  container_cmd+=$'  spaced=" ${exec_value} "\n'
  container_cmd+=$'  if [ -n "${exec_value}" ] && (printf "%s" "${spaced}" | grep -Fq " ${expected_binary} " || printf "%s" "${spaced}" | grep -Fq " ${binary_path} "); then\n'
  container_cmd+=$'    matching_desktop_files+=("$f")\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'done\n'
  container_cmd+=$'if [ "${#matching_desktop_files[@]}" -eq 0 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop files appear to target the expected executable ${expected_binary} (or ${binary_path}) in their Exec= entry." >&2\n'
  container_cmd+=$'  echo "Installed .desktop files inspected:" >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'    if [ -z "${exec_line}" ]; then\n'
  container_cmd+=$'      exec_line="(no Exec= entry)"\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'    echo "  - ${f}: ${exec_line}" >&2\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'desktop_files=("${matching_desktop_files[@]}")\n'
  container_cmd+=$'has_any_mimetype=0\n'
  container_cmd+=$'has_spreadsheet_mime=0\n'
  container_cmd+=$'has_xlsx_mime=0\n'
  container_cmd+=$'has_xlsx_integration=0\n'
  container_cmd+=$'has_parquet_mime=0\n'
  container_cmd+=$'declare -A found_scheme_mimes=()\n'
  container_cmd+=$'bad_exec_count=0\n'
  container_cmd+=$'for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'  mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'  if [ -z "${mime_line}" ]; then\n'
  container_cmd+=$'    continue\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  has_any_mimetype=1\n'
  container_cmd+=$'  mime_value="$(printf "%s" "${mime_line}" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//")"\n'
  container_cmd+=$'  IFS=";" read -r -a mime_tokens <<< "${mime_value}"\n'
  container_cmd+=$'  declare -A mime_set=()\n'
  container_cmd+=$'  for token in "${mime_tokens[@]}"; do\n'
  container_cmd+=$'    token="${token#"${token%%[![:space:]]*}"}"\n'
  container_cmd+=$'    token="${token%"${token##*[![:space:]]}"}"\n'
  container_cmd+=$'    token="$(printf "%s" "${token}" | tr "[:upper:]" "[:lower:]")"\n'
  container_cmd+=$'    if [ -n "${token}" ]; then\n'
  container_cmd+=$'      mime_set["$token"]=1\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Fqi "${required_parquet_mime}"; then\n'
  container_cmd+=$'    has_parquet_mime=1\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Fqi "${required_xlsx_mime}"; then\n'
  container_cmd+=$'    has_xlsx_mime=1\n'
  container_cmd+=$'    has_xlsx_integration=1\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Fqi "xlsx"; then\n'
  container_cmd+=$'    has_xlsx_integration=1\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Eqi "${spreadsheet_mime_regex}"; then\n'
  container_cmd+=$'    has_spreadsheet_mime=1\n'
  container_cmd+=$'    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'    if [ -z "${exec_line}" ]; then\n'
  container_cmd+=$'      echo "Installed desktop entry ${f} is missing an Exec= entry (required for file associations)" >&2\n'
  container_cmd+=$'      bad_exec_count=$((bad_exec_count + 1))\n'
  container_cmd+=$'    elif ! printf "%s" "${exec_line}" | grep -Eq "%[uUfF]"; then\n'
  container_cmd+=$'      echo "Installed desktop entry ${f} Exec= does not include a file/URL placeholder (%U/%u/%F/%f): ${exec_line}" >&2\n'
  container_cmd+=$'      bad_exec_count=$((bad_exec_count + 1))\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  has_any_expected_scheme_in_file=0\n'
  container_cmd+=$'  for scheme_mime in "${required_scheme_mimes[@]}"; do\n'
  container_cmd+=$'    if [ -n "${scheme_mime}" ] && [ -n "${mime_set["$scheme_mime"]+x}" ]; then\n'
  container_cmd+=$'      found_scheme_mimes["$scheme_mime"]=1\n'
  container_cmd+=$'      has_any_expected_scheme_in_file=1\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  if [ "${has_any_expected_scheme_in_file}" -eq 1 ]; then\n'
  container_cmd+=$'    exec_line="$(grep -Ei "^[[:space:]]*Exec[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'    if [ -z "${exec_line}" ]; then\n'
  container_cmd+=$'      echo "Installed desktop entry ${f} is missing an Exec= entry (required for URL scheme handlers)" >&2\n'
  container_cmd+=$'      bad_exec_count=$((bad_exec_count + 1))\n'
  container_cmd+=$'    elif ! printf "%s" "${exec_line}" | grep -Eq "%[uUfF]"; then\n'
  container_cmd+=$'      echo "Installed desktop entry ${f} Exec= does not include a URL placeholder (%U/%u/%F/%f): ${exec_line}" >&2\n'
  container_cmd+=$'      bad_exec_count=$((bad_exec_count + 1))\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'done\n'
  container_cmd+=$'if [ "${has_any_mimetype}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop file contained a MimeType= entry (file associations missing)." >&2\n'
  container_cmd+=$'  echo "Expected MimeType to include spreadsheet MIME types (tauri.conf.json bundle.fileAssociations)." >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    echo "  - ${f}" >&2\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${has_xlsx_integration}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop MimeType= entry advertised xlsx support (file associations missing)." >&2\n'
  container_cmd+=$'  echo "Expected MimeType= to include substring xlsx or MIME ${required_xlsx_mime}." >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" || true)"\n'
  container_cmd+=$'    if [ -n "${lines}" ]; then\n'
  container_cmd+=$'      while IFS= read -r l; do echo "  - ${f}: ${l}" >&2; done <<<"${lines}"\n'
  container_cmd+=$'    else\n'
  container_cmd+=$'      echo "  - ${f}: (no MimeType= entry)" >&2\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${has_parquet_mime}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop MimeType= entry advertised Parquet support (file associations missing)." >&2\n'
  container_cmd+=$'  echo "Expected MimeType= to include ${required_parquet_mime}." >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" || true)"\n'
  container_cmd+=$'    if [ -n "${lines}" ]; then\n'
  container_cmd+=$'      while IFS= read -r l; do echo "  - ${f}: ${l}" >&2; done <<<"${lines}"\n'
  container_cmd+=$'    else\n'
  container_cmd+=$'      echo "  - ${f}: (no MimeType= entry)" >&2\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${has_spreadsheet_mime}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop MimeType= entry advertised spreadsheet/xlsx support (file associations missing)." >&2\n'
  container_cmd+=$'  echo "Expected MimeType= to include ${required_xlsx_mime} (xlsx) or another spreadsheet MIME type." >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" || true)"\n'
  container_cmd+=$'    if [ -n "${lines}" ]; then\n'
  container_cmd+=$'      while IFS= read -r l; do echo "  - ${f}: ${l}" >&2; done <<<"${lines}"\n'
  container_cmd+=$'    else\n'
  container_cmd+=$'      echo "  - ${f}: (no MimeType= entry)" >&2\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'missing_scheme_mimes=()\n'
  container_cmd+=$'for scheme_mime in "${required_scheme_mimes[@]}"; do\n'
  container_cmd+=$'  if [ -n "${scheme_mime}" ] && [ -z "${found_scheme_mimes["$scheme_mime"]+x}" ]; then\n'
  container_cmd+=$'    missing_scheme_mimes+=("$scheme_mime")\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'done\n'
  container_cmd+=$'if [ "${#missing_scheme_mimes[@]}" -ne 0 ]; then\n'
  container_cmd+=$'  echo "No installed .desktop MimeType= entry advertised the expected URL scheme handler(s): ${missing_scheme_mimes[*]}." >&2\n'
  container_cmd+=$'  echo "Expected MimeType= to include: ${required_scheme_mimes[*]}" >&2\n'
  container_cmd+=$'  for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'    lines="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" || true)"\n'
  container_cmd+=$'    if [ -n "${lines}" ]; then\n'
  container_cmd+=$'      while IFS= read -r l; do echo "  - ${f}: ${l}" >&2; done <<<"${lines}"\n'
  container_cmd+=$'    else\n'
  container_cmd+=$'      echo "  - ${f}: (no MimeType= entry)" >&2\n'
  container_cmd+=$'    fi\n'
  container_cmd+=$'  done\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${bad_exec_count}" -ne 0 ]; then\n'
  container_cmd+=$'  echo "One or more installed .desktop entries had invalid Exec= lines for file association handling." >&2\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${has_xlsx_mime}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "WARN: No installed .desktop file explicitly listed xlsx MIME ${required_xlsx_mime}." >&2\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'set +e\n'
  container_cmd+=$'ldd_out="$(ldd "${binary_path}" 2>&1)"\n'
  container_cmd+=$'ldd_status=$?\n'
  container_cmd+=$'set -e\n'
  container_cmd+=$'echo "${ldd_out}"\n'
  container_cmd+=$'if echo "${ldd_out}" | grep -q "not found"; then\n'
  container_cmd+=$'  echo "Missing shared library dependencies detected:" >&2\n'
  container_cmd+=$'  echo "${ldd_out}" | grep "not found" >&2 || true\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'# ldd returns non-zero for static binaries ("not a dynamic executable"). Treat that as OK\n'
  container_cmd+=$'# as long as we did not detect missing shared libraries.\n'
  container_cmd+=$'if [ "${ldd_status}" -ne 0 ] && ! echo "${ldd_out}" | grep -q "not a dynamic executable" && ! echo "${ldd_out}" | grep -q "statically linked"; then\n'
  container_cmd+=$'  echo "ldd exited with status ${ldd_status}" >&2\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'

  set +e
  docker pull "${docker_platform_args[@]}" "${FEDORA_IMAGE}"
  docker run --rm \
    "${docker_platform_args[@]}" \
    -e "FORMULA_EXPECTED_SCHEME_MIMES=${EXPECTED_SCHEME_MIMES_ENV}" \
    -v "${mount_dir}:/rpms:ro" \
    "${FEDORA_IMAGE}" \
    bash -lc "${container_cmd}"
  local status=$?
  set -e
  rm -rf "${mount_dir}"
  if [[ "${status}" -ne 0 ]]; then
    die "Fedora container installability check failed for: ${rpm_path}"
  fi
}

main() {
  mapfile -t rpms < <(find_rpms)

  if [[ ${#rpms[@]} -eq 0 ]]; then
    err "No RPM files found."
    err "Searched (repo-relative):"
    if [[ -n "${CARGO_TARGET_DIR:-}" ]]; then
      err "  - \$CARGO_TARGET_DIR/release/bundle/rpm/*.rpm (\$CARGO_TARGET_DIR=${CARGO_TARGET_DIR})"
      err "  - \$CARGO_TARGET_DIR/*/release/bundle/rpm/*.rpm (\$CARGO_TARGET_DIR=${CARGO_TARGET_DIR})"
    fi
    err "  - apps/desktop/src-tauri/target/release/bundle/rpm/*.rpm"
    err "  - apps/desktop/src-tauri/target/*/release/bundle/rpm/*.rpm"
    err "  - apps/desktop/target/release/bundle/rpm/*.rpm"
    err "  - apps/desktop/target/*/release/bundle/rpm/*.rpm"
    err "  - target/release/bundle/rpm/*.rpm"
    err "  - target/*/release/bundle/rpm/*.rpm"
    err "Tip: Use --rpm <path> to specify one explicitly."
    exit 1
  fi

  note "Found ${#rpms[@]} RPM(s) to validate:"
  for rpm_path in "${rpms[@]}"; do
    echo "  - $(rel_path "${rpm_path}")"
  done
  for rpm_path in "${rpms[@]}"; do
    validate_static "${rpm_path}"
    if [[ "${NO_CONTAINER}" -eq 0 ]]; then
      validate_container "${rpm_path}"
    else
      note "Skipping container validation (--no-container)"
    fi
  done

  note "RPM validation passed"
}

main "$@"
