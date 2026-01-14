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
${SCRIPT_NAME} - validate a formula-desktop RPM bundle

Usage:
  ${SCRIPT_NAME} [--rpm <path>] [--no-container] [--image <fedora-image>]

Options:
  --rpm <path>        Validate a specific .rpm (or a directory containing .rpm files).
                      If omitted, the script searches common Tauri output locations:
                        - \$CARGO_TARGET_DIR/**/release/bundle/rpm/*.rpm (if set)
                        - apps/desktop/src-tauri/target/**/release/bundle/rpm/*.rpm
                        - apps/desktop/target/**/release/bundle/rpm/*.rpm
                        - target/**/release/bundle/rpm/*.rpm
  --no-container      Skip the Fedora container installability check (static checks only).
  --image <image>     Fedora image to use for the container step (default: fedora:40).
  -h, --help          Show this help text.
EOF
}

err() {
  echo "${SCRIPT_NAME}: ERROR: $*" >&2
}

note() {
  echo "${SCRIPT_NAME}: $*"
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
}

require_cmd rpm

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
SPREADSHEET_MIME_REGEX='application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|application/vnd\.ms-excel|application/vnd\.ms-excel\.sheet\.macroEnabled\.12|application/vnd\.ms-excel\.sheet\.binary\.macroEnabled\.12|application/vnd\.openxmlformats-officedocument\.spreadsheetml\.template|application/vnd\.ms-excel\.template\.macroEnabled\.12|application/vnd\.ms-excel\.addin\.macroEnabled\.12|text/csv'

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
        while IFS= read -r -d '' f; do rpms+=("$(abs_path "$f")"); done < <(find "${root}" -type f -path '*/release/bundle/rpm/*.rpm' -print0 2>/dev/null || true)
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

  note "Static validation: ${rpm_path}"

  rpm -qp --info "${rpm_path}" >/dev/null || die "rpm --info query failed for: ${rpm_path}"

  local file_list
  file_list="$(rpm -qp --list "${rpm_path}")" || die "rpm --list query failed for: ${rpm_path}"

  if ! grep -qx '/usr/bin/formula-desktop' <<<"${file_list}"; then
    err "RPM payload missing expected desktop binary path: /usr/bin/formula-desktop"
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
}

validate_container() {
  local rpm_path="$1"

  require_docker

  rpm_path="$(abs_path "${rpm_path}")"
  local rpm_basename
  rpm_basename="$(basename "${rpm_path}")"

  # Mount a temp directory that contains only the RPM under test so we don't
  # accidentally install multiple unrelated RPMs if the output directory has
  # leftovers from previous builds.
  local mount_dir
  mount_dir="$(mktemp -d)"
  if ! ln -s "${rpm_path}" "${mount_dir}/${rpm_basename}" 2>/dev/null; then
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
  container_cmd+=$'test -x /usr/bin/formula-desktop\n'
  container_cmd+=$'\n'
  container_cmd+=$'# Validate file association metadata is present in the installed .desktop entry.\n'
  container_cmd+=$'required_xlsx_mime="'"${REQUIRED_XLSX_MIME}"$'"\n'
  container_cmd+=$'spreadsheet_mime_regex="'"${SPREADSHEET_MIME_REGEX}"$'"\n'
  container_cmd+=$'desktop_files=(/usr/share/applications/*.desktop)\n'
  container_cmd+=$'if [ ! -e "${desktop_files[0]}" ]; then\n'
  container_cmd+=$'  echo "No .desktop files found under /usr/share/applications after RPM install." >&2\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'has_any_mimetype=0\n'
  container_cmd+=$'has_spreadsheet_mime=0\n'
  container_cmd+=$'has_xlsx_mime=0\n'
  container_cmd+=$'for f in "${desktop_files[@]}"; do\n'
  container_cmd+=$'  mime_line="$(grep -Ei "^[[:space:]]*MimeType[[:space:]]*=" "$f" | head -n 1 || true)"\n'
  container_cmd+=$'  if [ -z "${mime_line}" ]; then\n'
  container_cmd+=$'    continue\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  has_any_mimetype=1\n'
  container_cmd+=$'  mime_value="$(printf "%s" "${mime_line}" | sed -E "s/^[[:space:]]*MimeType[[:space:]]*=[[:space:]]*//")"\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Fqi "${required_xlsx_mime}"; then\n'
  container_cmd+=$'    has_xlsx_mime=1\n'
  container_cmd+=$'  fi\n'
  container_cmd+=$'  if printf "%s" "${mime_value}" | grep -Eqi "${spreadsheet_mime_regex}"; then\n'
  container_cmd+=$'    has_spreadsheet_mime=1\n'
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
  container_cmd+=$'if [ "${has_xlsx_mime}" -ne 1 ]; then\n'
  container_cmd+=$'  echo "WARN: No installed .desktop file explicitly listed xlsx MIME ${required_xlsx_mime}." >&2\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'set +e\n'
  container_cmd+=$'ldd_out="$(ldd /usr/bin/formula-desktop 2>&1)"\n'
  container_cmd+=$'ldd_status=$?\n'
  container_cmd+=$'set -e\n'
  container_cmd+=$'echo "${ldd_out}"\n'
  container_cmd+=$'if echo "${ldd_out}" | grep -q "not found"; then\n'
  container_cmd+=$'  echo "Missing shared library dependencies detected:" >&2\n'
  container_cmd+=$'  echo "${ldd_out}" | grep "not found" >&2 || true\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'
  container_cmd+=$'if [ "${ldd_status}" -ne 0 ] && ! echo "${ldd_out}" | grep -q "not a dynamic executable" && ! echo "${ldd_out}" | grep -q "statically linked"; then\n'
  container_cmd+=$'  echo "ldd exited with status ${ldd_status}" >&2\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'

  set +e
  docker run --rm -v "${mount_dir}:/rpms:ro" "${FEDORA_IMAGE}" bash -lc "${container_cmd}"
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
      err "  - \$CARGO_TARGET_DIR/**/release/bundle/rpm/*.rpm (\$CARGO_TARGET_DIR=${CARGO_TARGET_DIR})"
    fi
    err "  - apps/desktop/src-tauri/target/**/release/bundle/rpm/*.rpm"
    err "  - apps/desktop/target/**/release/bundle/rpm/*.rpm"
    err "  - target/**/release/bundle/rpm/*.rpm"
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
