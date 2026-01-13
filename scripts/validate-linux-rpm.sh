#!/bin/bash
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

usage() {
  cat <<EOF
${SCRIPT_NAME} - validate a formula-desktop RPM bundle

Usage:
  ${SCRIPT_NAME} [--rpm <path>] [--no-container] [--image <fedora-image>]

Options:
  --rpm <path>        Validate a specific .rpm (or a directory containing .rpm files).
                      If omitted, the script searches common Tauri output locations:
                        - apps/desktop/src-tauri/target/**/release/bundle/rpm/*.rpm
                        - target/**/release/bundle/rpm/*.rpm
  --no-container      Skip the Fedora container installability check (static checks only).
  --image <image>     Fedora image to use for the container step (default: fedora:40).
  -h, --help          Show this help text.
EOF
}

err() {
  echo "ERROR: $*" >&2
}

note() {
  echo "==> $*"
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

require_cmd rpm

find_rpms() {
  local -a rpms=()

  if [[ -n "${RPM_OVERRIDE}" ]]; then
    if [[ -d "${RPM_OVERRIDE}" ]]; then
      # Accept a directory to make local usage convenient.
      while IFS= read -r -d '' f; do rpms+=("$f"); done < <(find "${RPM_OVERRIDE}" -maxdepth 1 -type f -name '*.rpm' -print0)
    else
      rpms+=("${RPM_OVERRIDE}")
    fi
  else
    if [[ -d "apps/desktop/src-tauri/target" ]]; then
      while IFS= read -r -d '' f; do rpms+=("$f"); done < <(find "apps/desktop/src-tauri/target" -type f -path '*/release/bundle/rpm/*.rpm' -print0)
    fi
    if [[ -d "target" ]]; then
      while IFS= read -r -d '' f; do rpms+=("$f"); done < <(find "target" -type f -path '*/release/bundle/rpm/*.rpm' -print0)
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

  require_cmd docker

  # Docker bind mounts require an absolute path.
  local rpm_dir
  rpm_dir="$(cd "$(dirname "${rpm_path}")" && pwd -P)"

  note "Container validation (Fedora): ${rpm_path}"
  note "Using image: ${FEDORA_IMAGE}"

  local container_cmd
  container_cmd=$'set -euo pipefail\n'
  container_cmd+=$'echo "Fedora: $(cat /etc/fedora-release 2>/dev/null || true)"\n'
  container_cmd+=$'dnf -y install /rpms/*.rpm\n'
  container_cmd+=$'test -x /usr/bin/formula-desktop\n'
  # ldd returns non-zero for static binaries; treat that as OK as long as we
  # don\'t have any missing dynamic deps.
  container_cmd+=$'ldd_out="$(ldd /usr/bin/formula-desktop 2>&1 || true)"\n'
  container_cmd+=$'echo "${ldd_out}"\n'
  container_cmd+=$'if echo "${ldd_out}" | grep -q "not found"; then\n'
  container_cmd+=$'  echo "Missing shared library dependencies detected:" >&2\n'
  container_cmd+=$'  echo "${ldd_out}" | grep "not found" >&2 || true\n'
  container_cmd+=$'  exit 1\n'
  container_cmd+=$'fi\n'

  docker run --rm -v "${rpm_dir}:/rpms:ro" "${FEDORA_IMAGE}" bash -lc "${container_cmd}" \
    || die "Fedora container installability check failed for: ${rpm_path}"
}

main() {
  mapfile -t rpms < <(find_rpms)

  if [[ ${#rpms[@]} -eq 0 ]]; then
    die "No RPM files found. Use --rpm <path> to specify one explicitly."
  fi

  note "Found ${#rpms[@]} RPM(s) to validate"
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

