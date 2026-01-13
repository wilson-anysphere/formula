#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat <<'EOF'
usage: scripts/ci/linux-package-install-smoke.sh [deb|rpm|all]

Installs the built Linux desktop packages into clean containers and validates that
the installed binary has no missing shared libraries.

This is intended for CI/release workflows to catch:
  - missing runtime dependencies in packaging metadata
  - broken postinstall scripts
  - missing shared libraries (ldd "not found")

Environment variables:
  FORMULA_DEB_SMOKE_IMAGE  Ubuntu image to use (default: ubuntu:24.04)
  FORMULA_RPM_SMOKE_IMAGE  Fedora image to use (default: fedora:40)
EOF
}

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "${repo_root}"

kind="${1:-all}"
case "${kind}" in
  deb | rpm | all) ;;
  -h | --help)
    usage
    exit 0
    ;;
  *)
    usage >&2
    exit 2
    ;;
esac

require_docker() {
  if ! command -v docker >/dev/null 2>&1; then
    echo "linux-package-install-smoke: docker is required but was not found on PATH" >&2
    exit 1
  fi
  echo "::group::Docker version"
  docker version
  echo "::endgroup::"
}

find_pkg_dirs() {
  local pkg_type="$1" # deb|rpm
  local ext="$2" # .deb|.rpm
  local out_var="$3" # name of array var to populate
  # shellcheck disable=SC2178 # (used via nameref)
  local -n out_dirs="$out_var"

  local -a files
  mapfile -t files < <(
    find . -type f \( \
      -path "*/target/release/bundle/${pkg_type}/*${ext}" -o \
      -path "*/target/*/release/bundle/${pkg_type}/*${ext}" \
    \) -print | sort
  )

  if [[ ${#files[@]} -eq 0 ]]; then
    echo "linux-package-install-smoke: no ${ext} artifacts found under target/**/release/bundle/${pkg_type}" >&2
    exit 1
  fi

  echo "::group::Discovered ${ext} artifacts"
  printf '%s\n' "${files[@]}"
  echo "::endgroup::"

  local -A seen=()
  local -a dirs=()
  local f d
  for f in "${files[@]}"; do
    d="$(dirname "${f}")"
    if [[ -n "${seen[${d}]:-}" ]]; then
      continue
    fi
    seen["${d}"]=1
    dirs+=("${d}")
  done

  out_dirs=("${dirs[@]}")
}

deb_smoke_test_dir() {
  local deb_dir="$1"
  local image="${FORMULA_DEB_SMOKE_IMAGE:-ubuntu:24.04}"

  local deb_dir_abs
  deb_dir_abs="$(cd "${deb_dir}" && pwd)"

  echo "::group::.deb smoke test (${image}) - ${deb_dir}"
  echo "Mounting: ${deb_dir_abs} -> /mounted"
  docker pull "${image}"
  docker run --rm \
    -v "${deb_dir_abs}:/mounted:ro" \
    "${image}" \
    bash -euxo pipefail -c '
      export DEBIAN_FRONTEND=noninteractive
      echo "Container OS:"; cat /etc/os-release
      echo "Mounted artifacts:"; ls -lah /mounted
      apt-get update
      # Use --no-install-recommends to ensure runtime requirements are expressed as Depends
      apt-get install -y --no-install-recommends /mounted/*.deb

      test -x /usr/bin/formula-desktop
      set +e
      out="$(ldd /usr/bin/formula-desktop 2>&1)"
      status=$?
      set -e
      echo "ldd /usr/bin/formula-desktop:"
      echo "${out}"
      if echo "${out}" | grep -q "not found"; then
        echo "Missing shared libraries:"
        echo "${out}" | grep "not found" || true
        exit 1
      fi
      if [ "${status}" -ne 0 ]; then
        echo "ldd exited with status ${status}" >&2
        exit 1
      fi
    '
  echo "::endgroup::"
}

rpm_smoke_test_dir() {
  local rpm_dir="$1"
  local image="${FORMULA_RPM_SMOKE_IMAGE:-fedora:40}"

  local rpm_dir_abs
  rpm_dir_abs="$(cd "${rpm_dir}" && pwd)"

  echo "::group::.rpm smoke test (${image}) - ${rpm_dir}"
  echo "Mounting: ${rpm_dir_abs} -> /mounted"
  docker pull "${image}"
  docker run --rm \
    -v "${rpm_dir_abs}:/mounted:ro" \
    "${image}" \
    bash -euxo pipefail -c '
      echo "Container OS:"; cat /etc/os-release
      echo "Mounted artifacts:"; ls -lah /mounted
      # Ensure metadata is available, then install the local RPM.
      dnf -y install --setopt=install_weak_deps=False /mounted/*.rpm

      test -x /usr/bin/formula-desktop
      set +e
      out="$(ldd /usr/bin/formula-desktop 2>&1)"
      status=$?
      set -e
      echo "ldd /usr/bin/formula-desktop:"
      echo "${out}"
      if echo "${out}" | grep -q "not found"; then
        echo "Missing shared libraries:"
        echo "${out}" | grep "not found" || true
        exit 1
      fi
      if [ "${status}" -ne 0 ]; then
        echo "ldd exited with status ${status}" >&2
        exit 1
      fi
    '
  echo "::endgroup::"
}

require_docker

if [[ "${kind}" == "deb" || "${kind}" == "all" ]]; then
  deb_dirs=()
  find_pkg_dirs "deb" ".deb" deb_dirs
  for d in "${deb_dirs[@]}"; do
    deb_smoke_test_dir "${d}"
  done
fi

if [[ "${kind}" == "rpm" || "${kind}" == "all" ]]; then
  rpm_dirs=()
  find_pkg_dirs "rpm" ".rpm" rpm_dirs
  for d in "${rpm_dirs[@]}"; do
    rpm_smoke_test_dir "${d}"
  done
fi
