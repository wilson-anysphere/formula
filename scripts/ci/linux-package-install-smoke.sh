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
  CARGO_TARGET_DIR         Optional Cargo target directory override (also used for artifact discovery)
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

  # Avoid `find .` here: release jobs run `pnpm install`, so traversing the repo
  # (node_modules, pnpm store, etc) can be very slow. Only scan expected Tauri
  # target directories.
  local -a search_roots=()

  # Respect an explicit Cargo target dir override (common in CI caching setups).
  local cargo_target_dir="${CARGO_TARGET_DIR:-}"
  if [[ -n "${cargo_target_dir}" ]]; then
    # Cargo interprets relative paths relative to the working directory (repo root).
    if [[ "${cargo_target_dir}" != /* ]]; then
      cargo_target_dir="${repo_root}/${cargo_target_dir}"
    fi
    if [[ -d "${cargo_target_dir}" ]]; then
      search_roots+=("${cargo_target_dir}")
    fi
  fi
  local root
  for root in "apps/desktop/src-tauri/target" "apps/desktop/target" "target"; do
    if [[ -d "$root" ]]; then
      search_roots+=("$root")
    fi
  done

  if [[ ${#search_roots[@]} -eq 0 ]]; then
    echo "linux-package-install-smoke: no target directories found (expected CARGO_TARGET_DIR, apps/desktop/src-tauri/target, apps/desktop/target, or target)" >&2
    exit 1
  fi

  local -a files
  # Use globs instead of `find` to avoid traversing the entire Cargo target directory,
  # which can contain many build artifacts. The bundle output layout is predictable:
  #   <target>/release/bundle/<type>/*.deb|*.rpm
  #   <target>/<triple>/release/bundle/<type>/*.deb|*.rpm
  local nullglob_was_set=0
  if shopt -q nullglob; then
    nullglob_was_set=1
  fi
  shopt -s nullglob
  files=()
  for root in "${search_roots[@]}"; do
    files+=("$root"/release/bundle/"${pkg_type}"/*"${ext}")
    files+=("$root"/*/release/bundle/"${pkg_type}"/*"${ext}")
  done
  if [[ "${nullglob_was_set}" -eq 0 ]]; then
    shopt -u nullglob
  fi

  # Sort and de-dupe for deterministic logs.
  if [[ ${#files[@]} -gt 0 ]]; then
    mapfile -t files < <(printf '%s\n' "${files[@]}" | sort -u)
  fi

  if [[ ${#files[@]} -eq 0 ]]; then
    echo "linux-package-install-smoke: no ${ext} artifacts found under any target root." >&2
    echo "Expected something like: <target>/**/release/bundle/${pkg_type}/*${ext}" >&2
    echo "Search roots:" >&2
    printf '  - %s\n' "${search_roots[@]}" >&2
    echo "::group::linux-package-install-smoke: debug listing (release/bundle dirs)"
    for root in "${search_roots[@]}"; do
      echo "==> $root"
      ls -lah "$root/release/bundle" 2>/dev/null || true
      # Also show target-triple layout (target/<triple>/release/bundle).
      ls -lah "$root"/*/release/bundle 2>/dev/null || true
    done
    echo "::endgroup::"
    exit 1
  fi

  echo "::group::Discovered ${ext} artifacts"
  printf '%s\n' "${files[@]}"
  echo "::endgroup::"

  local -A seen=()
  local -a dirs=()
  local f d d_abs
  for f in "${files[@]}"; do
    d="$(dirname "${f}")"
    # Normalize to an absolute path so we don't accidentally run the same smoke
    # test twice (e.g. when `CARGO_TARGET_DIR=target` yields both absolute and
    # relative matches).
    d_abs="$(cd "${d}" && pwd -P)"
    if [[ -n "${seen[${d_abs}]:-}" ]]; then
      continue
    fi
    seen["${d_abs}"]=1
    dirs+=("${d_abs}")
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
      if [ "${status}" -ne 0 ] && ! echo "${out}" | grep -q "not a dynamic executable" && ! echo "${out}" | grep -q "statically linked"; then
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
      # The Tauri updater `.sig` files are not RPM GPG signatures, and we generally do not GPG-sign
      # the built RPM in CI. Use --nogpgcheck so the smoke test validates runtime deps rather than
      # failing on signing policy.
      dnf -y install --nogpgcheck --setopt=install_weak_deps=False /mounted/*.rpm

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
      if [ "${status}" -ne 0 ] && ! echo "${out}" | grep -q "not a dynamic executable" && ! echo "${out}" | grep -q "statically linked"; then
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
