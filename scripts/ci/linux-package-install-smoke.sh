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
  DOCKER_PLATFORM          Optional docker --platform override (default: host architecture)
  FORMULA_DEB_SMOKE_IMAGE  Ubuntu image to use (default: ubuntu:24.04)
  FORMULA_RPM_SMOKE_IMAGE  RPM-based image to use (default: fedora:40; supports openSUSE too)
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
  if [[ -n "${DOCKER_PLATFORM:-}" ]]; then
    echo "Docker platform: ${DOCKER_PLATFORM}"
  fi
  echo "::endgroup::"
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

# Force docker to use the host architecture image variant by default so we don't accidentally run
# an ARM container on an x86_64 runner (or vice versa). This avoids confusing failures like
# `exec format error` or `ldd: not a dynamic executable` when the installed binary can't run on the
# container architecture.
#
# Override for debugging by exporting DOCKER_PLATFORM explicitly.
DOCKER_PLATFORM="${DOCKER_PLATFORM:-$(detect_docker_platform)}"

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

  # Canonicalize + de-dupe roots (avoid double-scanning when CARGO_TARGET_DIR overlaps defaults).
  local -A seen_roots=()
  local -a uniq_roots=()
  local root_abs
  for root in "${search_roots[@]}"; do
    root_abs="$(cd "$root" && pwd -P)"
    if [[ -n "${seen_roots[${root_abs}]:-}" ]]; then
      continue
    fi
    seen_roots["${root_abs}"]=1
    uniq_roots+=("${root_abs}")
  done
  search_roots=("${uniq_roots[@]}")

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
  local image="$2"

  local deb_dir_abs
  deb_dir_abs="$(cd "${deb_dir}" && pwd)"

  echo "::group::.deb smoke test (${image}) - ${deb_dir}"
  echo "Mounting: ${deb_dir_abs} -> /mounted"
  docker run --rm \
    ${DOCKER_PLATFORM:+--platform "${DOCKER_PLATFORM}"} \
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

      # Validate that installer-time MIME integration ran successfully.
      #
      # Many distros do not ship a Parquet glob by default, so the package ships a
      # shared-mime-info definition under /usr/share/mime/packages and relies on the
      # shared-mime-info triggers to rebuild /usr/share/mime/globs2.
      test -f /usr/share/mime/packages/app.formula.desktop.xml
      if ! grep -Eq "application/vnd\.apache\.parquet:.*\*\.parquet" /usr/share/mime/globs2; then
        echo "Missing Parquet MIME mapping in /usr/share/mime/globs2 (expected application/vnd.apache.parquet -> *.parquet)" >&2
        echo "Parquet-related globs2 lines:" >&2
        grep -n parquet /usr/share/mime/globs2 || true
        exit 1
      fi

      desktop_file="$(grep -rlE "^[[:space:]]*Exec=.*formula-desktop" /usr/share/applications 2>/dev/null | head -n 1 || true)"
      if [ -z "${desktop_file}" ]; then
        echo "No installed .desktop file found with Exec referencing formula-desktop under /usr/share/applications" >&2
        ls -lah /usr/share/applications || true
        exit 1
      fi
      echo "Installed desktop entry: ${desktop_file}"
      grep -E "^[[:space:]]*(Exec|MimeType)=" "${desktop_file}" || true
      grep -Eq "^[[:space:]]*Exec=.*%[uUfF]" "${desktop_file}"
      grep -qi "x-scheme-handler/formula" "${desktop_file}"
      grep -qi "application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet" "${desktop_file}"
      grep -qi "application/vnd\.apache\.parquet" "${desktop_file}"
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
      # Treat any `ldd` failure as fatal. Note: `ldd` can emit "not a dynamic executable"
      # for non-native-arch binaries (e.g. ARM on x86_64) and still exit non-zero.
      if [ "${status}" -ne 0 ]; then
        echo "ldd exited with status ${status}" >&2
        exit 1
      fi
    '
  echo "::endgroup::"
}

rpm_smoke_test_dir() {
  local rpm_dir="$1"
  local image="$2"

  local rpm_dir_abs
  rpm_dir_abs="$(cd "${rpm_dir}" && pwd)"

  echo "::group::.rpm smoke test (${image}) - ${rpm_dir}"
  echo "Mounting: ${rpm_dir_abs} -> /mounted"
  docker run --rm \
    ${DOCKER_PLATFORM:+--platform "${DOCKER_PLATFORM}"} \
    -v "${rpm_dir_abs}:/mounted:ro" \
    "${image}" \
    bash -euxo pipefail -c '
      echo "Container OS:"; cat /etc/os-release
      echo "Mounted artifacts:"; ls -lah /mounted
      # Install the local RPM using whatever package manager is available in the container.
      #
      # - Fedora/RHEL-family: `dnf`
      # - openSUSE: `zypper`
      #
      # Note: the Tauri updater `.sig` files are not RPM GPG signatures, and we generally do not
      # GPG-sign the built RPM in CI. Use `--nogpgcheck` / `--allow-unsigned-rpm` so the smoke test
      # validates runtime deps rather than failing on signing policy.
      if command -v dnf >/dev/null 2>&1; then
        dnf -y install --nogpgcheck --setopt=install_weak_deps=False /mounted/*.rpm
      elif command -v zypper >/dev/null 2>&1; then
        zypper --non-interactive refresh
        zypper --non-interactive install --no-recommends --allow-unsigned-rpm /mounted/*.rpm
      else
        echo "linux-package-install-smoke: no supported RPM package manager found (expected dnf or zypper)" >&2
        exit 1
      fi

      test -x /usr/bin/formula-desktop

      # Validate that installer-time MIME integration ran successfully.
      test -f /usr/share/mime/packages/app.formula.desktop.xml
      if ! grep -Eq "application/vnd\.apache\.parquet:.*\*\.parquet" /usr/share/mime/globs2; then
        echo "Missing Parquet MIME mapping in /usr/share/mime/globs2 (expected application/vnd.apache.parquet -> *.parquet)" >&2
        echo "Parquet-related globs2 lines:" >&2
        grep -n parquet /usr/share/mime/globs2 || true
        exit 1
      fi

      desktop_file="$(grep -rlE "^[[:space:]]*Exec=.*formula-desktop" /usr/share/applications 2>/dev/null | head -n 1 || true)"
      if [ -z "${desktop_file}" ]; then
        echo "No installed .desktop file found with Exec referencing formula-desktop under /usr/share/applications" >&2
        ls -lah /usr/share/applications || true
        exit 1
      fi
      echo "Installed desktop entry: ${desktop_file}"
      grep -E "^[[:space:]]*(Exec|MimeType)=" "${desktop_file}" || true
      grep -Eq "^[[:space:]]*Exec=.*%[uUfF]" "${desktop_file}"
      grep -qi "x-scheme-handler/formula" "${desktop_file}"
      grep -qi "application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet" "${desktop_file}"
      grep -qi "application/vnd\.apache\.parquet" "${desktop_file}"
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
      # Treat any `ldd` failure as fatal. Note: `ldd` can emit "not a dynamic executable"
      # for non-native-arch binaries (e.g. ARM on x86_64) and still exit non-zero.
      if [ "${status}" -ne 0 ]; then
        echo "ldd exited with status ${status}" >&2
        exit 1
      fi
    '
  echo "::endgroup::"
}

require_docker

if [[ "${kind}" == "deb" || "${kind}" == "all" ]]; then
  deb_image="${FORMULA_DEB_SMOKE_IMAGE:-ubuntu:24.04}"
  deb_dirs=()
  find_pkg_dirs "deb" ".deb" deb_dirs

  echo "::group::Pull .deb smoke test image (${deb_image})"
  docker pull ${DOCKER_PLATFORM:+--platform "${DOCKER_PLATFORM}"} "${deb_image}"
  echo "::endgroup::"

  for d in "${deb_dirs[@]}"; do
    deb_smoke_test_dir "${d}" "${deb_image}"
  done
fi

if [[ "${kind}" == "rpm" || "${kind}" == "all" ]]; then
  rpm_image="${FORMULA_RPM_SMOKE_IMAGE:-fedora:40}"
  rpm_dirs=()
  find_pkg_dirs "rpm" ".rpm" rpm_dirs

  echo "::group::Pull .rpm smoke test image (${rpm_image})"
  docker pull ${DOCKER_PLATFORM:+--platform "${DOCKER_PLATFORM}"} "${rpm_image}"
  echo "::endgroup::"

  for d in "${rpm_dirs[@]}"; do
    rpm_smoke_test_dir "${d}" "${rpm_image}"
  done
fi
