#!/usr/bin/env bash

set -euo pipefail

##
# CI smoke test for produced Linux AppImage artifacts.
#
# What we check (without launching the GUI):
# - The AppImage is not corrupted and can be extracted via `--appimage-extract` (no FUSE required).
# - We can locate the main application ELF binary inside the extracted filesystem.
# - The binary architecture matches the current runner's architecture (e.g. x86-64).
# - `ldd` resolves all shared library dependencies (no "not found").
#
# This script is meant to be run after the Linux release build has produced:
#   target/**/release/bundle/appimage/*.AppImage
##

die() {
  echo "error: $*" >&2
  exit 1
}

require_cmd() {
  command -v "$1" >/dev/null 2>&1 || die "missing required command: $1"
}

expected_readelf_machine_substring() {
  # `readelf -h` prints e.g.:
  #   Machine:                           Advanced Micro Devices X86-64
  #   Machine:                           AArch64
  local arch
  arch="$(uname -m)"
  case "$arch" in
    x86_64) echo "X86-64" ;;
    aarch64) echo "AArch64" ;;
    armv7l) echo "ARM" ;;
    *)
      die "unsupported runner architecture '$arch' (add mapping to scripts/ci/check-appimage.sh)"
      ;;
  esac
}

find_appimages() {
  # We intentionally use `find` rather than globstar for portability and to avoid
  # relying on shell options.
  # Note: `find`'s default implicit `-print` interacts poorly with `-o` due to
  # operator precedence. Always parenthesize `-o` expressions and add an
  # explicit `-print`.
  find . \
    -type f \
    -name '*.AppImage' \
    \( \
      -path '*/target/*/release/bundle/appimage/*.AppImage' -o \
      -path '*/target/release/bundle/appimage/*.AppImage' \
    \) \
    -print \
    | sort
}

find_main_binary() {
  # Best-effort heuristic to locate the application's main ELF binary inside
  # an extracted AppImage squashfs-root directory.
  local squashfs_root="$1"

  # 1) Prefer AppRun if it's an ELF or a symlink to an ELF.
  if [ -e "$squashfs_root/AppRun" ]; then
    local apprun="$squashfs_root/AppRun"
    if [ -L "$apprun" ]; then
      local resolved
      resolved="$(readlink -f "$apprun")"
      if [ -n "$resolved" ] && [ -f "$resolved" ]; then
        if file -b "$resolved" | grep -q 'ELF'; then
          echo "$resolved"
          return 0
        fi
      fi
    elif [ -f "$apprun" ]; then
      if file -b "$apprun" | grep -q 'ELF'; then
        echo "$apprun"
        return 0
      fi
    fi
  fi

  # 2) Try to read the Exec= line from the desktop file.
  local desktop_file=""
  desktop_file="$(find "$squashfs_root" \
    -maxdepth 3 \
    -type f \
    -name '*.desktop' \
    2>/dev/null \
    | head -n1 || true)"
  if [ -n "$desktop_file" ] && [ -f "$desktop_file" ]; then
    local exec_line=""
    exec_line="$(grep -m1 '^Exec=' "$desktop_file" || true)"
    if [ -n "$exec_line" ]; then
      local exec_value exec_cmd candidate
      exec_value="${exec_line#Exec=}"
      # Take the first token; typical Tauri AppImages use:
      #   Exec=formula-desktop %U
      exec_cmd="${exec_value%% *}"
      exec_cmd="${exec_cmd%\"}"
      exec_cmd="${exec_cmd#\"}"

      if [[ "$exec_cmd" == /* ]]; then
        candidate="$squashfs_root$exec_cmd"
      else
        candidate="$squashfs_root/usr/bin/$exec_cmd"
      fi

      if [ -f "$candidate" ] && file -b "$candidate" | grep -q 'ELF'; then
        echo "$candidate"
        return 0
      fi
    fi
  fi

  # 3) Fall back to the first ELF executable in usr/bin.
  if [ -d "$squashfs_root/usr/bin" ]; then
    local bin
    while IFS= read -r -d '' bin; do
      if file -b "$bin" | grep -q 'ELF'; then
        echo "$bin"
        return 0
      fi
    done < <(find "$squashfs_root/usr/bin" -maxdepth 1 -type f -executable -print0 2>/dev/null)
  fi

  return 1
}

main() {
  require_cmd file
  require_cmd readelf
  require_cmd ldd

  mapfile -t appimages < <(find_appimages)
  if [ "${#appimages[@]}" -eq 0 ]; then
    die "no AppImage artifacts found (expected something like target/**/release/bundle/appimage/*.AppImage)"
  fi

  local expected_machine_substring
  expected_machine_substring="$(expected_readelf_machine_substring)"

  echo "Found ${#appimages[@]} AppImage artifact(s)."
  for appimage in "${appimages[@]}"; do
    echo "==> Checking AppImage: $appimage"

    if [ ! -s "$appimage" ]; then
      die "AppImage '$appimage' is empty or missing"
    fi

    chmod +x "$appimage"

    local appimage_abs tmp squashfs_root main_bin
    appimage_abs="$(realpath "$appimage")"
    tmp="$(mktemp -d)"
    # shellcheck disable=SC2064
    trap "rm -rf '$tmp'" EXIT

    (
      cd "$tmp"
      # Extract (no FUSE required). This will create ./squashfs-root
      "$appimage_abs" --appimage-extract >/dev/null
    )

    squashfs_root="$tmp/squashfs-root"
    if [ ! -d "$squashfs_root" ]; then
      die "expected extracted directory '$squashfs_root' to exist after --appimage-extract"
    fi

    if ! main_bin="$(find_main_binary "$squashfs_root")"; then
      die "failed to locate main ELF binary inside extracted AppImage (looked in '$squashfs_root')"
    fi

    if [ ! -x "$main_bin" ]; then
      die "located main binary '$main_bin' but it is not executable"
    fi

    echo "Main binary: $main_bin"
    file "$main_bin"

    # Architecture assertion.
    local machine_line=""
    machine_line="$(readelf -h "$main_bin" | awk -F: '/Machine:/{gsub(/^[ \t]+/, "", $2); print $2; exit}')"
    if [ -z "$machine_line" ]; then
      die "failed to read ELF machine type via 'readelf -h' for '$main_bin'"
    fi
    echo "ELF machine: $machine_line (expected contains: $expected_machine_substring)"
    if ! grep -qF "$expected_machine_substring" <<<"$machine_line"; then
      die "wrong architecture for '$main_bin': got '$machine_line', expected '$expected_machine_substring'"
    fi

    # Shared library dependency check.
    set +e
    ldd_out="$(ldd "$main_bin" 2>&1)"
    ldd_status=$?
    set -e

    echo "$ldd_out"
    if [ "$ldd_status" -ne 0 ] && ! grep -qF "not a dynamic executable" <<<"$ldd_out"; then
      die "ldd failed for '$main_bin' (status $ldd_status)"
    fi
    if grep -qF "not found" <<<"$ldd_out"; then
      die "missing shared library dependencies detected for '$main_bin' (see above for 'not found')"
    fi

    rm -rf "$tmp"
    trap - EXIT
    echo "OK: $appimage"
  done
}

main "$@"
