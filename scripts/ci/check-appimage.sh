#!/usr/bin/env bash

set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$repo_root"

usage() {
  cat <<'EOF'
usage: scripts/ci/check-appimage.sh [AppImage...]

If no arguments are provided, this script searches for AppImages under the usual
Cargo/Tauri output directories (e.g. target/**/release/bundle/appimage/*.AppImage).

Environment overrides:
  FORMULA_APPIMAGE_MAIN_BINARY
    Expected basename of the main binary under squashfs-root/usr/bin/.
    Defaults to `apps/desktop/src-tauri/tauri.conf.json` mainBinaryName when available.

  FORMULA_EXPECTED_ELF_MACHINE_SUBSTRING
    Expected substring from `readelf -h` "Machine:" line (defaults based on uname -m).

  FORMULA_CHECK_APPIMAGE_USE_HOST_LD_LIBRARY_PATH=1
    By default the script ignores any existing host LD_LIBRARY_PATH when running `ldd`
    (so results are deterministic and closer to a normal end-user environment). Set
    this to include the host LD_LIBRARY_PATH in addition to the extracted AppImage
    library directories.
EOF
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

# Optional overrides for CI/debugging.
# - FORMULA_APPIMAGE_MAIN_BINARY: expected basename of the main binary under `squashfs-root/usr/bin/`
# - FORMULA_EXPECTED_ELF_MACHINE_SUBSTRING: expected substring from `readelf -h` "Machine:" line
if [ -z "${FORMULA_APPIMAGE_MAIN_BINARY:-}" ]; then
  # Keep this in sync with apps/desktop/src-tauri/tauri.conf.json `mainBinaryName`.
  # Prefer reading it dynamically so the smoke test doesn't silently drift if the binary is renamed.
  if [ -f "apps/desktop/src-tauri/tauri.conf.json" ] && command -v python3 >/dev/null 2>&1; then
    FORMULA_APPIMAGE_MAIN_BINARY="$(
      python3 - <<'PY' 2>/dev/null || true
import json
with open("apps/desktop/src-tauri/tauri.conf.json", "r", encoding="utf-8") as f:
    conf = json.load(f)
print(conf.get("mainBinaryName", ""))
PY
    )"
  fi
  : "${FORMULA_APPIMAGE_MAIN_BINARY:=formula-desktop}"
fi

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
  # Use a GitHub Actions error annotation when possible; still readable locally.
  #
  # If we're inside a GitHub Actions log group, close it so failures don't swallow the rest of the
  # job logs into the group.
  if [[ "${CHECK_APPIMAGE_GROUP_OPEN:-0}" -eq 1 ]]; then
    echo "::endgroup::" >&2
    CHECK_APPIMAGE_GROUP_OPEN=0
  fi
  echo "::error::check-appimage: $*" >&2
  exit 1
}

require_cmd() {
  command -v "$1" >/dev/null 2>&1 || die "missing required command: $1"
}

expected_readelf_machine_substring() {
  if [ -n "${FORMULA_EXPECTED_ELF_MACHINE_SUBSTRING:-}" ]; then
    echo "$FORMULA_EXPECTED_ELF_MACHINE_SUBSTRING"
    return 0
  fi

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

expected_file_arch_substring() {
  # `file` prints e.g.:
  #   ELF 64-bit LSB executable, x86-64, ...
  #   ELF 64-bit LSB executable, ARM aarch64, ...
  local arch
  arch="$(uname -m)"
  case "$arch" in
    x86_64) echo "x86-64" ;;
    aarch64) echo "aarch64" ;;
    armv7l) echo "ARM" ;;
    *)
      die "unsupported runner architecture '$arch' (add mapping to scripts/ci/check-appimage.sh)"
      ;;
  esac
}

find_appimages() {
  # Prefer searching known Cargo/Tauri target directories rather than traversing the whole repo.
  local roots=()

  # Respect `CARGO_TARGET_DIR` if set (common in CI caching setups). Cargo interprets relative paths
  # relative to the build working directory (repo root in CI).
  if [[ -n "${CARGO_TARGET_DIR:-}" ]]; then
    local cargo_target_dir="${CARGO_TARGET_DIR}"
    if [[ "${cargo_target_dir}" != /* ]]; then
      cargo_target_dir="${repo_root}/${cargo_target_dir}"
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

  if [[ "${#roots[@]}" -eq 0 ]]; then
    die "no target directories found (expected CARGO_TARGET_DIR, apps/desktop/src-tauri/target, apps/desktop/target, and/or target)"
  fi

  # Canonicalize and de-dupe roots (avoid duplicate scanning when CARGO_TARGET_DIR overlaps defaults).
  local -A seen=()
  local -a uniq_roots=()
  local r abs
  for r in "${roots[@]}"; do
    abs="${r}"
    if [[ "${abs}" != /* ]]; then
      abs="${repo_root}/${abs}"
    fi
    if [[ ! -d "${abs}" ]]; then
      continue
    fi
    abs="$(cd "${abs}" && pwd -P)"
    if [[ -n "${seen[${abs}]:-}" ]]; then
      continue
    fi
    seen["${abs}"]=1
    uniq_roots+=("${abs}")
  done
  roots=("${uniq_roots[@]}")

  # Use predictable bundle globs (fast) and fall back to `find` if the layout is unexpected.
  local nullglob_was_set=0
  if shopt -q nullglob; then
    nullglob_was_set=1
  fi
  shopt -s nullglob

  local -a matches=()
  for r in "${roots[@]}"; do
    matches+=("${r}/release/bundle/appimage/"*.AppImage)
    matches+=("${r}/"*/release/bundle/appimage/*.AppImage)
  done

  if [[ "${nullglob_was_set}" -eq 0 ]]; then
    shopt -u nullglob
  fi

  if [[ "${#matches[@]}" -gt 0 ]]; then
    printf '%s\n' "${matches[@]}" | sort -u
    return 0
  fi

  # Fallback: traverse roots to locate AppImage bundles.
  find "${roots[@]}" \
    -type f \
    -name '*.AppImage' \
    -path '*/release/bundle/appimage/*.AppImage' \
    -print \
    | sort
}

find_main_binary() {
  # Best-effort heuristic to locate the application's main ELF binary inside
  # an extracted AppImage squashfs-root directory.
  local squashfs_root="$1"

  # 0) Prefer the expected main binary name (most deterministic).
  local expected_bin="$squashfs_root/usr/bin/$FORMULA_APPIMAGE_MAIN_BINARY"
  if [ -f "$expected_bin" ] && file -b "$expected_bin" | grep -q 'ELF'; then
    echo "$expected_bin"
    return 0
  fi

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
    -maxdepth 5 \
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

build_appdir_ld_library_path() {
  # Build an LD_LIBRARY_PATH that approximates what AppRun would set for an AppImage,
  # so `ldd` resolves bundled libraries inside the extracted squashfs.
  local squashfs_root="$1"
  local -a dirs=()

  for d in \
    "$squashfs_root/usr/lib" \
    "$squashfs_root/usr/lib64" \
    "$squashfs_root/lib" \
    "$squashfs_root/lib64"; do
    if [ -d "$d" ]; then
      dirs+=("$d")
    fi
  done

  # Include common multiarch subdirectories when present.
  for parent in "$squashfs_root/usr/lib" "$squashfs_root/lib"; do
    if [ -d "$parent" ]; then
      while IFS= read -r -d '' d; do
        dirs+=("$d")
      done < <(find "$parent" -maxdepth 1 -type d -name '*-linux-gnu' -print0 2>/dev/null || true)
    fi
  done

  # De-duplicate while preserving order.
  local -A seen=()
  local -a uniq=()
  local d
  for d in "${dirs[@]}"; do
    if [[ -z "${seen[$d]:-}" ]]; then
      seen["$d"]=1
      uniq+=("$d")
    fi
  done

  local joined=""
  for d in "${uniq[@]}"; do
    if [ -z "$joined" ]; then
      joined="$d"
    else
      joined="$joined:$d"
    fi
  done

  echo "$joined"
}

main() {
  require_cmd file
  require_cmd readelf
  require_cmd ldd
  require_cmd realpath

  local -a appimages=()

  if [[ $# -gt 0 ]]; then
    # Accept explicit AppImage paths (or directories containing them) for local debugging.
    local arg
    for arg in "$@"; do
      if [[ -d "$arg" ]]; then
        while IFS= read -r -d '' f; do
          appimages+=("$f")
        done < <(find "$arg" -type f -name '*.AppImage' -print0)
      else
        appimages+=("$arg")
      fi
    done
  else
    mapfile -t appimages < <(find_appimages)
  fi

  # De-dupe and validate.
  if [[ "${#appimages[@]}" -gt 0 ]]; then
    mapfile -t appimages < <(printf '%s\n' "${appimages[@]}" | sort -u)
  fi
  if [ "${#appimages[@]}" -eq 0 ]; then
    die "no AppImage artifacts found (expected something like target/**/release/bundle/appimage/*.AppImage)"
  fi
  local f
  for f in "${appimages[@]}"; do
    if [[ ! -f "$f" ]]; then
      die "AppImage path does not exist or is not a file: $f"
    fi
  done

  local expected_machine_substring
  expected_machine_substring="$(expected_readelf_machine_substring)"
  local expected_file_substring
  expected_file_substring="$(expected_file_arch_substring)"

  echo "Found ${#appimages[@]} AppImage artifact(s)."
  for appimage in "${appimages[@]}"; do
    echo "::group::check-appimage: $appimage"
    CHECK_APPIMAGE_GROUP_OPEN=1
    echo "==> Checking AppImage: $appimage"

    if [ ! -s "$appimage" ]; then
      die "AppImage '$appimage' is empty or missing"
    fi

    chmod +x "$appimage"

    local appimage_abs tmp="" squashfs_root main_bin appdir_ld_library_path ld_library_path extract_out extract_status
    appimage_abs="$(realpath "$appimage")"

    # Sanity check the AppImage runtime architecture before we try to execute it.
    # If this is the wrong arch, extraction will fail with "Exec format error"; make that
    # failure mode explicit and actionable.
    local appimage_file_info=""
    appimage_file_info="$(file "$appimage_abs" 2>/dev/null || true)"
    echo "$appimage_file_info"
    if [[ -n "$appimage_file_info" ]] && ! grep -qiF "$expected_file_substring" <<<"$appimage_file_info"; then
      die "wrong AppImage architecture for '$appimage': expected '$expected_file_substring' (runner arch $(uname -m)), got: $appimage_file_info"
    fi

    # Set up cleanup *before* mktemp so failures (e.g. mktemp itself) still close the log group.
    # shellcheck disable=SC2064
    trap 'rm -rf "${tmp:-}"; if [[ "${CHECK_APPIMAGE_GROUP_OPEN:-0}" -eq 1 ]]; then echo "::endgroup::"; fi' EXIT
    tmp="$(mktemp -d)"

    # Extract (no FUSE required). This will create ./squashfs-root.
    #
    # Capture output so failures surface clearly in CI logs without spamming on success.
    set +e
    extract_out="$(
      cd "$tmp" && "$appimage_abs" --appimage-extract 2>&1
    )"
    extract_status=$?
    set -e
    if [[ "$extract_status" -ne 0 ]]; then
      echo "$extract_out"
      die "failed to extract AppImage via --appimage-extract (status $extract_status)"
    fi

    squashfs_root="$tmp/squashfs-root"
    if [ ! -d "$squashfs_root" ]; then
      die "expected extracted directory '$squashfs_root' to exist after --appimage-extract"
    fi

    # Validate that the AppImage's `.desktop` metadata advertises our expected file associations
    # and deep-link scheme handler (x-scheme-handler/*). This is the metadata that AppImage
    # integration tools (e.g. AppImageLauncher) install into the user's desktop database.
    if command -v python3 >/dev/null 2>&1; then
      echo "::group::check-appimage: desktop integration (.desktop MimeType/Exec)"
      python3 scripts/ci/verify_linux_desktop_integration.py --package-root "$squashfs_root"
      echo "::endgroup::"
    else
      die "python3 is required to validate AppImage desktop integration (missing on PATH)"
    fi

    if ! main_bin="$(find_main_binary "$squashfs_root")"; then
      die "failed to locate main ELF binary inside extracted AppImage (looked in '$squashfs_root')"
    fi

    if [ ! -x "$main_bin" ]; then
      die "located main binary '$main_bin' but it is not executable"
    fi

    echo "Main binary: $main_bin"
    local file_out
    file_out="$(file -b "$main_bin")"
    echo "$main_bin: $file_out"
    if grep -q "not stripped" <<<"$file_out"; then
      die "main binary is not stripped (expected stripped): $file_out"
    fi

    # Ensure no DWARF debug sections are present in the shipped AppImage binary.
    # (Stripping should remove these, but guard against regressions.)
    if readelf -S --wide "$main_bin" | grep -q '\.debug_'; then
      die "main binary contains .debug_* sections (expected stripped)"
    fi

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
    appdir_ld_library_path="$(build_appdir_ld_library_path "$squashfs_root")"
    if [ -n "$appdir_ld_library_path" ]; then
      echo "Using AppImage LD_LIBRARY_PATH for ldd: $appdir_ld_library_path"
    fi

    # Prefer a deterministic library search path:
    # - Always include bundled AppImage libraries.
    # - Ignore host LD_LIBRARY_PATH by default (it can hide missing deps in CI/dev shells).
    ld_library_path="$appdir_ld_library_path"
    if [[ -n "${FORMULA_CHECK_APPIMAGE_USE_HOST_LD_LIBRARY_PATH:-}" && -n "${LD_LIBRARY_PATH:-}" ]]; then
      if [[ -n "$ld_library_path" ]]; then
        ld_library_path="$ld_library_path:$LD_LIBRARY_PATH"
      else
        ld_library_path="$LD_LIBRARY_PATH"
      fi
    fi

    set +e
    ldd_out="$(LD_LIBRARY_PATH="$ld_library_path" LD_PRELOAD= ldd "$main_bin" 2>&1)"
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
    echo "::endgroup::"
    CHECK_APPIMAGE_GROUP_OPEN=0
  done
}

main "$@"
