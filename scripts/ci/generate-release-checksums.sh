#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat <<'EOF' >&2
Usage:
  bash scripts/ci/generate-release-checksums.sh <assets_dir> [output_path]

Generate a SHA256SUMS.txt file for desktop release assets.

This script is intended to be used by the GitHub Actions release workflow after
all platform builds have uploaded assets to the draft GitHub Release.

Validation:
  The script fails if it cannot find expected platform installers/bundles and
  updater metadata in <assets_dir>.
EOF
}

if [ "${1:-}" = "-h" ] || [ "${1:-}" = "--help" ] || [ $# -lt 1 ] || [ $# -gt 2 ]; then
  usage
  exit 2
fi

assets_dir="$1"
output_path="${2:-SHA256SUMS.txt}"

if [ ! -d "$assets_dir" ]; then
  echo "checksums: ERROR assets_dir not found: $assets_dir" >&2
  exit 2
fi

hash_cmd() {
  if command -v sha256sum >/dev/null 2>&1; then
    echo "sha256sum"
    return 0
  fi
  if command -v shasum >/dev/null 2>&1; then
    # `shasum -a 256` is available on macOS and most CI images.
    echo "shasum -a 256"
    return 0
  fi
  if command -v openssl >/dev/null 2>&1; then
    # Fallback for environments without coreutils / shasum.
    echo "openssl dgst -sha256"
    return 0
  fi

  return 1
}

hash_tool="$(hash_cmd || true)"
if [ -z "${hash_tool}" ]; then
  echo "checksums: ERROR No SHA256 tool found (need sha256sum, shasum, or openssl)" >&2
  exit 2
fi

shopt -s nullglob

missing=()

# macOS: installer + updater archive
mac_dmg=( "$assets_dir"/*.dmg )
mac_app_targz=( "$assets_dir"/*.app.tar.gz "$assets_dir"/*.app.tgz )
if [ ${#mac_dmg[@]} -eq 0 ]; then
  missing+=("*.dmg")
fi
if [ ${#mac_app_targz[@]} -eq 0 ]; then
  # Some bundlers may produce a plain .tar.gz for the macOS updater bundle; accept it
  # as a fallback, but avoid confusing it with Linux AppImage tarballs.
  fallback_targz=( "$assets_dir"/*.tar.gz "$assets_dir"/*.tgz )
  filtered_fallback=()
  for f in "${fallback_targz[@]}"; do
    base="$(basename "$f")"
    if [[ "$base" == *.AppImage.tar.gz ]] || [[ "$base" == *.AppImage.tgz ]]; then
      continue
    fi
    if [[ "$base" == *.tar.gz.sig ]] || [[ "$base" == *.tgz.sig ]]; then
      continue
    fi
    filtered_fallback+=( "$f" )
  done
  if [ ${#filtered_fallback[@]} -eq 0 ]; then
    missing+=("*.app.tar.gz (or a non-AppImage *.tar.gz/*.tgz)")
  fi
fi

# Windows: installer(s)
win_msi=( "$assets_dir"/*.msi )
win_exe=( "$assets_dir"/*.exe )
if [ ${#win_msi[@]} -eq 0 ]; then
  missing+=("*.msi")
fi
if [ ${#win_exe[@]} -eq 0 ]; then
  missing+=("*.exe")
fi

# Linux: AppImage + package formats
linux_appimage=( "$assets_dir"/*.AppImage )
linux_deb=( "$assets_dir"/*.deb )
linux_rpm=( "$assets_dir"/*.rpm )
if [ ${#linux_appimage[@]} -eq 0 ]; then
  missing+=("*.AppImage")
fi
if [ ${#linux_deb[@]} -eq 0 ]; then
  missing+=("*.deb")
fi
if [ ${#linux_rpm[@]} -eq 0 ]; then
  missing+=("*.rpm")
fi

# Updater metadata (used by Tauri's built-in updater endpoints).
if [ ! -f "$assets_dir/latest.json" ]; then
  missing+=("latest.json")
fi
if [ ! -f "$assets_dir/latest.json.sig" ]; then
  missing+=("latest.json.sig")
fi

if [ ${#missing[@]} -ne 0 ]; then
  {
    echo "checksums: ERROR Missing expected release assets:"
    for m in "${missing[@]}"; do
      echo "checksums: - $m"
    done
    echo ""
    echo "checksums: Found assets:"
    (cd "$assets_dir" && find . -maxdepth 1 -type f -print | sed 's|^\./|checksums: - |' | sort) || true
  } >&2
  exit 1
fi

# Hash all files in the release asset directory (excluding a pre-existing
# SHA256SUMS.txt from previous workflow runs).
mapfile -t rel_files < <(
  cd "$assets_dir"
  find . -maxdepth 1 -type f -print \
    | sed 's|^\./||' \
    | awk '$0 != "SHA256SUMS.txt"' \
    | sort
)

if [ ${#rel_files[@]} -eq 0 ]; then
  echo "checksums: ERROR No release assets found to hash in $assets_dir" >&2
  exit 1
fi

tmp_out="$(mktemp "${TMPDIR:-/tmp}/SHA256SUMS.XXXXXX")"
(
  cd "$assets_dir"
  if [[ "$hash_tool" == "openssl dgst -sha256" ]]; then
    # openssl output format:
    #   SHA256(filename)= <hash>
    for f in "${rel_files[@]}"; do
      openssl dgst -sha256 "$f" | awk -v file="$f" '{print $NF "  " file}'
    done >"$tmp_out"
  else
    # sha256sum / shasum format:
    #   <hash>  <filename>
    # `shasum -a 256` also prints `<hash>  <filename>`.
    # shellcheck disable=SC2086
    $hash_tool "${rel_files[@]}" >"$tmp_out"
  fi
)

mv "$tmp_out" "$output_path"

echo "checksums: Wrote $(wc -l <"$output_path" | tr -d ' ') checksums to $output_path"
