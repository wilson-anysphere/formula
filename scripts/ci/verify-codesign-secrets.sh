#!/usr/bin/env bash

set -euo pipefail

usage() {
  cat >&2 <<'EOF'
Usage: verify-codesign-secrets.sh <platform>

Validates that code signing certificate secrets are:
  - valid base64
  - valid PKCS#12 archives
  - decryptable with the provided password

Platforms:
  macos    Requires APPLE_CERTIFICATE + APPLE_CERTIFICATE_PASSWORD
  windows  Requires WINDOWS_CERTIFICATE + WINDOWS_CERTIFICATE_PASSWORD

If FORMULA_REQUIRE_CODESIGN=1, missing required secrets will fail CI with a
message pointing at docs/release.md.
EOF
}

platform="${1:-}"
if [[ -z "$platform" ]]; then
  usage
  exit 2
fi

require_codesign="${FORMULA_REQUIRE_CODESIGN:-}"
codesign_required=0
if [[ "$require_codesign" == "1" ]]; then
  codesign_required=1
fi

if ! command -v openssl >/dev/null 2>&1; then
  echo "openssl is required for code signing secret validation but was not found on PATH." >&2
  exit 1
fi

tmp_dir="$(mktemp -d 2>/dev/null || mktemp -d -t 'formula-codesign-secrets')"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT

docs_hint="See docs/release.md for required GitHub Actions secrets."

fail_missing_secret() {
  local message="$1"
  echo "::error::${message} ${docs_hint}" >&2
  exit 1
}

verify_pkcs12_secret() {
  local secret_name="$1"
  local secret_value="$2"
  local password_name="$3"
  local password_value="$4"
  local out_file="$5"

  local base64_decoder=()
  if command -v base64 >/dev/null 2>&1; then
    if printf 'TQ==' | base64 --decode >/dev/null 2>&1; then
      base64_decoder=(base64 --decode)
    elif printf 'TQ==' | base64 -D >/dev/null 2>&1; then
      base64_decoder=(base64 -D)
    elif printf 'TQ==' | base64 -d >/dev/null 2>&1; then
      base64_decoder=(base64 -d)
    fi
  fi

  if [[ -z "$secret_value" ]]; then
    if [[ "$codesign_required" -eq 1 ]]; then
      fail_missing_secret "FORMULA_REQUIRE_CODESIGN=1 but ${secret_name} is missing."
    fi
    echo "${secret_name} not set; skipping code signing secret preflight."
    return 0
  fi

  if [[ -z "$password_value" ]]; then
    fail_missing_secret "${secret_name} is set but ${password_name} is missing."
  fi

  # Decode base64 -> PKCS#12 archive
  if ((${#base64_decoder[@]})); then
    if ! printf '%s' "$secret_value" | "${base64_decoder[@]}" >"$out_file"; then
      echo "::error::Failed to base64-decode ${secret_name}. Ensure it is a base64-encoded PKCS#12 archive (.p12/.pfx). ${docs_hint}" >&2
      exit 1
    fi
  elif command -v python3 >/dev/null 2>&1; then
    if ! CODESIGN_SECRET_VALUE="$secret_value" CODESIGN_OUT_FILE="$out_file" python3 - <<'PY'
import base64
import os
import sys

data = os.environ.get("CODESIGN_SECRET_VALUE", "")
out_path = os.environ.get("CODESIGN_OUT_FILE", "")

# Secrets are often generated via `base64` with newlines; allow whitespace and
# validate the underlying base64 payload strictly.
data = "".join(data.split())

try:
    raw = base64.b64decode(data, validate=True)
except Exception as e:
    print(f"Invalid base64: {e}", file=sys.stderr)
    sys.exit(1)

with open(out_path, "wb") as f:
    f.write(raw)
PY
    then
      echo "::error::Failed to base64-decode ${secret_name}. Ensure it is a base64-encoded PKCS#12 archive (.p12/.pfx). ${docs_hint}" >&2
      exit 1
    fi
  else
    # Fallback: OpenSSL's base64 decoder is more permissive, but still better than nothing.
    if ! printf '%s' "$secret_value" | openssl base64 -d -A -out "$out_file"; then
      echo "::error::Failed to base64-decode ${secret_name}. Ensure it is a base64-encoded PKCS#12 archive (.p12/.pfx). ${docs_hint}" >&2
      exit 1
    fi
  fi

  if [[ ! -s "$out_file" ]]; then
    echo "::error::Failed to base64-decode ${secret_name}. Ensure it is a base64-encoded PKCS#12 archive (.p12/.pfx). ${docs_hint}" >&2
    exit 1
  fi

  # Verify PKCS#12 can be decrypted/parsed with the given password.
  #
  # - `-nodes` ensures OpenSSL actually decrypts the private key material.
  # - stdout is discarded to avoid leaking key material in logs.
  if ! openssl pkcs12 -in "$out_file" -nodes -passin "pass:${password_value}" >/dev/null; then
    echo "::error::Failed to decrypt/parse ${secret_name} with ${password_name}. Ensure the password matches the certificate export password. ${docs_hint}" >&2
    exit 1
  fi
}

case "$platform" in
  macos)
    verify_pkcs12_secret \
      "APPLE_CERTIFICATE" \
      "${APPLE_CERTIFICATE:-}" \
      "APPLE_CERTIFICATE_PASSWORD" \
      "${APPLE_CERTIFICATE_PASSWORD:-}" \
      "${tmp_dir}/apple-certificate.p12"
    ;;
  windows)
    verify_pkcs12_secret \
      "WINDOWS_CERTIFICATE" \
      "${WINDOWS_CERTIFICATE:-}" \
      "WINDOWS_CERTIFICATE_PASSWORD" \
      "${WINDOWS_CERTIFICATE_PASSWORD:-}" \
      "${tmp_dir}/windows-certificate.pfx"
    ;;
  *)
    echo "Unknown platform '$platform'." >&2
    usage
    exit 2
    ;;
esac
