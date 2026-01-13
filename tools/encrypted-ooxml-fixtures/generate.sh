#!/usr/bin/env bash
set -euo pipefail

# Generate encrypted OOXML `.xlsx` fixtures using Apache POI.
#
# This script intentionally does NOT commit jar binaries to the repo. Instead it:
#   1) downloads a pinned set of jars from Maven Central into a local cache dir
#   2) verifies SHA-256 checksums (supply-chain safety)
#   3) compiles `GenerateEncryptedXlsx.java`
#   4) runs it with the provided args
#
# Apache POI version: 5.2.5
#
# Usage:
#   tools/encrypted-ooxml-fixtures/generate.sh agile|standard <password> <in_plaintext_xlsx> <out_encrypted_xlsx>
#
# Example (from repo root):
#   tools/encrypted-ooxml-fixtures/generate.sh agile password fixtures/xlsx/basic/basic.xlsx fixtures/encrypted/ooxml/agile.xlsx
#   tools/encrypted-ooxml-fixtures/generate.sh standard password fixtures/xlsx/basic/basic.xlsx fixtures/encrypted/ooxml/standard.xlsx

usage() {
  cat >&2 <<'EOF'
Usage: generate.sh <mode> <password> <in_plaintext_xlsx> <out_encrypted_xlsx>
  mode: agile | standard

Example:
  tools/encrypted-ooxml-fixtures/generate.sh agile password fixtures/xlsx/basic/basic.xlsx fixtures/encrypted/ooxml/agile.xlsx
EOF
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

if [[ $# -ne 4 ]]; then
  usage
  exit 2
fi

MODE="$1"
PASSWORD="$2"
IN_PLAINTEXT_XLSX="$3"
OUT_ENCRYPTED_XLSX="$4"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CACHE_DIR="${SCRIPT_DIR}/.cache"
JARS_DIR="${CACHE_DIR}/jars"
CLASSES_DIR="${CACHE_DIR}/classes"

mkdir -p "${JARS_DIR}" "${CLASSES_DIR}"

download_jar() {
  local filename="$1"
  local url="$2"
  local expected_sha="$3"
  local dest="${JARS_DIR}/${filename}"

  if [[ -f "${dest}" ]]; then
    local actual_sha
    actual_sha="$(sha256sum "${dest}" | awk '{print $1}')"
    if [[ "${actual_sha}" == "${expected_sha}" ]]; then
      return 0
    fi
    echo "SHA256 mismatch for cached ${filename} (expected ${expected_sha}, got ${actual_sha}); re-downloading" >&2
    rm -f "${dest}"
  fi

  echo "Downloading ${filename}..." >&2
  curl -sSL --fail -o "${dest}.tmp" "${url}"

  local actual_sha
  actual_sha="$(sha256sum "${dest}.tmp" | awk '{print $1}')"
  if [[ "${actual_sha}" != "${expected_sha}" ]]; then
    echo "SHA256 mismatch for ${filename} (expected ${expected_sha}, got ${actual_sha})" >&2
    rm -f "${dest}.tmp"
    exit 1
  fi

  mv "${dest}.tmp" "${dest}"
}

BASE_URL="https://repo1.maven.org/maven2"

# Pinned jar set (Apache POI 5.2.5 + runtime deps) with SHA-256 checksums.
download_jar "poi-5.2.5.jar" "${BASE_URL}/org/apache/poi/poi/5.2.5/poi-5.2.5.jar" "352e1b44a5777af2df3d7dc408cda9f75f932d0e0125fa1a7d336a13c0a663a7"
download_jar "commons-io-2.15.0.jar" "${BASE_URL}/commons-io/commons-io/2.15.0/commons-io-2.15.0.jar" "a328dad730921d197b6a9b195dffa00e41c974c2dac8fe37e84d31706bca7792"
download_jar "commons-collections4-4.4.jar" "${BASE_URL}/org/apache/commons/commons-collections4/4.4/commons-collections4-4.4.jar" "1df8b9430b5c8ed143d7815e403e33ef5371b2400aadbe9bda0883762e0846d1"
download_jar "commons-codec-1.16.0.jar" "${BASE_URL}/commons-codec/commons-codec/1.16.0/commons-codec-1.16.0.jar" "56595fb20b0b85bc91d0d503dad50bb7f1b9afc0eed5dffa6cbb25929000484d"
download_jar "commons-math3-3.6.1.jar" "${BASE_URL}/org/apache/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar" "1e56d7b058d28b65abd256b8458e3885b674c1d588fa43cd7d1cbb9c7ef2b308"
download_jar "SparseBitSet-1.3.jar" "${BASE_URL}/com/zaxxer/SparseBitSet/1.3/SparseBitSet-1.3.jar" "f76b85adb0c00721ae267b7cfde4da7f71d3121cc2160c9fc00c0c89f8c53c8a"
download_jar "log4j-api-2.21.1.jar" "${BASE_URL}/org/apache/logging/log4j/log4j-api/2.21.1/log4j-api-2.21.1.jar" "1db48e180881bef1deb502022006a025a248d8f6a26186789b0c7ce487c602d6"
download_jar "log4j-core-2.21.1.jar" "${BASE_URL}/org/apache/logging/log4j/log4j-core/2.21.1/log4j-core-2.21.1.jar" "ad00ba17c77ff3efd7e461cedf3b825888cea95abe343467c8adb5e3912a72a1"

REQUIRED_JARS=(
  "${JARS_DIR}/poi-5.2.5.jar"
  "${JARS_DIR}/commons-codec-1.16.0.jar"
  "${JARS_DIR}/commons-collections4-4.4.jar"
  "${JARS_DIR}/commons-math3-3.6.1.jar"
  "${JARS_DIR}/commons-io-2.15.0.jar"
  "${JARS_DIR}/SparseBitSet-1.3.jar"
  "${JARS_DIR}/log4j-api-2.21.1.jar"
  "${JARS_DIR}/log4j-core-2.21.1.jar"
)

CP_JARS=""
for jar in "${REQUIRED_JARS[@]}"; do
  CP_JARS="${CP_JARS}:${jar}"
done
CP_JARS="${CP_JARS#:}"

SOURCE="${SCRIPT_DIR}/GenerateEncryptedXlsx.java"
MAIN_CLASS="GenerateEncryptedXlsx"
MAIN_CLASSFILE="${CLASSES_DIR}/${MAIN_CLASS}.class"

if [[ ! -f "${MAIN_CLASSFILE}" || "${SOURCE}" -nt "${MAIN_CLASSFILE}" ]]; then
  echo "Compiling ${SOURCE}..." >&2
  # Some jars on the classpath include annotation processors; disable annotation processing to avoid
  # noisy warnings (this tool doesn't rely on annotation processing).
  javac -proc:none -classpath "${CP_JARS}" -d "${CLASSES_DIR}" "${SOURCE}"
fi

java -classpath "${CLASSES_DIR}:${CP_JARS}" "${MAIN_CLASS}" "${MODE}" "${PASSWORD}" "${IN_PLAINTEXT_XLSX}" "${OUT_ENCRYPTED_XLSX}"
