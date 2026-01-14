#!/usr/bin/env bash
set -euo pipefail

# Regenerate the deterministic MS-OFFCRYPTO `EncryptedPackage` fixtures under `fixtures/encryption/`.
#
# Usage:
#   ./tools/generate_office_crypto_fixtures.sh
#
# Note: This uses the `formula-office-crypto` internal fixture generator binary. The generated
# files are committed to git so CI does not need to run this.

bash scripts/cargo_agent.sh run -p formula-office-crypto --bin generate_fixtures

