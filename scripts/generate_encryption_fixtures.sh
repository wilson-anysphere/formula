#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${repo_root}"

# Regenerate `fixtures/encryption/*` from the small plaintext templates shipped in the repo.
bash scripts/cargo_agent.sh run -p formula-office-crypto --example generate_fixtures

