#!/usr/bin/env bash
set -euo pipefail

# Cargo runner for wasm32 tests.
#
# This repo uses a per-repo `CARGO_HOME` (see `scripts/agent-init.sh` / `scripts/cargo_agent.sh`)
# to avoid cross-agent lock contention. That means `wasm-bindgen-test-runner` is often *not*
# available on PATH when running `cargo test --target wasm32-unknown-unknown`.
#
# Ensure the runner is installed into `$CARGO_HOME/bin` on demand, then delegate.

WASM_BINDGEN_CLI_VERSION="0.2.106"

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd "${script_dir}/.." && pwd)"

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml`. Some environments set it globally
# (often to `stable`), which would bypass the pinned toolchain and reintroduce drift for cargo
# installs invoked by this runner.
if [[ -n "${RUSTUP_TOOLCHAIN:-}" && -f "${repo_root}/rust-toolchain.toml" ]]; then
  unset RUSTUP_TOOLCHAIN
fi

cargo_home="${CARGO_HOME:-${HOME:-/root}/.cargo}"
bin_dir="${cargo_home%/}/bin"
runner="${bin_dir}/wasm-bindgen-test-runner"

if [[ ! -x "${runner}" ]]; then
  mkdir -p "${bin_dir}"

  # Best-effort cross-process lock: `cargo test` may spawn multiple wasm test binaries.
  lock_dir="${bin_dir}/.wasm-bindgen-test-runner.install-lock"
  if mkdir "${lock_dir}" 2>/dev/null; then
    trap 'rmdir "${lock_dir}" 2>/dev/null || true' EXIT
    if [[ ! -x "${runner}" ]]; then
      echo "Installing wasm-bindgen-test-runner (wasm-bindgen-cli ${WASM_BINDGEN_CLI_VERSION})..." >&2
      cargo install wasm-bindgen-cli --version "${WASM_BINDGEN_CLI_VERSION}" \
        --bin wasm-bindgen-test-runner --quiet
    fi
  else
    # Another process is installing; wait until the binary appears.
    while [[ ! -x "${runner}" ]]; do
      sleep 0.1
    done
  fi
fi

exec "${runner}" "$@"
