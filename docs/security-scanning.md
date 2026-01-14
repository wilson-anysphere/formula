# Security Scanning (CI + Local)

This repository is intended to be a **Rust + Node (pnpm) monorepo** (Tauri desktop + supporting services). To keep the supply chain and configuration surface hardened over time, we run a dedicated security workflow on every PR and on a weekly schedule.

## What runs in CI

GitHub Actions workflow: `.github/workflows/security.yml`

Each run produces a **single “security-report” artifact** containing raw tool output plus a human-readable summary (`security-report/summary.md`).

### 1) Dependency vulnerability scanning

**Rust**
- Tool: `cargo audit`
- Policy: CI fails on **high/critical** vulnerabilities.
- Severity handling: if an advisory does not include an explicit severity, CI derives it from the CVSS vector when available; advisories without severity *and* without CVSS are treated as **high** (fail-safe) to force explicit triage/allowlisting.
- Allowlist: `security/allowlist/cargo-audit.txt` (one `RUSTSEC-*` ID per line, comments allowed)

**Node**
- Tool: `pnpm audit` (fallback to `npm audit` if pnpm is not used)
- Policy: CI fails on **high/critical** vulnerabilities.
- Allowlist: `security/allowlist/node-audit.txt` (one `GHSA-*`, `CVE-*`, or npm advisory ID per line, comments allowed)

### 2) Static analysis (SAST)

**Rust**
- Tool: `cargo clippy`
- Policy: CI fails on placeholder macros like `todo!()` / `unimplemented!()`.
  - `unwrap` / `expect` / `panic!` are intentionally *not* denied at the workflow level yet because
    they are still used throughout the codebase; tighten these once the repo is ready to enforce
    them consistently.
  - Clippy runs with default features (not `--all-features`) so optional desktop/system-dependency
    builds don't break headless CI.

**JS/TS**
- Tooling expectation:
  - ESLint configured with security rules (e.g. `eslint-plugin-security`)
  - TypeScript compiler in `strict` mode

The security workflow will enforce these expectations once the monorepo includes `package.json`/`tsconfig.json`.

### 3) Secret scanning

- Tool: `gitleaks`
- Scope: scans the working tree (files in the PR), and fails if any secrets are detected.

### 4) Configuration hardening checks

The workflow includes best-effort checks for:
- **Tauri (desktop)** capabilities/CSP hardening (Tauri v2 uses “capabilities” instead of the old v1 allowlist; see `apps/desktop/src-tauri/tauri.conf.json` and `docs/11-desktop-shell.md`)
- **API/sync services** secure defaults (detects obviously unsafe patterns like wildcard CORS and disabled TLS verification)

These checks are designed to **fail only when relevant files exist**. As the codebase grows, expand the checks to cover additional configurations.

## How to run locally

The CI workflow runs a single entrypoint script:

```bash
./scripts/security/ci.sh
```

Prerequisites:
- Rust toolchain pinned by `rust-toolchain.toml` (includes `clippy` + `rustfmt` components)
- `cargo-audit` (`bash scripts/cargo_agent.sh install cargo-audit --locked`)
- Node + pnpm (via Corepack)
- `gitleaks` (install from https://github.com/gitleaks/gitleaks)

## Handling findings

### Dependency vulnerabilities

1. **Confirm severity and reachability**
   - Is the vulnerable dependency only used in dev tooling?
   - Is the affected code path reachable in the shipped product?
2. **Upgrade / patch first**
   - Prefer upgrading the direct dependency.
   - If blocked, upgrade transitive dependencies via overrides/resolutions.
3. **Allowlist only as a last resort**
   - If the vulnerability is a false positive or non-reachable, allowlist it with context.

#### Allowlisting (Rust)

Edit: `security/allowlist/cargo-audit.txt`

Example:

```text
# RUSTSEC-2024-0000 — false positive in build-only dependency (tracked in ISSUE-123)
RUSTSEC-2024-0000
```

#### Allowlisting (Node)

Edit: `security/allowlist/node-audit.txt`

Example:

```text
# GHSA-xxxx-xxxx-xxxx — dev-only tooling, blocked on upstream fix (tracked in ISSUE-456)
GHSA-xxxx-xxxx-xxxx
```

### Secrets

1. **Rotate the secret immediately** (assume it is compromised).
2. **Remove it from the repo** and update any deployments/config.
3. If the secret may exist in git history, follow GitHub’s guidance for rewriting history and invalidating tokens.

## Pre-commit hook (recommended)

Add a local pre-commit hook to prevent accidental secret commits:

```bash
cat > .git/hooks/pre-commit <<'HOOK'
#!/usr/bin/env bash
set -euo pipefail

if command -v gitleaks >/dev/null 2>&1; then
  gitleaks protect --staged --redact
else
  echo "gitleaks not installed; skipping secret scan"
fi
HOOK

chmod +x .git/hooks/pre-commit
```

If you use the `pre-commit` framework, add `gitleaks` as a hook and run it on staged changes.
