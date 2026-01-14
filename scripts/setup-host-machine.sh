#!/bin/bash
# ONE-TIME HOST MACHINE SETUP
# Run this ONCE on the EC2 instance before agents start
# Requires sudo privileges
#
# Usage: sudo ./scripts/setup-host-machine.sh

set -e

if [ "$EUID" -ne 0 ]; then
  echo "Please run with sudo: sudo $0"
  exit 1
fi

echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Formula Development Host Setup                                 ║"
echo "╚════════════════════════════════════════════════════════════════╝"
echo ""

# ============================================================================
# System Limits
# ============================================================================
echo "=== Configuring System Limits ==="

# Increase inotify limits (for file watchers across many repos)
if ! grep -q "fs.inotify.max_user_watches" /etc/sysctl.conf; then
  echo "fs.inotify.max_user_watches=524288" >> /etc/sysctl.conf
  echo "fs.inotify.max_user_instances=8192" >> /etc/sysctl.conf
  sysctl -p
  echo "✓ Increased inotify limits"
else
  echo "✓ inotify limits already configured"
fi

# Increase file descriptor limits
if ! grep -q "nofile" /etc/security/limits.conf | grep -q 1048576; then
  cat >> /etc/security/limits.conf << 'EOF'
*               soft    nofile          1048576
*               hard    nofile          1048576
root            soft    nofile          1048576
root            hard    nofile          1048576
EOF
  echo "✓ Increased file descriptor limits"
else
  echo "✓ File descriptor limits already configured"
fi

# ============================================================================
# Essential Packages
# ============================================================================
echo ""
echo "=== Installing Essential Packages ==="

apt-get update

# Build essentials
apt-get install -y \
  build-essential \
  pkg-config \
  cmake \
  git \
  curl \
  wget \
  unzip

# For Canvas/graphics
apt-get install -y \
  libcairo2-dev \
  libjpeg-dev \
  libpango1.0-dev \
  libgif-dev \
  librsvg2-dev \
  libpixman-1-dev

# For Tauri
apt-get install -y \
  file \
  libgtk-3-dev \
  libssl-dev \
  librsvg2-dev \
  patchelf \
  squashfs-tools \
  fakeroot \
  rpm \
  cpio

# WebKitGTK dev package name differs across Ubuntu versions.
apt-get install -y libwebkit2gtk-4.1-dev || apt-get install -y libwebkit2gtk-4.0-dev

# AppIndicator dev package name differs across Ubuntu versions (Ubuntu 24.04 prefers Ayatana).
apt-get install -y libayatana-appindicator3-dev || apt-get install -y libappindicator3-dev

# `appimagetool` is distributed as an AppImage and requires the FUSE 2 runtime.
# Ubuntu 24.04 uses `libfuse2t64` as part of the time_t 64-bit transition.
apt-get install -y libfuse2 || apt-get install -y libfuse2t64

# For headless browser testing
apt-get install -y \
  xvfb \
  libnss3 \
  libatk1.0-0 \
  libatk-bridge2.0-0 \
  libcups2 \
  libxkbcommon0 \
  libxcomposite1 \
  libxdamage1 \
  libxfixes3 \
  libxrandr2 \
  libgbm1 \
  libpango-1.0-0 \
  libcairo2 \
  libasound2

# Fonts (for consistent text rendering)
apt-get install -y \
  fonts-liberation \
  fonts-dejavu-core \
  fonts-noto-core \
  fonts-noto-cjk \
  fontconfig

# Refresh font cache
fc-cache -f -v >/dev/null 2>&1

echo "✓ Essential packages installed"

# ============================================================================
# Development Tools (if not already installed)
# ============================================================================
echo ""
echo "=== Checking Development Tools ==="

# Node.js (keep in sync with CI/release workflows)
if ! command -v node &> /dev/null; then
  echo "⚠️  Node.js not found. Install via nvm or your preferred method."
else
  node_version="$(node --version)"
  echo "✓ Node.js: ${node_version}"
  node_major="$(node -p "process.versions.node.split('.')[0]")"

  # Prefer the repo-pinned Node major from `.nvmrc` when available.
  # Fall back to `.node-version` (asdf-style) if present.
  # Fall back to `mise.toml` if the repo uses mise for local tooling.
  # Fall back to 22 to preserve historical guidance if neither file is present/unparseable.
  script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
  repo_root="$(cd "${script_dir}/.." && pwd)"
  expected_node_major=""
  if [ -f "${repo_root}/.nvmrc" ]; then
    expected_node_major="$(head -n 1 "${repo_root}/.nvmrc" | tr -d '[:space:]' | sed -E 's/^[vV]?([0-9]+).*/\\1/')"
  elif [ -f "${repo_root}/.node-version" ]; then
    expected_node_major="$(head -n 1 "${repo_root}/.node-version" | tr -d '[:space:]' | sed -E 's/^[vV]?([0-9]+).*/\\1/')"
  elif [ -f "${repo_root}/mise.toml" ]; then
    expected_node_major="$(grep -E '^[[:space:]]*node[[:space:]]*=' "${repo_root}/mise.toml" | head -n 1 | sed -E 's/.*=[[:space:]]*\"?([0-9]+).*/\\1/')"
  fi
  if ! [[ "${expected_node_major}" =~ ^[0-9]+$ ]] && [ -f "${repo_root}/.github/workflows/ci.yml" ]; then
    # Avoid matching YAML-like strings inside block scalar bodies (e.g. `run: |` scripts) so
    # non-semantic text can't influence this best-effort extraction.
    expected_node_major="$(
      awk '
        function indent(s) {
          match(s, /^[ ]*/);
          return RLENGTH;
        }
        BEGIN {
          in_block = 0;
          block_indent = 0;
          block_re = ":[[:space:]]*[>|][0-9+-]*[[:space:]]*$";
        }
        {
          raw = $0;
          sub(/\r$/, "", raw);
          ind = indent(raw);

          if (in_block) {
            if (raw ~ /^[[:space:]]*$/) next;
            if (ind > block_indent) next;
            in_block = 0;
          }

          trimmed = raw;
          sub(/^[[:space:]]*/, "", trimmed);
          if (trimmed ~ /^#/) next;

          line = raw;
          sub(/#.*/, "", line);
          is_block = (line ~ block_re);

          if (line ~ /^[[:space:]]*NODE_VERSION[[:space:]]*:/) {
            value = line;
            sub(/^[[:space:]]*NODE_VERSION[[:space:]]*:[[:space:]]*/, "", value);
            if (match(value, /[0-9]+/)) {
              print substr(value, RSTART, RLENGTH);
            }
            exit;
          }

          if (is_block) {
            in_block = 1;
            block_indent = ind;
          }
        }
      ' "${repo_root}/.github/workflows/ci.yml"
    )"
  fi
  if ! [[ "${expected_node_major}" =~ ^[0-9]+$ ]]; then
    expected_node_major="22"
  fi

  if [ "${node_major}" -lt "${expected_node_major}" ]; then
    echo "⚠️  Node.js ${node_version} detected; CI/release workflows run on Node ${expected_node_major}. Consider upgrading for maximum parity."
  fi
fi

# Rust
if ! command -v rustc &> /dev/null; then
  echo "⚠️  Rust not found. Install via: curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh"
else
  echo "✓ Rust: $(rustc --version)"
fi

# ============================================================================
# Optional: sccache for shared Rust compilation cache
# ============================================================================
echo ""
echo "=== Optional: Shared Compilation Cache ==="

if command -v sccache &> /dev/null; then
  echo "✓ sccache already installed"
else
  echo "ℹ️  sccache not installed. To enable shared compilation caching:"
  echo "   bash scripts/cargo_agent.sh install sccache"
  echo "   Then configure RUSTC_WRAPPER=sccache in agent-init.sh"
fi

# Create shared cache directory if desired
if [ ! -d /shared ]; then
  echo "ℹ️  Consider creating /shared directory for shared caches:"
  echo "   mkdir -p /shared/sccache /shared/npm-cache"
  echo "   chmod 1777 /shared /shared/sccache /shared/npm-cache"
fi

# ============================================================================
# Summary
# ============================================================================
echo ""
echo "╔════════════════════════════════════════════════════════════════╗"
echo "║  Setup Complete                                                 ║"
echo "╠════════════════════════════════════════════════════════════════╣"
echo "║  • inotify limits increased                                     ║"
echo "║  • File descriptor limits increased                             ║"
echo "║  • Canvas/graphics libraries installed                          ║"
echo "║  • Tauri dependencies installed                                 ║"
echo "║  • Headless browser dependencies installed                      ║"
echo "║  • Fonts installed                                              ║"
echo "╠════════════════════════════════════════════════════════════════╣"
echo "║  Next Steps:                                                    ║"
echo "║  1. Reboot or re-login for limits to take effect                ║"
echo "║  2. Ensure Node.js and Rust are installed                       ║"
echo "║  3. Agents should run: . scripts/agent-init.sh                  ║"
echo "╚════════════════════════════════════════════════════════════════╝"
