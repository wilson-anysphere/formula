# Formula: Agent Guidelines

> **This is a Cursor product.** All AI features are powered by Cursor's backend—no local models, no API keys, no provider configuration.

---

## ⛔ READ THIS FIRST ⛔

**Before doing ANYTHING, understand these rules. They are non-negotiable.**

1. **Memory limits are critical** — see [Development Environment](#development-environment)
2. **Use wrapper scripts for cargo** — `bash scripts/cargo_agent.sh`, never bare `cargo`
3. **Excel compatibility is non-negotiable** — 100% formula behavioral compatibility
4. **Cursor product identity** — all AI goes through Cursor servers, period

---

## Workstream Instructions

**You are assigned to ONE workstream. Read your workstream file and follow it.**

| Workstream | File | Focus |
|------------|------|-------|
| **A: Core Engine** | [`instructions/core-engine.md`](./instructions/core-engine.md) | Formula parser, dependency graph, functions, Rust/WASM |
| **B: UI/UX** | [`instructions/ui.md`](./instructions/ui.md) | Canvas grid, formula bar, panels, React/TypeScript |
| **C: File I/O** | [`instructions/file-io.md`](./instructions/file-io.md) | XLSX/XLSB reader/writer, round-trip preservation |
| **D: AI Integration** | [`instructions/ai.md`](./instructions/ai.md) | Context management, tool calling, Cursor backend |
| **E: Collaboration** | [`instructions/collaboration.md`](./instructions/collaboration.md) | CRDT sync, presence, version history, Yjs |
| **F: Platform** | [`instructions/platform.md`](./instructions/platform.md) | Tauri shell, native integration, distribution |

**All workstream files reference back to this document.** This is the shared foundation.

---

## Mission Statement

Build a spreadsheet that achieves **100% Excel compatibility** while introducing **radical AI-native capabilities** and **modern architectural foundations**.

**Strategic Imperatives:**

1. **Zero-Compromise Excel Compatibility** — every `.xlsx` loads perfectly
2. **AI as Co-Pilot, Not Gimmick** — woven into formulas, not a chatbot sidebar
3. **Performance That Scales** — 60fps with millions of rows
4. **Modern Foundation** — Git-like versioning, real-time collaboration, Python/TypeScript scripting
5. **Win Power Users First** — finance modelers, data analysts, quant researchers

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  PRESENTATION LAYER (TypeScript/React)                                       │
│  ├── Canvas-Based Grid Renderer (60fps virtualized scrolling)               │
│  ├── Formula Bar with AI Autocomplete                                        │
│  ├── Command Palette (Cmd/Ctrl+Shift+P)                                      │
│  └── AI Chat Panel                                                           │
├─────────────────────────────────────────────────────────────────────────────┤
│  CORE ENGINE (Rust/WASM - runs in Worker thread)                            │
│  ├── Formula Parser (A1/R1C1/Structured References)                         │
│  ├── Dependency Graph (incremental dirty marking)                           │
│  ├── Calculation Engine (multi-threaded, SIMD)                              │
│  └── Function Library (500+ Excel-compatible)                               │
├─────────────────────────────────────────────────────────────────────────────┤
│  DATA LAYER                                                                  │
│  ├── SQLite (CRDT-enabled for sync)                                         │
│  ├── Format Converters (XLSX, XLSB, CSV, Parquet)                           │
│  └── Yjs (CRDT for collaboration)                                           │
├─────────────────────────────────────────────────────────────────────────────┤
│  AI LAYER (Cursor-managed)                                                   │
│  ├── Cursor AI Backend (all inference)                                      │
│  ├── Context Manager (schema extraction, sampling, RAG)                     │
│  └── Tool Calling Interface                                                 │
└─────────────────────────────────────────────────────────────────────────────┘
```

**Key Decisions:**

| Component | Choice | Rationale |
|-----------|--------|-----------|
| Desktop Framework | Tauri + Rust | 10x smaller than Electron, memory safety |
| Calculation Engine | Rust → WASM | Near-native performance in browser |
| Grid Rendering | Canvas-based | Only way to hit 60fps with millions of rows |
| Collaboration | CRDT (Yjs) | Better offline support than OT |
| AI | Cursor Backend | Cursor controls all AI inference |

---

## Development Environment

### ⚠️ Memory Management (CRITICAL)

This machine runs **~200 concurrent agents**. Memory is the constraint.

```
Total RAM:     1,500 GB
Per-agent:        ~7 GB soft target
```

### Memory Limits by Operation

| Operation | Expected Peak | Limit |
|-----------|---------------|-------|
| Node.js process | 512MB-2GB | 4GB |
| Rust compilation | 2-8GB | 14GB |
| TypeScript check | 500MB-2GB | 2GB |
| Total concurrent | - | 14GB |

### Required: Use Wrapper Scripts

**MANDATORY for all cargo commands:**

```bash
# CORRECT:
bash scripts/cargo_agent.sh build --release
bash scripts/cargo_agent.sh test -p formula-engine
bash scripts/cargo_agent.sh check

# WRONG - can exhaust RAM:
cargo build
cargo test
```

### Environment Setup

```bash
# Optional: override default Cargo parallelism for the session
export FORMULA_CARGO_JOBS=8

# Initialize safe defaults (do this first)
. scripts/agent-init.sh
```

### Helper Scripts

| Script | Purpose |
|--------|---------|
| `scripts/agent-init.sh` | Source at session start |
| `scripts/cargo_agent.sh` | **Use for ALL cargo commands** |
| `scripts/run_limited.sh` | Run any command with memory cap |
| `scripts/check-memory.sh` | Show memory status |

**Full details:** [`docs/99-agent-development-guide.md`](./docs/99-agent-development-guide.md)

---

## Design Philosophy

### Excel Functionality + Cursor Polish

**This is the formula: HIGH functionality + CLEAN styling.**

```
FUNCTIONALITY (# of Excel buttons/features)
        ↑
  LOW   │   HIGH
────────┼────────→ STYLING (clean vs bloated)
        │
```

**CORRECT:** All Excel features, styled with Cursor's clean aesthetic.

**FAILURE MODES:**
- ❌ Microsoft Bureaucracy: bloated styling
- ❌ Stripped Minimal: removed features to "look minimal"
- ❌ AI Slop: giant icons, gradient backgrounds, chat-bot aesthetics

### Design Rules

| ✅ DO | ❌ DON'T |
|-------|----------|
| Light mode | Dark mode by default |
| Full ribbon interface | Hamburger menus |
| Sheet tabs at bottom | Hide sheet navigation |
| Visible formula bar | Collapsed inputs |
| Monospace for cells | Sans for numbers |
| Professional blue accent | Purple AI gradients |

### Mockups

```bash
mockups/spreadsheet-main.html    # Main app layout
mockups/ai-agent-mode.html       # Agent execution view
mockups/command-palette.html     # Quick command search
mockups/README.md                # Full design system
```

> ⚠️ **Mockups are directional, not literal.**
> Use them for vision and design language. Apply judgment—add polish, fix inconsistencies. Follow **principles** over exact pixels.

---

## Cross-Cutting Concerns

### Performance Targets

| Metric | Target |
|--------|--------|
| Cold start | <1 second |
| Scroll FPS | 60fps with 1M+ rows |
| Recalculation | <100ms for 100K cell chain |
| File open | <3 seconds for 100MB xlsx |
| Memory | <500MB for 100MB xlsx |
| AI response | <2 seconds for tab completion |

### Excel Compatibility Levels

| Level | Description | Target |
|-------|-------------|--------|
| L1: Read | File opens, all data visible | 100% |
| L2: Calculate | All formulas produce correct results | 99.9% |
| L3: Render | Visual appearance matches Excel | 98% |
| L4: Round-trip | Save and reopen without changes | 97% |

### Security

1. **Sandboxed execution** — scripts run isolated
2. **Permission system** — explicit grants for file/network/clipboard
3. **Data encryption** — at-rest and in-transit
4. **Audit logging** — all changes tracked

### Accessibility

- Screen reader support (ARIA labels)
- Complete keyboard navigation
- High contrast mode
- Font scaling
- Reduced motion

---

## Coordination Points

Regardless of workstream, these integration points require coordination:

1. **Formula Engine ↔ UI**: Cell rendering, formula bar, error display
2. **Data Model ↔ File I/O**: Serialization format, lazy loading
3. **AI ↔ Core Engine**: Context extraction, tool execution
4. **Collaboration ↔ Data Model**: CRDT operations, conflict resolution
5. **All Systems ↔ Performance**: Shared benchmarking, profiling

---

## Documentation

### Core Systems

| Document | Description |
|----------|-------------|
| [`docs/01-formula-engine.md`](./docs/01-formula-engine.md) | Formula parsing, evaluation, functions |
| [`docs/02-xlsx-compatibility.md`](./docs/02-xlsx-compatibility.md) | File format handling, preservation |
| [`docs/03-rendering-ui.md`](./docs/03-rendering-ui.md) | Canvas rendering, virtualization |
| [`docs/04-data-model-storage.md`](./docs/04-data-model-storage.md) | Cell storage, compression |
| [`docs/05-ai-integration.md`](./docs/05-ai-integration.md) | AI modes, context management |
| [`docs/06-collaboration.md`](./docs/06-collaboration.md) | CRDT sync, presence, versioning |

### Supporting

| Document | Description |
|----------|-------------|
| [`docs/11-desktop-shell.md`](./docs/11-desktop-shell.md) | Tauri integration |
| [`docs/12-ux-design.md`](./docs/12-ux-design.md) | Interface design, shortcuts |
| [`docs/99-agent-development-guide.md`](./docs/99-agent-development-guide.md) | **CRITICAL**: Memory limits, build setup |

---

## Quick Reference

### Build Commands

```bash
# Setup
. scripts/agent-init.sh
pnpm install

# Build WASM engine
pnpm build:wasm

# Run desktop app
pnpm dev:desktop

# Rust (ALWAYS use wrapper)
bash scripts/cargo_agent.sh build --release
bash scripts/cargo_agent.sh test -p formula-engine

# Tests
pnpm test
pnpm test:node
```

### DANGEROUS Commands (NEVER USE)

```bash
cargo build                    # No memory limit!
cargo test                     # No memory limit!
cargo build -j$(nproc)        # 192 parallel rustc = OOM
pkill cargo                    # Kills OTHER agents
killall rustc                  # Kills OTHER agents
```

---

## Remember

1. **Read your workstream file** in `instructions/`
2. **Use `cargo_agent.sh`** for all Rust commands
3. **Excel compatibility is non-negotiable**
4. **All AI goes through Cursor servers**
5. **Mockups are direction, not specification**
6. **HIGH functionality + CLEAN styling**
