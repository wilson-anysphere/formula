# Formula: Next-Generation AI-Native Spreadsheet

> **⚠️ AGENT DEVELOPMENT CONSTRAINTS**: Before running any build commands, read [docs/99-agent-development-guide.md](./docs/99-agent-development-guide.md) and run `source scripts/agent-init.sh`. Memory limits are critical—this machine runs ~200 concurrent agents.

## Mission Statement

Build a spreadsheet application that achieves **100% Excel compatibility** while introducing **radical AI-native capabilities** and **modern architectural foundations** that make it objectively superior to Excel in every measurable dimension. This is not incremental improvement—it's a generational leap that will make users say "I can't imagine going back."

---

## Table of Contents

1. [Vision & Strategic Goals](#vision--strategic-goals)
2. [Architecture Overview](#architecture-overview)
3. [Core Systems](#core-systems)
4. [AI Integration Strategy](#ai-integration-strategy)
5. [Excel Compatibility Strategy](#excel-compatibility-strategy)
6. [User Experience Philosophy](#user-experience-philosophy)
7. [Cross-Cutting Concerns](#cross-cutting-concerns)
8. [Implementation Strategy](#implementation-strategy)
9. [Success Metrics](#success-metrics)
10. [Linked Detail Documents](#linked-detail-documents)

---

## Vision & Strategic Goals

### The Opportunity

Excel's architecture dates to the 1980s. Despite Microsoft's layering of modern features, fundamental constraints remain:

- **1,048,576 row limit** makes "real" data work impossible
- **Single-threaded VBA** blocks UI during macro execution
- **Calculation-blocking UI thread** creates lag during recalc
- **Binary VBA model** prevents modern development practices
- **Primitive version control** creates "v61_final_final_REAL.xlsx" chaos
- **Formula debugging** requires manual precedent tracing
- **Collaboration** breaks at scale with co-author limits

Microsoft Copilot for Excel demonstrates the opportunity but is **bolted onto legacy architecture**. Building AI-native from the ground up means the LLM has full context of the dependency graph, understands formula semantics, and can propose multi-step transformations that Excel cannot.

### Strategic Imperatives

1. **Zero-Compromise Excel Compatibility**: Every `.xlsx` file loads perfectly. Every formula works identically. Users can switch with zero friction.

2. **AI as Co-Pilot, Not Gimmick**: The AI is woven into formulas, data connectivity, and user assistance—not a chatbot sidebar. It supercharges experts while enabling novices.

3. **Performance That Scales**: Handle millions of rows with 60fps scrolling. Recalculate 100K cells in <100ms. Start in <1 second.

4. **Modern Foundation**: Git-like version control, relational data model, real-time collaboration, Python/TypeScript extensibility, plugin architecture.

5. **Win Power Users First**: Finance modelers, data analysts, and quant researchers are the most demanding users. Win them and the rest follow.

---

## Architecture Overview

### System Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  PRESENTATION LAYER                                                          │
│  ├── Canvas-Based Grid Renderer (60fps virtualized scrolling)               │
│  ├── Formula Bar with AI Autocomplete                                        │
│  ├── Command Palette (Cmd+K)                                                 │
│  ├── AI Chat Panel                                                           │
│  └── Overlay Elements (Selection, Cell Editor, Context Menus)               │
├─────────────────────────────────────────────────────────────────────────────┤
│  IPC BRIDGE (MessageChannel / Tauri Commands)                               │
├─────────────────────────────────────────────────────────────────────────────┤
│  APPLICATION LAYER                                                           │
│  ├── Document Controller (undo/redo, dirty tracking)                        │
│  ├── Selection Manager                                                       │
│  ├── Clipboard Handler (rich formats)                                        │
│  ├── AI Orchestrator (context management, tool calling)                     │
│  └── Collaboration Engine (CRDT/OT sync)                                    │
├─────────────────────────────────────────────────────────────────────────────┤
│  CORE ENGINE (Rust/WASM - runs in Worker thread)                            │
│  ├── Formula Parser (Chevrotain-style LL(k), A1/R1C1/Structured References) │
│  ├── Dependency Graph (incremental dirty marking, range nodes)              │
│  ├── Calculation Engine (multi-threaded, SIMD-optimized)                    │
│  ├── Function Library (500+ Excel-compatible functions)                     │
│  ├── Data Model (sparse storage, columnar compression)                      │
│  └── Format Engine (number formats, conditional formatting, styles)         │
├─────────────────────────────────────────────────────────────────────────────┤
│  DATA LAYER                                                                  │
│  ├── SQLite (CRDT-enabled for sync, auto-versioning)                        │
│  ├── VertiPaq-style Columnar Store (for Power Pivot data)                   │
│  ├── Format Converters (XLSX, XLSB, CSV, Parquet, Arrow)                    │
│  └── External Connectors (databases, APIs, cloud storage)                   │
├─────────────────────────────────────────────────────────────────────────────┤
│  AI LAYER                                                                    │
│  ├── Local Models (formula completion, code assistance)                      │
│  ├── Cloud Models (complex analysis, data fetching agents)                  │
│  ├── Context Manager (schema extraction, sampling, RAG)                     │
│  └── Tool Calling Interface (spreadsheet operations)                        │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Key Architectural Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Desktop Framework | **Tauri + Rust** | 10x smaller bundle than Electron, 4-8x less memory, <500ms startup, memory safety |
| Calculation Engine | **Rust compiled to WASM** | Near-native performance in browser context, multi-threaded via Web Workers |
| Grid Rendering | **Canvas-based** | Only viable approach for 60fps with millions of rows; DOM cannot scale |
| Collaboration | **CRDT (Yjs)** | Better offline support than OT; peer-to-peer capable; proven at scale |
| Storage | **SQLite + custom columnar** | Relational queries for metadata, columnar for large data ranges |
| AI Integration | **Multi-modal (local + cloud)** | Fast completion locally, powerful reasoning in cloud |

### Technology Stack

**Frontend:**
- TypeScript + React (UI components)
- Canvas 2D API (grid rendering)
- Yjs (CRDT for collaboration)
- TanStack Virtual (for non-grid virtualization)

**Backend (Tauri/Rust):**
- Rust (core calculation engine)
- wasm-bindgen (WASM interop)
- rayon (parallel computation)
- calamine/rust_xlsxwriter (XLSX I/O)
- SQLite (via rusqlite)
- tokio (async runtime)

**AI:**
- Local: Ollama or custom fine-tuned models
- Cloud: Claude/GPT-4 via API
- Embeddings: local sentence transformers for RAG

---

## Core Systems

### 1. Formula Engine
**Detail Document:** [docs/01-formula-engine.md](./docs/01-formula-engine.md)

The formula engine is the heart of any spreadsheet. Our implementation must:

- Parse **all Excel formula syntax** including A1, R1C1, structured references, dynamic arrays
- Implement **500+ functions** with identical behavior to Excel (including edge cases)
- Build and maintain **dependency graphs** with incremental updates
- Support **multi-threaded calculation** across independent branches
- Handle **volatile functions** (NOW, RAND, OFFSET, INDIRECT) correctly
- Implement **dynamic array spilling** with #SPILL! error handling
- Support **LAMBDA** and **LET** for user-defined functions

**Critical Specifications:**
- 8,192 character formula display limit (16,384 bytes tokenized)
- 64 levels of nested functions
- 255 arguments per function
- 15 significant digit IEEE 754 double precision
- 65,536 dependency threshold before full recalc mode

### 2. XLSX Compatibility Layer
**Detail Document:** [docs/02-xlsx-compatibility.md](./docs/02-xlsx-compatibility.md)

Perfect Excel compatibility is non-negotiable. This means:

- **Read/write all XLSX components** (worksheets, styles, charts, drawings, pivot tables, Power Query, VBA)
- **Preserve everything on round-trip** - no data loss, no formatting changes
- **Handle version differences** (Excel 2007 vs 2010+ conditional formatting, chart types)
- **Support XLSB** for faster loading of large files
- **Emulate Excel bugs** for compatibility (1900 leap year bug) with opt-out

**The Five Hardest Problems:**
1. Conditional formatting rules (version divergence)
2. Chart fidelity (DrawingML complexity)
3. Date systems (1900 vs 1904, Lotus bug)
4. Dynamic array function prefixes (`_xlfn.`)
5. VBA macro preservation

### 3. Rendering & UI
**Detail Document:** [docs/03-rendering-ui.md](./docs/03-rendering-ui.md)

Canvas-based rendering is mandatory for performance. Implementation requires:

- **Full grid on Canvas** - no DOM cells (Google Sheets approach)
- **Virtualized scrolling** with O(v) complexity (v = visible cells only)
- **Batch draw calls** to minimize GPU context switches
- **Device pixel ratio awareness** for crisp rendering on Retina displays
- **Smooth scrolling** to 33M pixel browser limits (~1M rows at 30px)
- **Overlay system** for selection, cell editing, and menus

### 4. Data Model & Storage
**Detail Document:** [docs/04-data-model-storage.md](./docs/04-data-model-storage.md)

Modern data handling that breaks Excel's limits:

- **Sparse HashMap storage** for cells (most spreadsheets are sparse)
- **Columnar compression** (VertiPaq-style) for large datasets
- **Relational tables** with enforced referential integrity
- **Rich data types** in cells (images, JSON, attachments)
- **No arbitrary row limits** - scale to 100M+ rows via streaming

### 5. AI Integration
**Detail Document:** [docs/05-ai-integration.md](./docs/05-ai-integration.md)

AI is not an add-on; it's woven into the fabric of the application:

| Mode | Trigger | Capability |
|------|---------|------------|
| **Tab Completion** | Typing in formula bar | Suggest formulas, values, completions based on context |
| **Inline Edit (Cmd+K)** | Selection + keyboard shortcut | Transform selected data via natural language |
| **Chat Panel** | Toggle panel | Ask questions about data, request analysis, explain formulas |
| **Cell Functions** | `=AI("prompt", range)` | AI as a formula that recalculates with data changes |
| **Agent Mode** | Autonomous toggle | Multi-step data gathering, cleaning, analysis workflows |

### 6. Collaboration
**Detail Document:** [docs/06-collaboration.md](./docs/06-collaboration.md)

Real-time collaboration that doesn't break at scale:

- **CRDT-based sync** (Yjs) for conflict-free concurrent editing
- **Presence indicators** showing who's editing what
- **Cell-level commenting** with threaded discussions
- **Version history** with named checkpoints and semantic diff
- **Offline-first** with automatic merge on reconnection
- **Branch-and-merge** for scenario analysis

### 7. Power Features
**Detail Document:** [docs/07-power-features.md](./docs/07-power-features.md)

The features that make Excel indispensable to power users:

- **Pivot Tables** with AI-assisted setup ("show sales by region and product")
- **Power Query equivalent** with query folding to push computation to sources
- **Data Model** with DAX-like calculated columns and measures
- **What-If Analysis** (Goal Seek, Solver, Scenario Manager)
- **Statistical tools** (regression, Monte Carlo with 100K+ iterations)
- **Native probability distributions** (20+ types)

### 8. Macro & Scripting Compatibility
**Detail Document:** [docs/08-macro-compatibility.md](./docs/08-macro-compatibility.md)

Legacy support with modern alternatives:

- **VBA preservation** - load and preserve vbaProject.bin
- **VBA execution** via interpreter or automated translation
- **Modern scripting** with Python and TypeScript
- **Macro recorder** that generates Python/TypeScript (not just VBA)
- **API parity** - everything VBA can do, modern scripts can do

---

## AI Integration Strategy

### The Autonomy Slider (Cursor's Paradigm)

Following Cursor's proven model, users control AI autonomy:

```
[Tab Completion] ←→ [Inline Assist] ←→ [Chat Panel] ←→ [Composer] ←→ [Full Agent]
     Low Autonomy                                                    High Autonomy
```

### AI Context Management

A 1M-row spreadsheet could exceed 100M tokens if naively serialized. Required strategies:

1. **Schema-First**: Send structure (headers, types, samples) not raw data
2. **Selective Sampling**: Random/stratified sample of rows for statistical questions
3. **RAG Over Cells**: Embed cells, retrieve relevant context via semantic search
4. **Hierarchical Summarization**: Pre-computed aggregates at sheet/region level
5. **Query-Aware Loading**: Dynamically load only ranges relevant to current query

**Context Budget Allocation (128K token model):**
```
System prompt:        ~2K tokens
Schema:               ~1K tokens
Sample data (100 rows): ~5K tokens
Conversation history: ~10K tokens
Retrieved context:    ~20K tokens
Reserved for output:  ~20K tokens
─────────────────────────────────
Available buffer:     ~70K tokens
```

### Tool Calling Schema

```json
{
  "tools": [
    {
      "name": "write_cell",
      "description": "Write a value or formula to a cell",
      "parameters": {
        "cell": "A1",
        "value": "=SUM(B1:B10)",
        "is_formula": true
      }
    },
    {
      "name": "apply_formula_column", 
      "description": "Apply a formula pattern to an entire column",
      "parameters": {
        "column": "C",
        "formula_template": "=A{row}*B{row}",
        "start_row": 2,
        "end_row": -1
      }
    },
    {
      "name": "create_pivot_table",
      "description": "Create a pivot table from a data range",
      "parameters": {
        "source_range": "A1:F1000",
        "rows": ["Region", "Product"],
        "values": [{"field": "Sales", "aggregation": "SUM"}],
        "destination": "Sheet2!A1"
      }
    },
    {
      "name": "detect_anomalies",
      "description": "Find outliers in a data range",
      "parameters": {
        "range": "D2:D100",
        "method": "zscore",
        "threshold": 2.5
      }
    },
    {
      "name": "fetch_external_data",
      "description": "Retrieve data from external source",
      "parameters": {
        "source_type": "api",
        "url": "https://api.example.com/data",
        "destination": "Sheet1!A1",
        "transform": "json_to_table"
      }
    }
  ]
}
```

### AI Safety and Verification

AI outputs must be verifiable:

1. **Computation over hallucination**: For factual queries about data, AI should use formulas to compute answers rather than guessing
2. **Source attribution**: External data fetches include citation metadata
3. **Preview before apply**: Major changes show diff preview for user approval
4. **Confidence indicators**: Flag uncertain outputs
5. **Audit trail**: Log all AI actions for review

---

## Excel Compatibility Strategy

### Compatibility Levels

| Level | Description | Target |
|-------|-------------|--------|
| **L1: Read** | File opens, all data visible | 100% |
| **L2: Calculate** | All formulas produce correct results | 99.9% |
| **L3: Render** | Visual appearance matches Excel | 98% |
| **L4: Round-trip** | Save and reopen in Excel with no changes | 97% |
| **L5: Execute** | VBA macros run correctly | 90% (stretch) |

### Function Compatibility Matrix

| Category | Functions | Priority | Complexity |
|----------|-----------|----------|------------|
| Math & Trig | SUM, AVERAGE, ROUND, etc. | P0 | Low |
| Lookup | VLOOKUP, INDEX, MATCH, XLOOKUP | P0 | Medium |
| Text | CONCATENATE, LEFT, RIGHT, FIND | P0 | Low |
| Logical | IF, AND, OR, IFS, SWITCH | P0 | Low |
| Date/Time | DATE, TODAY, NETWORKDAYS | P0 | Medium |
| Statistical | STDEV, CORREL, LINEST | P1 | Medium |
| Financial | NPV, IRR, PMT, XNPV | P1 | High |
| Dynamic Arrays | FILTER, SORT, UNIQUE, SEQUENCE | P0 | High |
| Information | ISBLANK, ISERROR, TYPE | P0 | Low |
| Engineering | COMPLEX, IMSUM, CONVERT | P2 | Medium |
| Cube | CUBEVALUE, CUBEMEMBER | P2 | High |
| Database | DSUM, DCOUNT, DGET | P2 | Medium |

### The xlsx Preservation Strategy

Use **Markup Compatibility (MC) namespace** for forward compatibility:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="x14">
    <!-- Excel 2010+ feature -->
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for older apps -->
  </mc:Fallback>
</mc:AlternateContent>
```

**Critical rules:**
- Always store both formula text (`f` attribute) AND cached value (`v` attribute)
- Preserve relationship IDs exactly—never regenerate
- Store `_xlfn.` prefixes for newer functions
- Maintain `calcChain.xml` for calculation order hints

---

## User Experience Philosophy

### Power User First

The interface should make experts faster, not slower:

1. **Keyboard-driven**: Every action accessible via keyboard
2. **Command Palette (Cmd+K)**: VS Code-style command search
3. **Vim-mode optional**: Modal editing for those who want it
4. **No modal dialogs**: Prefer inline editing and panels
5. **Persistent layouts**: Remember window arrangements, zoom levels, scroll positions

### Pain Points We Must Solve

| Pain Point | Excel's Approach | Our Solution |
|------------|------------------|--------------|
| Version chaos | "Save As" creates copies | Git-like branches, named checkpoints |
| Formula debugging | Evaluate Formula dialog | Always-on step debugger, hover previews |
| Collaboration limits | Co-author limits, shared filtering | CRDT sync, independent views |
| Large data | 1M row limit, slow scrolling | No limits, virtualized 60fps rendering |
| Learning curve | Help menu, web search | Integrated AI tutor, contextual examples |

### Formula Debugging UX

Based on VL/HCC 2023 research (FxD paper), implement:

1. **Always-on debugging**: Execution steps visible without initialization
2. **Step-wise evaluation**: Show all steps at once with collapsible precedents
3. **Sub-formula trace coloring**: Hover highlights expression + result
4. **Information inspector**: Range previews, cell provenance "pills"

### Semantic Version Control

What Git is to code, we are to spreadsheets:

- **Cell-by-cell diff** with color coding (green=added, red=removed, yellow=modified)
- **Formula diff** showing specific changes, not just "formula changed"
- **Side-by-side comparison** of conflicting cells with merge preview
- **Named checkpoints** ("Q3 Budget Approved") with annotations
- **Scenario branches** for what-if analysis without file duplication

---

## Cross-Cutting Concerns

### Performance Targets

| Metric | Target | Measurement |
|--------|--------|-------------|
| Cold start | <1 second | Time to interactive grid |
| Scroll FPS | 60fps | With 1M+ rows visible range |
| Recalculation | <100ms | 100K cell dependency chain |
| File open | <3 seconds | 100MB xlsx file |
| Memory | <500MB | For 100MB xlsx loaded |
| AI response | <2 seconds | Tab completion suggestions |

### Security Model

1. **Sandboxed execution**: Scripts run in isolated contexts
2. **Permission system**: Explicit grants for file system, network, clipboard
3. **Data encryption**: At-rest and in-transit encryption
4. **Audit logging**: All changes tracked with user attribution
5. **Cell-level permissions**: Granular access control for enterprise

### Accessibility

- **Screen reader support**: Full ARIA labels, logical navigation
- **Keyboard navigation**: Complete keyboard accessibility
- **High contrast mode**: System preference respected
- **Font scaling**: Respect system font size preferences
- **Reduced motion**: Disable animations when requested

### Internationalization

- **Formula localization**: Support localized function names (SUMME for German SUM)
- **Number formats**: Locale-aware decimal/thousands separators
- **Date formats**: Regional date format support
- **RTL support**: Right-to-left language layouts
- **Translation infrastructure**: All UI strings externalizable

---

## Implementation Strategy

### Phase Overview

| Phase | Duration | Focus | Key Deliverables |
|-------|----------|-------|------------------|
| **Phase 1** | Months 1-6 | Core Engine | Formula parser, dependency graph, basic grid rendering, xlsx read |
| **Phase 2** | Months 4-8 | Desktop Shell | Tauri wrapper, Rust engine integration, SQLite storage, native OS integration |
| **Phase 3** | Months 6-12 | AI Integration | Tab completion, chat panel, tool calling, formula assistance |
| **Phase 4** | Months 10-18 | Power Features | Power Query equivalent, Data Model, version control, agent mode, plugins |
| **Phase 5** | Months 16-24 | Polish & Scale | Performance optimization, enterprise features, advanced collaboration |

### Parallel Workstreams

These workstreams can proceed in parallel with appropriate coordination:

```
Workstream A: Core Engine (Rust)
├── Formula Parser
├── Dependency Graph  
├── Function Library
├── Calculation Engine
└── Data Model

Workstream B: UI/UX (TypeScript/React)
├── Canvas Grid Renderer
├── Formula Bar & Editing
├── Command Palette
├── Panels & Dialogs
└── Theming System

Workstream C: File I/O
├── XLSX Reader
├── XLSX Writer
├── XLSB Support
├── CSV/Parquet
└── Round-trip Testing

Workstream D: AI Integration
├── Context Management
├── Local Model Integration
├── Cloud API Integration
├── Tool Calling Framework
└── Safety & Verification

Workstream E: Collaboration
├── CRDT Implementation
├── Sync Protocol
├── Presence System
├── Version History
└── Conflict Resolution

Workstream F: Platform
├── Tauri Shell
├── Auto-update
├── Crash Reporting
├── Analytics
└── Installer/Distribution
```

### Risk Mitigation

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| xlsx compatibility gaps | High | Critical | Extensive test suite against real-world files; community bug reporting |
| Performance not meeting targets | Medium | High | Continuous benchmarking; profiling-driven optimization; fallback strategies |
| AI accuracy concerns | Medium | Medium | Verification layer; preview before apply; confidence indicators |
| VBA execution complexity | High | Medium | Prioritize preservation over execution; offer migration tools |
| Scope creep | High | Medium | Strict feature prioritization; MVP discipline |

---

## Success Metrics

### Technical Metrics

- **Formula compatibility**: 99.9% of Excel formulas produce identical results
- **File compatibility**: 99% of xlsx files round-trip without user-visible changes
- **Performance**: Meet all targets in Performance Targets section
- **Crash rate**: <0.1% of sessions experience crashes
- **AI accuracy**: >95% user acceptance of AI suggestions

### User Metrics

- **Time to productivity**: New users complete tutorial in <10 minutes
- **Power user efficiency**: Tasks completed 20% faster than Excel
- **Net Promoter Score**: >50
- **Daily Active Users**: Retention curve flattens above 40% at day 30
- **Feature adoption**: >30% of users engage with AI features weekly

### Business Metrics

- **Excel migration rate**: >80% of xlsx files imported successfully on first try
- **Support ticket volume**: <5 tickets per 1000 users per month
- **Enterprise adoption**: SOC2/ISO27001 certification within 18 months

---

## Linked Detail Documents

### Core Systems

| Document | Description |
|----------|-------------|
| [01-formula-engine.md](./docs/01-formula-engine.md) | Formula parsing, evaluation, dependency tracking, function library |
| [02-xlsx-compatibility.md](./docs/02-xlsx-compatibility.md) | File format handling, preservation strategy, compatibility testing |
| [03-rendering-ui.md](./docs/03-rendering-ui.md) | Canvas rendering, virtualization, scroll performance, overlays |
| [17-charts.md](./docs/17-charts.md) | DrawingML/ChartEx chart parsing, rendering fidelity, and round-trip preservation |
| [04-data-model-storage.md](./docs/04-data-model-storage.md) | Cell storage, columnar compression, relational tables, rich types |
| [05-ai-integration.md](./docs/05-ai-integration.md) | AI modes, context management, tool calling, safety |
| [06-collaboration.md](./docs/06-collaboration.md) | CRDT sync, presence, versioning, conflict resolution |
| [07-power-features.md](./docs/07-power-features.md) | Pivot tables, Power Query, Data Model, analysis tools |
| [08-macro-compatibility.md](./docs/08-macro-compatibility.md) | VBA handling, modern scripting, automation |

### Supporting Systems

| Document | Description |
|----------|-------------|
| [09-security-enterprise.md](./docs/09-security-enterprise.md) | Permissions, encryption, audit, compliance |
| [10-extensibility.md](./docs/10-extensibility.md) | Plugin architecture, custom functions, API |
| [11-desktop-shell.md](./docs/11-desktop-shell.md) | Tauri integration, native features, distribution |
| [12-ux-design.md](./docs/12-ux-design.md) | Interface design, keyboard shortcuts, command palette |

### Reference

| Document | Description |
|----------|-------------|
| [13-testing-validation.md](./docs/13-testing-validation.md) | Test strategy, compatibility validation, benchmarking |
| [14-competitive-analysis.md](./docs/14-competitive-analysis.md) | Excel, Google Sheets, Notion, Airtable, others |
| [15-excel-feature-parity.md](./docs/15-excel-feature-parity.md) | Complete feature checklist with priority |
| [16-performance-targets.md](./docs/16-performance-targets.md) | Detailed benchmarks and optimization strategies |

### Agent Development

| Document | Description |
|----------|-------------|
| [99-agent-development-guide.md](./docs/99-agent-development-guide.md) | **CRITICAL**: Memory limits, parallelism, headless setup for multi-agent development |

---

## Work Breakdown Suggestions

The following are suggestions for organizing work across multiple agents/teams. The actual structure should be determined based on available resources and coordination preferences.

### Suggested Team Boundaries

**Option A: By System**
- Team per major system (Formula Engine, UI, AI, etc.)
- Clear interfaces between teams
- Risk: Integration complexity

**Option B: By Feature Slice**
- Cross-functional teams owning end-to-end features
- Better integration, harder coordination
- Risk: Duplicated infrastructure work

**Option C: Hybrid**
- Platform team builds shared infrastructure
- Feature teams build on platform
- Balance of integration and specialization

### Coordination Points

Regardless of organization, these integration points require careful coordination:

1. **Formula Engine ↔ UI**: Cell rendering, formula bar, error display
2. **Data Model ↔ File I/O**: Serialization format, lazy loading
3. **AI ↔ Core Engine**: Context extraction, tool execution
4. **Collaboration ↔ Data Model**: CRDT operations, conflict resolution
5. **All Systems ↔ Performance**: Shared benchmarking, profiling tools

### Documentation Standards

All development should maintain:

1. **API documentation**: Generated from code annotations
2. **Architecture Decision Records (ADRs)**: For significant decisions
3. **Test coverage reports**: Automated and visible
4. **Performance dashboards**: Continuous benchmarking
5. **Compatibility reports**: xlsx test suite results

---

## Conclusion

This plan represents a comprehensive blueprint for building a next-generation spreadsheet that doesn't just compete with Excel—it leapfrogs it. The combination of **perfect compatibility**, **AI-native design**, **modern architecture**, and **power-user focus** creates a product that can genuinely transform how people work with data.

The challenge is immense. Excel has 40 years of accumulated features and the trust of over a billion users. But its architectural constraints create an opportunity for a clean-slate implementation that can deliver capabilities Excel simply cannot achieve.

Success requires unwavering commitment to:
1. **Compatibility first**: Users must trust us with their files
2. **Performance always**: Speed is a feature
3. **AI thoughtfully**: Enhance, don't replace, human judgment
4. **Power users respected**: Don't dumb down for lowest common denominator

The linked documents provide the detailed specifications needed to execute this vision. Each represents a significant undertaking; together, they form the blueprint for the spreadsheet of the future.
