# Architecture Decision Records (ADRs)

This directory contains short, implementation-backed decisions that clarify **ownership boundaries**
and prevent architectural drift.

When adding a new ADR:

- pick the next available `ADR-XXXX` number (4 digits),
- include `Status` + `Date`,
- prefer linking to concrete code paths so future contributors can validate assumptions.

## Index

- [ADR-0001: Platform target](./ADR-0001-platform-target.md)
- [ADR-0002: Engine execution model (Tauri invoke vs WASM Worker)](./ADR-0002-engine-execution-model.md)
- [ADR-0003: Crypto envelope KMS](./ADR-0003-crypto-envelope-kms.md)
- [ADR-0003: Engine protocol parity](./ADR-0003-engine-protocol-parity.md)
- [ADR-0004: Collaboration semantics for sheet view state and undo](./ADR-0004-collab-sheet-view-and-undo.md)
- [ADR-0005: PivotTables ownership and data flow across crates](./ADR-0005-pivot-tables-ownership-and-data-flow.md)

