//! Debugging and auditing support.
//!
//! The core `Engine` maintains a dependency graph for recalculation. This
//! module provides additional, on-demand tools used by the UX described in
//! `docs/12-ux-design.md`:
//! - Evaluation tracing with byte spans (for a step-through formula debugger).
//! - Trace data structures that can be serialized over IPC.

mod trace;

pub use trace::{DebugEvaluation, Span, TraceKind, TraceNode, TraceRef};

pub(crate) use trace::{evaluate_with_trace, parse_spanned_formula};
