//! Helpers for parsing canonical external workbook sheet keys.
//!
//! The engine represents external workbook references using a bracketed "external sheet key"
//! string such as `"[Book.xlsx]Sheet1"`. Centralizing parsing here ensures consistent validation
//! across the evaluator, engine, and debug tooling.

/// Split an external workbook key on the bracketed workbook boundary.
///
/// Accepts both single-sheet keys (`"[Book]Sheet"`) and 3D span keys (`"[Book]Start:End"`).
/// The returned `sheet_part` is everything after the closing bracket (it may contain `:`).
pub(crate) fn split_external_sheet_key_parts(key: &str) -> Option<(&str, &str)> {
    formula_model::external_refs::split_external_sheet_key_parts(key)
}

/// Parse a workbook-only external key in the canonical bracketed form: `"[Book]"`.
///
/// This is used for workbook-scoped external structured references like `[Book.xlsx]Table1[Col]`,
/// which lower to a `SheetReference::External("[Book.xlsx]")` key (no explicit sheet name).
///
/// Returns the workbook identifier slice (borrowed from `key`).
pub(crate) fn parse_external_workbook_key(key: &str) -> Option<&str> {
    formula_model::external_refs::parse_external_workbook_key(key)
}

/// Parse an external workbook sheet key in the canonical bracketed form: `"[Book]Sheet"`.
///
/// Returns the workbook name and sheet name slices (borrowed from `key`).
///
/// Notes:
/// - External 3D spans (`"[Book]Sheet1:Sheet3"`) are **not** accepted here; use
///   [`parse_external_span_key`] instead.
pub(crate) fn parse_external_key(key: &str) -> Option<(&str, &str)> {
    formula_model::external_refs::parse_external_key(key)
}

/// Parse an external workbook 3D span key in the canonical bracketed form: `"[Book]Start:End"`.
///
/// Returns the workbook name, start sheet, and end sheet slices (borrowed from `key`).
pub(crate) fn parse_external_span_key(key: &str) -> Option<(&str, &str, &str)> {
    formula_model::external_refs::parse_external_span_key(key)
}

/// Expand an external workbook 3D sheet span into per-sheet external keys.
///
/// Given:
/// - a `workbook` identifier (no surrounding brackets),
/// - span endpoints `start` and `end` (sheet names),
/// - the external workbook's sheet names in tab order (no `[workbook]` prefix),
///
/// returns canonical external sheet keys like `"[Book.xlsx]Sheet2"` for each sheet in the span.
///
/// Notes:
/// - Endpoint matching uses Excel-like Unicode-aware, case-insensitive comparison via
///   [`formula_model::sheet_name_eq_case_insensitive`].
/// - If either endpoint is missing from `sheet_names`, returns `None`.
pub(crate) fn expand_external_sheet_span_from_order(
    workbook: &str,
    start: &str,
    end: &str,
    sheet_names: &[String],
) -> Option<Vec<String>> {
    formula_model::external_refs::expand_external_sheet_span_from_order(
        workbook,
        start,
        end,
        sheet_names,
    )
}
