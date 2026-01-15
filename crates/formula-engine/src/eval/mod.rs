mod address;
mod ast;
mod compiler;
mod evaluator;
mod parser;

/// Internal sentinel prefix used to track omitted LAMBDA parameters.
///
/// The leading NUL character ensures the key cannot be referenced by user formulas.
pub(crate) const LAMBDA_OMITTED_PREFIX: &str = "\u{0}LAMBDA_OMITTED:";

pub use address::{parse_a1, AddressParseError, CellAddr};
pub use ast::{
    BinaryOp, CellRef, CompareOp, CompiledExpr, Expr, NameRef, ParsedExpr, PostfixOp, RangeRef,
    Ref, SheetReference, StructuredRefExpr, UnaryOp,
};
pub use compiler::{compile_canonical_expr, lower_ast, lower_expr};
pub(crate) use evaluator::MAX_MATERIALIZED_ARRAY_CELLS;
pub(crate) use evaluator::{
    is_valid_external_single_sheet_key, split_external_sheet_key_parts, split_external_sheet_span_key,
};
pub use evaluator::{
    DependencyTrace, EvalContext, Evaluator, RecalcContext, ResolvedName, ValueResolver,
};
pub use parser::{FormulaParseError, Parser};

#[cfg(test)]
mod tests {
    use super::{
        is_valid_external_single_sheet_key, split_external_sheet_key_parts,
        split_external_sheet_span_key,
    };

    #[test]
    fn split_external_sheet_key_parses_basic() {
        assert_eq!(
            split_external_sheet_key_parts("[Book.xlsx]Sheet1"),
            Some(("Book.xlsx", "Sheet1"))
        );
    }

    #[test]
    fn split_external_sheet_key_uses_last_bracket_when_workbook_contains_brackets() {
        // Workbook ids can include path prefixes from quoted external references like
        // `'C:\\[foo]\\[Book.xlsx]Sheet1'!A1`. The parser folds the path prefix into the workbook
        // id (dropping the `[...]` around the workbook name), producing a canonical external key
        // like:
        //   workbook: `C:\[foo]\Book.xlsx`
        //   key:      `[C:\[foo]\Book.xlsx]Sheet1`
        //
        // The folder name `[foo]` introduces an interior `]`, so we must split on the last `]`.
        let key = "[C:\\[foo]\\Book.xlsx]Sheet1";
        assert_eq!(
            split_external_sheet_key_parts(key),
            Some(("C:\\[foo]\\Book.xlsx", "Sheet1"))
        );
    }

    #[test]
    fn split_external_sheet_key_rejects_missing_components() {
        assert_eq!(split_external_sheet_key_parts("Book.xlsx]Sheet1"), None);
        assert_eq!(split_external_sheet_key_parts("[Book.xlsxSheet1"), None);
        assert_eq!(split_external_sheet_key_parts("[]Sheet1"), None);
        assert_eq!(split_external_sheet_key_parts("[Book.xlsx]"), None);
    }

    #[test]
    fn split_external_sheet_span_key_parses_sheet_ranges() {
        assert_eq!(
            split_external_sheet_span_key("[Book.xlsx]Sheet1:Sheet3"),
            Some(("Book.xlsx", "Sheet1", "Sheet3"))
        );
    }

    #[test]
    fn split_external_sheet_span_key_rejects_invalid_ranges() {
        assert_eq!(split_external_sheet_span_key("[Book.xlsx]Sheet1"), None);
        assert_eq!(split_external_sheet_span_key("[Book.xlsx]:Sheet3"), None);
        assert_eq!(split_external_sheet_span_key("[Book.xlsx]Sheet1:"), None);
    }

    #[test]
    fn is_valid_external_single_sheet_key_is_strict_about_spans() {
        assert!(is_valid_external_single_sheet_key("[Book.xlsx]Sheet1"));
        assert!(!is_valid_external_single_sheet_key("[Book.xlsx]Sheet1:Sheet3"));
        assert!(!is_valid_external_single_sheet_key("not a key"));
    }
}
