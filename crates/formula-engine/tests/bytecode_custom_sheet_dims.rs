#![cfg(not(target_arch = "wasm32"))]

use formula_engine::{metadata::FormatRun, BytecodeCompileReason, Engine, ErrorKind, Value};

#[test]
fn bytecode_custom_sheet_dims_whole_row_and_column_refs() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set sheet dimensions");

    // Populate a few values in column A (leave implicit blanks elsewhere).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A10", 3.0).unwrap();

    // Avoid circular references by keeping these formulas out of the referenced row/column.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ROWS(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=COLUMNS(1:1)")
        .unwrap();

    // These formulas should compile to bytecode (no AST fallback reasons).
    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );
    assert!(
        !report
            .iter()
            .any(|e| e.reason == BytecodeCompileReason::NonDefaultSheetDimensions),
        "NonDefaultSheetDimensions should not be reported after dimension-aware lowering"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(5.0));
}

#[test]
fn col_format_runs_sheet_dim_growth_invalidates_stale_bytecode_programs() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2, 2)
        .expect("set sheet dimensions");

    // Avoid circular references by keeping the formula out of the referenced column.
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(A:A)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formula to compile to bytecode; got: {report:?}"
    );

    // Grow sheet dimensions by applying formatting runs that extend beyond the current row count.
    engine
        .set_col_format_runs(
            "Sheet1",
            0,
            vec![FormatRun {
                start_row: 0,
                end_row_exclusive: 4,
                style_id: 1,
            }],
        )
        .unwrap();
    engine.recalculate_single_threaded();

    // The formula must observe the updated sheet dimensions (4 rows). This requires bumping
    // `sheet_dims_generation` so bytecode programs are treated as stale and fall back to AST
    // evaluation.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
}

#[test]
fn bytecode_custom_sheet_dims_row_spill_respects_sheet_row_count() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set sheet dimensions");

    // Place the formula outside the referenced column so it doesn't create a circular reference.
    // `ROW(A:A)` should spill a 10x1 array (one element per sheet row).
    engine
        .set_cell_formula("Sheet1", "B1", "=ROW(A:A)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();

    // Spill range should be exactly B1:B10 (10 rows).
    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start.row, 0);
    assert_eq!(start.col, 1);
    assert_eq!(end.row, 9);
    assert_eq!(end.col, 1);

    // Origin cell stores the top-left value of the spill.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B10"), Value::Number(10.0));
}

#[test]
fn bytecode_custom_sheet_dims_column_spill_respects_sheet_col_count() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set sheet dimensions");

    // `COLUMN(1:1)` should spill a 1x5 array (one element per sheet column).
    // Place it on row 2 so it doesn't reference itself (row 1 is referenced).
    engine
        .set_cell_formula("Sheet1", "A2", "=COLUMN(1:1)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();

    // Spill range should be exactly A2:E2 (5 cols).
    let (start, end) = engine.spill_range("Sheet1", "A2").expect("spill range");
    assert_eq!(start.row, 1);
    assert_eq!(start.col, 0);
    assert_eq!(end.row, 1);
    assert_eq!(end.col, 4);

    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(5.0));
}

#[test]
fn bytecode_custom_sheet_dims_row_spill_uses_referenced_sheet_dimensions() {
    let mut engine = Engine::new();
    // Make Sheet1 larger than Sheet2 so we can detect if the spill shape accidentally uses the
    // formula's sheet dimensions.
    engine
        .set_sheet_dimensions("Sheet1", 30, 10)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 7)
        .expect("set Sheet2 dimensions");

    engine
        .set_cell_formula("Sheet1", "B1", "=ROW(Sheet2!A:A)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();

    // Spill range should be based on Sheet2's row count (20), not Sheet1's (30).
    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start.row, 0);
    assert_eq!(start.col, 1);
    assert_eq!(end.row, 19);
    assert_eq!(end.col, 1);

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B20"), Value::Number(20.0));
}

#[test]
fn bytecode_custom_sheet_dims_column_spill_uses_referenced_sheet_dimensions() {
    let mut engine = Engine::new();
    // Make Sheet1 larger than Sheet2 so we can detect if the spill shape accidentally uses the
    // formula's sheet dimensions.
    engine
        .set_sheet_dimensions("Sheet1", 30, 10)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 7)
        .expect("set Sheet2 dimensions");

    engine
        .set_cell_formula("Sheet1", "A2", "=COLUMN(Sheet2!1:1)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();

    // Spill range should be based on Sheet2's column count (7), not Sheet1's (10).
    let (start, end) = engine.spill_range("Sheet1", "A2").expect("spill range");
    assert_eq!(start.row, 1);
    assert_eq!(start.col, 0);
    assert_eq!(end.row, 1);
    assert_eq!(end.col, 6);

    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(7.0));
}

#[test]
fn bytecode_custom_sheet_dims_column_sheet_prefixed_whole_column_does_not_trigger_range_cell_limit()
{
    // Regression test: `COLUMN(Sheet2!A:A)` produces a single value even if Sheet2 has a very large
    // row count, so the bytecode compiler should not reject it based on the referenced range's
    // cell count.
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 6_000_000, 10)
        .expect("set Sheet2 dimensions");

    engine
        .set_cell_formula("Sheet1", "B1", "=COLUMN(Sheet2!A:A)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formula to compile to bytecode; got report: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn bytecode_custom_sheet_dims_column_sheet_prefixed_column_range_does_not_trigger_range_cell_limit()
{
    // Regression test: `COLUMN(Sheet2!A:B)` returns a 1x2 vector even though `A:B` spans all rows.
    // The compiler should not reject it based on the dense range cell count.
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 6_000_000, 10)
        .expect("set Sheet2 dimensions");

    engine
        .set_cell_formula("Sheet1", "A1", "=COLUMN(Sheet2!A:B)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formula to compile to bytecode; got report: {report:?}"
    );

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start.row, 0);
    assert_eq!(start.col, 0);
    assert_eq!(end.row, 0);
    assert_eq!(end.col, 1);

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_custom_sheet_dims_row_sheet_prefixed_full_row_range_does_not_trigger_range_cell_limit()
{
    // Regression test: `ROW(Sheet2!1:1000)` spans an entire *row range* (all columns), but the
    // output is a 1000x1 vector (one element per row), not a dense 1000xN grid. The bytecode
    // compiler should therefore apply the same 1-D cell-count limit logic used for unprefixed
    // whole-row references.
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_000, 10)
        .expect("set Sheet1 dimensions");
    // Configure Sheet2 to be wide enough that the referenced row range would exceed the global
    // materialization cap if treated as a dense rectangle.
    engine
        .set_sheet_dimensions("Sheet2", 1_000, 6_000)
        .expect("set Sheet2 dimensions");

    engine
        .set_cell_formula("Sheet1", "B1", "=ROW(Sheet2!1:1000)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formula to compile to bytecode; got report: {report:?}"
    );

    engine.recalculate_single_threaded();

    // The output should be a 1000x1 spill containing {1;2;...;1000}.
    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start.row, 0);
    assert_eq!(start.col, 1);
    assert_eq!(end.row, 999);
    assert_eq!(end.col, 1);

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B500"),
        Value::Number(500.0)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1000"),
        Value::Number(1000.0)
    );
}

#[test]
fn bytecode_custom_sheet_dims_use_referenced_sheet_for_sheet_prefixed_whole_row_col_refs() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 7)
        .expect("set Sheet2 dimensions");

    // Put values on Sheet2 in rows/cols that would be out of bounds if the lowerer accidentally
    // used Sheet1's dimensions.
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A20", 3.0).unwrap();

    // Evaluate on Sheet1 but reference Sheet2's whole column/row.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet2!A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ROWS(Sheet2!A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=COLUMNS(Sheet2!1:1)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );
    assert!(
        !report
            .iter()
            .any(|e| e.reason == BytecodeCompileReason::NonDefaultSheetDimensions),
        "NonDefaultSheetDimensions should not be reported after dimension-aware lowering"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(7.0));
}

#[test]
fn bytecode_custom_sheet_dims_expand_3d_whole_column_per_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 5)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A20", 2.0).unwrap();

    // Put the formula on Sheet1 but avoid circular refs by keeping it out of column A.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet2!A:A)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn bytecode_custom_sheet_dims_expand_3d_whole_row_per_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 10, 7)
        .expect("set Sheet2 dimensions");

    // Row 1 spans different column counts on each sheet.
    engine.set_cell_value("Sheet1", "E1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "G1", 2.0).unwrap();

    // Put the formula on row 2 to avoid a circular reference (since `1:1` includes row 1).
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(Sheet1:Sheet2!1:1)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
}

#[test]
fn bytecode_custom_sheet_dims_sheet_prefixed_whole_column_can_reference_wider_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 10, 7)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet2", "G1", 2.0).unwrap();
    engine
        // Sheet1 only has 5 columns, but can still reference a wider Sheet2.
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet2!G:G)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_custom_sheet_dims_3d_whole_column_returns_ref_when_out_of_bounds_for_any_sheet() {
    let mut engine = Engine::new();
    // Sheet1 is narrower than Sheet2, so column G is out-of-bounds on Sheet1 but valid on Sheet2.
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 10, 7)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet2", "G1", 2.0).unwrap();
    engine
        // 3D refs should still compile to bytecode even if some areas are out-of-bounds; evaluation
        // should surface `#REF!`.
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet2!G:G)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn bytecode_custom_sheet_dims_sheet_prefixed_whole_row_can_reference_taller_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 5, 10)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 7, 10)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet2", "A7", 2.0).unwrap();
    engine
        // Sheet1 only has 5 rows, but can still reference a taller Sheet2.
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet2!7:7)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_custom_sheet_dims_3d_whole_row_returns_ref_when_out_of_bounds_for_any_sheet() {
    let mut engine = Engine::new();
    // Sheet1 is shorter than Sheet2, so row 7 is out-of-bounds on Sheet1 but valid on Sheet2.
    engine
        .set_sheet_dimensions("Sheet1", 5, 10)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 7, 10)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet2", "A7", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet2!7:7)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}
