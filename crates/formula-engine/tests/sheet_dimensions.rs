use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, PrecedentNode, Value};
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

#[test]
fn sheet_dimensions_affect_full_column_rows() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_000_000, EXCEL_MAX_COLS)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(A:A)")
        .unwrap();

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(2_000_000.0)
    );

    // Updating sheet dimensions should mark the cell dirty and recompute.
    engine
        .set_sheet_dimensions("Sheet1", 3_000_000, EXCEL_MAX_COLS)
        .unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(3_000_000.0)
    );
}

#[test]
fn set_sheet_dimensions_is_noop_when_dimensions_unchanged() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 10, 10).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(A:A)")
        .unwrap();
    engine.recalculate();

    assert!(!engine.is_dirty("Sheet1", "B1"));

    // Setting the same dimensions again should not dirty/recompile formulas.
    engine.set_sheet_dimensions("Sheet1", 10, 10).unwrap();
    assert!(!engine.is_dirty("Sheet1", "B1"));
}

#[test]
fn sheet_dimensions_affect_full_row_columns() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 10, 100).unwrap();
    // Avoid a circular reference: `1:1` includes row 1, so the formula can't be on row 1.
    engine
        .set_cell_formula("Sheet1", "B2", "=COLUMNS(1:1)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(100.0));

    // Updating sheet dimensions should mark the cell dirty and recompute.
    engine.set_sheet_dimensions("Sheet1", 10, 120).unwrap();
    assert!(engine.is_dirty("Sheet1", "B2"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(120.0));
}

#[test]
fn row_and_column_handle_row_and_column_refs_with_custom_dimensions() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_000_000, 100)
        .unwrap();

    // Avoid a circular ref: `5:7` includes row 5..7, so keep the formula outside those rows.
    engine
        .set_cell_formula("Sheet1", "J1", "=ROW(5:7)")
        .unwrap();
    // Avoid a circular ref: `D:F` includes column D..F, so keep the formula in another column.
    engine
        .set_cell_formula("Sheet1", "A10", "=COLUMN(D:F)")
        .unwrap();

    engine.recalculate();

    // ROW(5:7) -> {5;6;7}
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J3"), Value::Number(7.0));

    // COLUMN(D:F) -> {4,5,6}
    assert_eq!(engine.get_cell_value("Sheet1", "A10"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B10"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C10"), Value::Number(6.0));
}

#[test]
fn row_and_array_lift_return_spill_for_huge_whole_column_outputs() {
    // The engine supports sheets larger than Excel's default row count, but array-producing
    // functions should not attempt to materialize arbitrarily large arrays based solely on the
    // sheet's configured dimensions.
    //
    // Use a sheet height that exceeds the engine's materialization cap so functions like ROW(A:A)
    // and ABS(A:A) fail fast with `#SPILL!` instead of trying to allocate gigabytes.
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 5_000_001, EXCEL_MAX_COLS)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=ROW(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ABS(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=VALUE(A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=FILTER(A:A,A:A)")
        .unwrap();

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B4"),
        Value::Error(ErrorKind::Spill)
    );
}

#[test]
fn defined_name_whole_column_tracks_sheet_dimensions() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 10, 10).unwrap();
    engine
        .define_name(
            "MyCol",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A:A".to_string()),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(MyCol)")
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));

    // Growing the sheet should also grow whole-column references inside defined names.
    engine
        .set_sheet_dimensions("Sheet1", 2_000_000, 10)
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(2_000_000.0)
    );
}

#[test]
fn defined_name_whole_row_tracks_sheet_dimensions() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 100, 10).unwrap();
    engine
        .define_name(
            "MyRow",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!1:1".to_string()),
        )
        .unwrap();
    // Avoid a circular reference: `1:1` includes row 1, so the formula can't be on row 1.
    engine
        .set_cell_formula("Sheet1", "A2", "=COLUMNS(MyRow)")
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(10.0));

    // Growing the sheet should also grow whole-row references inside defined names.
    engine.set_sheet_dimensions("Sheet1", 100, 12).unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(12.0));
}

#[test]
fn indirect_whole_row_and_column_respect_sheet_dimensions() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_000_000, 10)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(INDIRECT(\"A:A\"))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=COLUMNS(INDIRECT(\"1:1\"))")
        .unwrap();

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(2_000_000.0)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(10.0));
}

#[test]
fn shrinking_sheet_dimensions_invalidates_bytecode_compiled_references() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=J1").unwrap();

    // The bytecode backend should be active for simple in-bounds formulas by default.
    assert!(
        engine.bytecode_program_count() > 0,
        "expected bytecode to compile for a simple cell ref"
    );

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);

    // Shrink the sheet so column J is out of bounds; the formula should now evaluate to #REF!.
    engine.set_sheet_dimensions("Sheet1", 10, 5).unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn precedents_clamp_whole_row_and_column_to_sheet_dimensions() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 100, 10).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    // Avoid a circular reference: `1:1` includes row 1, so keep the formula outside row 1.
    engine
        .set_cell_formula("Sheet1", "A2", "=SUM(1:1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::Range {
            sheet: 0,
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 99, col: 0 },
        }]
    );

    assert_eq!(
        engine.precedents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Range {
            sheet: 0,
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 0, col: 9 },
        }]
    );
}

#[test]
fn precedents_clamp_whole_row_and_column_from_defined_names_and_indirect() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 100, 10).unwrap();

    engine
        .define_name(
            "MyCol",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A:A".to_string()),
        )
        .unwrap();

    // A workbook-defined name that expands to a whole-column reference should clamp to sheet dims
    // in the auditing API.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyCol)")
        .unwrap();

    // INDIRECT dynamic dependencies should also clamp whole-column refs to sheet dims.
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(INDIRECT(\"A:A\"))")
        .unwrap();

    engine.recalculate();

    for addr in ["B1", "B2"] {
        assert_eq!(
            engine.precedents("Sheet1", addr).unwrap(),
            vec![PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 99, col: 0 },
            }],
            "unexpected precedents for {addr}"
        );
    }
}

#[test]
fn precedents_clamp_external_whole_row_and_column_to_excel_dimensions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=SUM([Book.xlsx]Sheet1!1:1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::ExternalRange {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr {
                row: EXCEL_MAX_ROWS.saturating_sub(1),
                col: 0
            },
        }]
    );

    assert_eq!(
        engine.precedents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::ExternalRange {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr {
                row: 0,
                col: EXCEL_MAX_COLS.saturating_sub(1)
            },
        }]
    );
}
