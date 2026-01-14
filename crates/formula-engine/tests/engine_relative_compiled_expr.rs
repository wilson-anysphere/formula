use formula_engine::eval::CellAddr;
use formula_engine::{Engine, PrecedentNode, Value};
use formula_model::EXCEL_MAX_ROWS;
use pretty_assertions::assert_eq;

#[test]
fn fill_down_formula_uses_per_cell_origin_and_shares_bytecode_program() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();

    // Filled pattern:
    // C1: =A1+B1
    // C2: =A2+B2
    engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A2+B2").unwrap();

    // These two formulas should share a single normalized bytecode program (interning) because
    // their references are the same once expressed as offsets from the formula origin cell.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let sheet1_id = engine.sheet_id("Sheet1").unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(30.0));

    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 1 }
            }
        ]
    );
    assert_eq!(
        engine.precedents("Sheet1", "C2").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 1, col: 1 }
            }
        ]
    );

    // Dependency graph regression guard: editing A1 should only dirty C1, not C2.
    engine.set_cell_value("Sheet1", "A1", 5.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "C1"));
    assert!(!engine.is_dirty("Sheet1", "C2"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(30.0));
}

#[test]
fn whole_column_references_shift_with_fill_and_share_bytecode_program() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();

    // Add a value in the destination column so SUM(B:B) differs from SUM(A:A).
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    // Equivalent to copying B1 one column to the right.
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(B:B)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let sheet1_id = engine.sheet_id("Sheet1").unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(103.0));

    let max_row = EXCEL_MAX_ROWS - 1;
    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::Range {
            sheet: sheet1_id,
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr {
                row: max_row,
                col: 0
            }
        }]
    );
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Range {
            sheet: sheet1_id,
            start: CellAddr { row: 0, col: 1 },
            end: CellAddr {
                row: max_row,
                col: 1
            }
        }]
    );

    // Dependency regression guard: editing B2 should dirty C1 (SUM(B:B)) but not B1 (SUM(A:A)).
    engine.set_cell_value("Sheet1", "B2", 200.0).unwrap();
    assert!(!engine.is_dirty("Sheet1", "B1"));
    assert!(engine.is_dirty("Sheet1", "C1"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(203.0));
}

#[test]
fn sheet_span_refs_shift_with_fill_and_share_bytecode_program() {
    let mut engine = Engine::new();

    // Create sheets up-front so 3D spans resolve deterministically.
    engine.ensure_sheet("Summary");
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    for (sheet, a1, a2) in [
        ("Sheet1", 1.0, 10.0),
        ("Sheet2", 2.0, 20.0),
        ("Sheet3", 3.0, 30.0),
    ] {
        engine.set_cell_value(sheet, "A1", a1).unwrap();
        engine.set_cell_value(sheet, "A2", a2).unwrap();
    }

    // Filled pattern on Summary sheet:
    // B1: =SUM(Sheet1:Sheet3!A1)
    // B2: =SUM(Sheet1:Sheet3!A2)
    engine
        .set_cell_formula("Summary", "B1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine
        .set_cell_formula("Summary", "B2", "=SUM(Sheet1:Sheet3!A2)")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let sheet1_id = engine.sheet_id("Sheet1").unwrap();
    let sheet2_id = engine.sheet_id("Sheet2").unwrap();
    let sheet3_id = engine.sheet_id("Sheet3").unwrap();
    assert_eq!(engine.get_cell_value("Summary", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Summary", "B2"), Value::Number(60.0));

    assert_eq!(
        engine.precedents("Summary", "B1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet2_id,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet3_id,
                addr: CellAddr { row: 0, col: 0 }
            }
        ]
    );
    assert_eq!(
        engine.precedents("Summary", "B2").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet2_id,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet3_id,
                addr: CellAddr { row: 1, col: 0 }
            }
        ]
    );

    // Dependency regression guard: editing Sheet2!A2 should only dirty Summary!B2.
    engine.set_cell_value("Sheet2", "A2", 99.0).unwrap();
    assert!(!engine.is_dirty("Summary", "B1"));
    assert!(engine.is_dirty("Summary", "B2"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Summary", "B2"), Value::Number(139.0));
}

#[test]
fn mixed_absolute_and_relative_refs_use_per_cell_origin_and_share_bytecode_program() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();

    // Filled pattern with mixed absolute/relative refs:
    // C1: =$A1+B$1
    // C2: =$A2+B$1
    //
    // Once expressed as offsets + absolute flags from the formula origin cell, these should share
    // a single normalized bytecode program.
    engine.set_cell_formula("Sheet1", "C1", "=$A1+B$1").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=$A2+B$1").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let sheet1_id = engine.sheet_id("Sheet1").unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(12.0));

    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 1 }
            }
        ]
    );
    // Precedents are returned sorted by `(row, col)`, so B1 comes before A2 here.
    assert_eq!(
        engine.precedents("Sheet1", "C2").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 0, col: 1 }
            },
            PrecedentNode::Cell {
                sheet: sheet1_id,
                addr: CellAddr { row: 1, col: 0 }
            }
        ]
    );

    // Dependency graph regression guard: editing A1 should only dirty C1.
    engine.set_cell_value("Sheet1", "A1", 5.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "C1"));
    assert!(!engine.is_dirty("Sheet1", "C2"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(12.0));

    // Editing the absolute-row input B$1 should dirty both formulas.
    engine.set_cell_value("Sheet1", "B1", 7.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "C1"));
    assert!(engine.is_dirty("Sheet1", "C2"));
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(17.0));
}
