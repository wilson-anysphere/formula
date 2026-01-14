use formula_engine::eval::CellAddr;
use formula_engine::value::EntityValue;
use formula_engine::{Engine, PrecedentNode, Value};
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
use std::collections::HashMap;

#[test]
fn precedents_and_dependents_queries_work() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_formula("Sheet1", "A2", "=A1*2").unwrap();
    engine.set_cell_formula("Sheet1", "A3", "=A2+1").unwrap();
    engine.recalculate();

    let p_a3_direct = engine.precedents("Sheet1", "A3").unwrap();
    assert_eq!(
        p_a3_direct,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 1, col: 0 }
        }]
    );

    let p_a3_trans = engine.precedents_transitive("Sheet1", "A3").unwrap();
    assert_eq!(
        p_a3_trans,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }
            }
        ]
    );

    let d_a1_direct = engine.dependents("Sheet1", "A1").unwrap();
    assert_eq!(
        d_a1_direct,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 1, col: 0 }
        }]
    );

    let d_a1_trans = engine.dependents_transitive("Sheet1", "A1").unwrap();
    assert_eq!(
        d_a1_trans,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 2, col: 0 }
            }
        ]
    );
}

#[test]
fn precedents_for_filled_formulas_resolve_against_formula_cell() {
    let mut engine = Engine::new();

    // Simulate a fill-down pattern where each row references cells in the same row.
    engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A2+B2").unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "C2").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 1 }, // B2
            },
        ]
    );
}

#[test]
fn dirty_dependency_path_explains_why_cell_is_dirty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_formula("Sheet1", "A2", "=A1*2").unwrap();
    engine.set_cell_formula("Sheet1", "A3", "=A2+1").unwrap();
    engine.recalculate();

    engine.set_cell_value("Sheet1", "A1", 20.0).unwrap();

    assert!(engine.is_dirty("Sheet1", "A2"));
    assert!(engine.is_dirty("Sheet1", "A3"));

    let path = engine.dirty_dependency_path("Sheet1", "A3").unwrap();
    assert_eq!(
        path,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 2, col: 0 }
            }
        ]
    );

    // Recalc should clear dirty flags and the debug path.
    engine.recalculate();
    assert!(!engine.is_dirty("Sheet1", "A3"));
    assert!(engine.dirty_dependency_path("Sheet1", "A3").is_none());
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(41.0));
}

#[test]
fn field_access_dependencies_include_base_cell() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(123.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue::with_fields("Apple Inc.", fields)),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(123.0));

    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 0 }
        }]
    );
    assert_eq!(
        engine.dependents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }
        }]
    );
}

#[test]
fn full_column_precedent_is_compact_and_expansion_is_deterministic() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();

    let max_row = EXCEL_MAX_ROWS - 1;

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![PrecedentNode::Range {
            sheet: 0,
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr {
                row: max_row,
                col: 0
            }
        }]
    );

    // Expanded queries should be capped and ordered deterministically.
    let expanded_1 = engine.precedents_expanded("Sheet1", "B1", 3).unwrap();
    let expanded_2 = engine.precedents_expanded("Sheet1", "B1", 3).unwrap();
    assert_eq!(expanded_1, expanded_2);
    assert_eq!(
        expanded_1,
        vec![
            (0, CellAddr { row: 0, col: 0 }),
            (0, CellAddr { row: 1, col: 0 }),
            (0, CellAddr { row: 2, col: 0 })
        ]
    );

    // Reverse lookup via the range-node index should work without per-cell expansion.
    let dependents = engine.dependents("Sheet1", "A123").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }
        }]
    );
}

#[test]
fn expanded_precedents_merge_multiple_ranges_in_sheet_order() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SUM(A:A)+SUM(B:B)")
        .unwrap();

    let expanded = engine.precedents_expanded("Sheet1", "C1", 6).unwrap();
    assert_eq!(
        expanded,
        vec![
            (0, CellAddr { row: 0, col: 0 }), // A1
            (0, CellAddr { row: 0, col: 1 }), // B1
            (0, CellAddr { row: 1, col: 0 }), // A2
            (0, CellAddr { row: 1, col: 1 }), // B2
            (0, CellAddr { row: 2, col: 0 }), // A3
            (0, CellAddr { row: 2, col: 1 })  // B3
        ]
    );
}

#[test]
fn dirty_dependency_path_includes_range_node_for_range_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=B1+1").unwrap();
    engine.recalculate();

    engine.set_cell_value("Sheet1", "A2", 20.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    assert!(engine.is_dirty("Sheet1", "C1"));

    let max_row = EXCEL_MAX_ROWS - 1;
    let path = engine.dirty_dependency_path("Sheet1", "C1").unwrap();
    assert_eq!(
        path,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: 0
                }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 2 }
            }
        ]
    );
}

#[test]
fn sheet_range_3d_refs_participate_in_precedents_and_dependents_queries() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    assert_eq!(
        engine.precedents("Summary", "A1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 1,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 2,
                addr: CellAddr { row: 0, col: 0 }
            }
        ]
    );

    assert_eq!(
        engine.dependents("Sheet2", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 3,
            addr: CellAddr { row: 0, col: 0 }
        }]
    );
}

#[test]
fn sheet_range_precedents_follow_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");
    engine.ensure_sheet("Summary");
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    assert_eq!(
        engine.precedents("Summary", "A1").unwrap(),
        vec![
            PrecedentNode::Cell {
                sheet: 2,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 1,
                addr: CellAddr { row: 0, col: 0 }
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }
            }
        ]
    );
}

#[test]
fn expanded_sheet_range_precedents_follow_tab_order_after_reorder() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");
    engine.ensure_sheet("Summary");
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1:A2)")
        .unwrap();

    // Reverse the sheet tab order: Sheet3, Sheet2, Sheet1, Summary.
    assert!(engine.reorder_sheet("Sheet3", 0));
    assert!(engine.reorder_sheet("Sheet2", 1));

    assert_eq!(
        engine.precedents_expanded("Summary", "A1", 6).unwrap(),
        vec![
            (2, CellAddr { row: 0, col: 0 }), // Sheet3!A1
            (2, CellAddr { row: 1, col: 0 }), // Sheet3!A2
            (1, CellAddr { row: 0, col: 0 }), // Sheet2!A1
            (1, CellAddr { row: 1, col: 0 }), // Sheet2!A2
            (0, CellAddr { row: 0, col: 0 }), // Sheet1!A1
            (0, CellAddr { row: 1, col: 0 }), // Sheet1!A2
        ]
    );
}

#[test]
fn sheet_range_3d_refs_create_per_sheet_range_nodes_for_full_column_refs() {
    let mut engine = Engine::new();
    // Ensure the sheet order is deterministic for the 3D span resolution.
    engine.ensure_sheet("Summary");
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A:A)")
        .unwrap();

    let max_row = EXCEL_MAX_ROWS - 1;
    assert_eq!(
        engine.precedents("Summary", "A1").unwrap(),
        vec![
            PrecedentNode::Range {
                sheet: 1,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: 0
                }
            },
            PrecedentNode::Range {
                sheet: 2,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: 0
                }
            },
            PrecedentNode::Range {
                sheet: 3,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: 0
                }
            }
        ]
    );

    // Reverse lookup via the range-node index should work for intermediate sheets in a 3D span.
    assert_eq!(
        engine.dependents("Sheet2", "A123").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 0 }
        }]
    );
}

#[test]
fn sheet_range_3d_refs_create_per_sheet_range_nodes_for_full_sheet_refs() {
    let mut engine = Engine::new();
    // Ensure the sheet order is deterministic for the 3D span resolution.
    engine.ensure_sheet("Summary");
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");
    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A:XFD)")
        .unwrap();

    let max_row = EXCEL_MAX_ROWS - 1;
    let max_col = EXCEL_MAX_COLS - 1;
    assert_eq!(
        engine.precedents("Summary", "A1").unwrap(),
        vec![
            PrecedentNode::Range {
                sheet: 1,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: max_col
                }
            },
            PrecedentNode::Range {
                sheet: 2,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: max_col
                }
            },
            PrecedentNode::Range {
                sheet: 3,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr {
                    row: max_row,
                    col: max_col
                }
            },
        ]
    );

    // Reverse lookup via the range-node index should work for intermediate sheets in a 3D span.
    assert_eq!(
        engine.dependents("Sheet2", "C3").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 0 }
        }]
    );
}
