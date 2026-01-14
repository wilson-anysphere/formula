use formula_engine::eval::CellAddr;
use formula_engine::{Engine, PrecedentNode, Value};

#[test]
fn offset_updates_range_precedents_and_dependents() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(OFFSET(A1,1,0,3,1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 0 },   // A4
            },
        ]
    );

    // The dynamically-referenced cells should be visible as dependents once the formula has
    // evaluated and updated the dependency graph.
    let dependents = engine.dependents("Sheet1", "A3").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn tocol_updates_range_precedents_and_dependents_for_offset_input() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=TOCOL(OFFSET(A1,1,0,3,1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(3.0));

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 0 },   // A4
            },
        ]
    );

    let dependents = engine.dependents("Sheet1", "A3").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn xmatch_updates_range_precedents_and_dependents_for_offset_lookup_array() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=XMATCH(2, OFFSET(A1,1,0,3,1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 0 },   // A4
            },
        ]
    );

    let dependents = engine.dependents("Sheet1", "A3").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn xlookup_updates_range_precedents_and_dependents_for_offset_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 30.0).unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=XLOOKUP(2, OFFSET(A1,1,0,3,1), OFFSET(B1,1,0,3,1))",
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(20.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 }, // B1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 0 },   // A4
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 1 }, // B2
                end: CellAddr { row: 3, col: 1 },   // B4
            },
        ]
    );

    let dependents_a3 = engine.dependents("Sheet1", "A3").unwrap();
    let dependents_b3 = engine.dependents("Sheet1", "B3").unwrap();
    let expected = vec![PrecedentNode::Cell {
        sheet: 0,
        addr: CellAddr { row: 0, col: 2 }, // C1
    }];
    assert_eq!(dependents_a3, expected);
    assert_eq!(dependents_b3, expected);
}

#[test]
fn match_updates_range_precedents_and_dependents_for_offset_lookup_array() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=MATCH(2, OFFSET(A1,1,0,3,1), 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 0 },   // A4
            },
        ]
    );

    let dependents = engine.dependents("Sheet1", "A3").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn vlookup_updates_range_precedents_and_dependents_for_offset_table_array() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 30.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=VLOOKUP(2, OFFSET(A1,1,0,3,2), 2, FALSE)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(20.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 1 },   // B4
            },
        ]
    );

    let dependents_a3 = engine.dependents("Sheet1", "A3").unwrap();
    let dependents_b3 = engine.dependents("Sheet1", "B3").unwrap();
    let expected = vec![PrecedentNode::Cell {
        sheet: 0,
        addr: CellAddr { row: 0, col: 2 }, // C1
    }];
    assert_eq!(dependents_a3, expected);
    assert_eq!(dependents_b3, expected);
}

#[test]
fn hlookup_updates_range_precedents_and_dependents_for_offset_table_array() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", "A").unwrap();
    engine.set_cell_value("Sheet1", "B2", "B").unwrap();
    engine.set_cell_value("Sheet1", "C2", "C").unwrap();
    engine.set_cell_value("Sheet1", "A3", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 30.0).unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D1",
            "=HLOOKUP(\"B\", OFFSET(A1,1,0,2,3), 2, FALSE)",
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(20.0));

    let precedents = engine.precedents("Sheet1", "D1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 2, col: 2 },   // C3
            },
        ]
    );

    let dependents_b2 = engine.dependents("Sheet1", "B2").unwrap();
    let dependents_b3 = engine.dependents("Sheet1", "B3").unwrap();
    let expected = vec![PrecedentNode::Cell {
        sheet: 0,
        addr: CellAddr { row: 0, col: 3 }, // D1
    }];
    assert_eq!(dependents_b2, expected);
    assert_eq!(dependents_b3, expected);
}

#[test]
fn index_updates_range_precedents_and_dependents_for_offset_array() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 6.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=INDEX(OFFSET(A1,1,0,3,2), 2, 2)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(4.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 1, col: 0 }, // A2
                end: CellAddr { row: 3, col: 1 },   // B4
            },
        ]
    );

    let dependents = engine.dependents("Sheet1", "B3").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn indirect_establishes_dependencies_for_calc_order() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1").unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=INDIRECT(\"B1\")")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );
}

#[test]
fn indirect_dependency_switching_updates_graph_edges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=INDIRECT(D1)")
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(10.0));

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 } // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 } // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(20.0));

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 } // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 } // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_address_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"address\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("$A$1".to_string())
    );

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("$A$2".to_string())
    );

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A2").unwrap().is_empty());
}

#[test]
fn cell_col_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"col\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target to a different column.
    engine.set_cell_value("Sheet1", "D1", "B1").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));

    let precedents_b1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_b1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "B1").unwrap().is_empty());
}

#[test]
fn cell_row_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"row\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target to a different row.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A2").unwrap().is_empty());
}

#[test]
fn cell_width_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"width\",INDIRECT(D1))")
        .unwrap();

    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A2").unwrap().is_empty());
}

#[test]
fn row_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=ROW(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A2").unwrap().is_empty());
}

#[test]
fn column_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=COLUMN(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "B1").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "B1").unwrap().is_empty());
}

#[test]
fn rows_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1:A3").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=ROWS(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A1:A2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A2").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "A3").unwrap().is_empty());
}

#[test]
fn columns_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1:C1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=COLUMNS(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A1:B1").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "B1").unwrap().is_empty());
    assert!(engine.dependents("Sheet1", "C1").unwrap().is_empty());
}

#[test]
fn sheet_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    // Ensure Sheet1 is first in tab order by creating it before Sheet2.
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SHEET(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target sheet.
    engine.set_cell_value("Sheet1", "D1", "Sheet2!A1").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet2", "A1").unwrap().is_empty());
}

#[test]
fn isref_does_not_record_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=ISREF(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    // Switch the target to an invalid reference.
    engine.set_cell_value("Sheet1", "D1", "NotARef").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(false));
    assert_eq!(
        engine.precedents("Sheet1", "C1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 3 }, // D1
        }]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
}

#[test]
fn cell_protect_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"protect\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_prefix_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"prefix\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_format_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"format\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_color_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"color\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_parentheses_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"parentheses\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_contents_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"contents\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_type_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"type\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_filename_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=CELL(\"filename\",INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_filename_without_reference_does_not_create_self_edge_in_dynamic_dep_tracing() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=LEN(CELL("filename"))+INDIRECT("B1")"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }, // B1
        }]
    );
    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
}

#[test]
fn formulatext_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    engine.set_cell_formula("Sheet1", "A2", "=2+2").unwrap();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=FORMULATEXT(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("=1+1".to_string())
    );

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("=2+2".to_string())
    );

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn isformula_records_dynamic_dependencies_for_indirect_reference_arg() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    engine.set_cell_value("Sheet1", "A2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", "A1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=ISFORMULA(INDIRECT(D1))")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));

    let precedents_a1 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a1,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 0 }, // A1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
        ]
    );

    // Switch the target.
    engine.set_cell_value("Sheet1", "D1", "A2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(false));

    let precedents_a2 = engine.precedents("Sheet1", "C1").unwrap();
    assert_eq!(
        precedents_a2,
        vec![
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 3 }, // D1
            },
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 1, col: 0 }, // A2
            },
        ]
    );

    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
    assert_eq!(
        engine.dependents("Sheet1", "A2").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn cell_format_without_reference_does_not_create_self_edge_in_dynamic_dep_tracing() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=LEN(CELL("format"))+INDIRECT("B1")"#)
        .unwrap();
    engine.recalculate();

    // `CELL("format")` defaults to using the current cell as its implicit reference, but that
    // implicit self-reference should not be recorded as a calculation dependency when dynamic
    // dependency tracing is active.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(
        engine.precedents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }, // B1
        }]
    );
    assert!(engine.dependents("Sheet1", "A1").unwrap().is_empty());
}

#[test]
fn indirect_updates_precedents_and_dependents_when_reference_is_dereferenced_by_array_lift() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=ABS(INDIRECT("B1"))"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 },
            }
        )),
        "expected C1 precedents to include Sheet1!B1, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "B1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 2 },
            }
        )),
        "expected B1 dependents to include Sheet1!C1, got: {dependents:?}"
    );
}

#[test]
fn indirect_updates_precedents_and_dependents_for_concat_reference_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.set_cell_value("Sheet1", "A2", "y").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=CONCAT(INDIRECT("A1:A2"))"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("xy".to_string())
    );

    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 1, col: 0 },
            }
        )),
        "expected B1 precedents to include Sheet1!A1:A2, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "A1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 },
            }
        )),
        "expected A1 dependents to include Sheet1!B1, got: {dependents:?}"
    );
}

#[test]
fn indirect_updates_precedents_and_dependents_for_sumproduct_reference_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            r#"=SUMPRODUCT(INDIRECT("A1:A2"),INDIRECT("B1:B2"))"#,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(50.0));

    let precedents = engine.precedents("Sheet1", "C1").unwrap();
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 0 },
                end: CellAddr { row: 1, col: 0 },
            }
        )),
        "expected C1 precedents to include Sheet1!A1:A2, got: {precedents:?}"
    );
    assert!(
        precedents.iter().any(|node| matches!(
            node,
            PrecedentNode::Range {
                sheet: 0,
                start: CellAddr { row: 0, col: 1 },
                end: CellAddr { row: 1, col: 1 },
            }
        )),
        "expected C1 precedents to include Sheet1!B1:B2, got: {precedents:?}"
    );

    let dependents = engine.dependents("Sheet1", "A1").unwrap();
    assert!(
        dependents.iter().any(|node| matches!(
            node,
            PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 2 },
            }
        )),
        "expected A1 dependents to include Sheet1!C1, got: {dependents:?}"
    );
}

#[test]
fn indirect_resolves_sheet_names_with_trailing_spaces() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1 ", "A1", 42.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=INDIRECT(\"'Sheet1 '!A1\")")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(42.0));
}

#[test]
fn indirect_supports_degenerate_3d_sheet_spans() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();
    engine
        .set_cell_formula("Summary", "A1", "=INDIRECT(\"Sheet1:Sheet1!A1\")")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(42.0));
    assert_eq!(
        engine.precedents("Summary", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 0 }
        }]
    );
}

#[test]
fn indirect_rejects_non_degenerate_3d_sheet_spans() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=INDIRECT(\"Sheet1:Sheet3!A1\")")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Summary", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
    assert!(engine.precedents("Summary", "A1").unwrap().is_empty());
    assert!(engine.dependents("Sheet2", "A1").unwrap().is_empty());
}
