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
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=VLOOKUP(2, OFFSET(A1,1,0,3,2), 2, FALSE)",
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
fn indirect_resolves_sheet_names_with_trailing_spaces() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1 ", "A1", 42.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=INDIRECT(\"'Sheet1 '!A1\")")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(42.0));
}
