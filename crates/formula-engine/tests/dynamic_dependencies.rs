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
