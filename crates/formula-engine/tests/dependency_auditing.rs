use formula_engine::eval::CellAddr;
use formula_engine::{Engine, Value};

#[test]
fn precedents_and_dependents_queries_work() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=A1*2")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=A2+1")
        .unwrap();
    engine.recalculate();

    let p_a3_direct = engine.precedents("Sheet1", "A3").unwrap();
    assert_eq!(p_a3_direct, vec![(0, CellAddr { row: 1, col: 0 })]);

    let p_a3_trans = engine.precedents_transitive("Sheet1", "A3").unwrap();
    assert_eq!(
        p_a3_trans,
        vec![(0, CellAddr { row: 0, col: 0 }), (0, CellAddr { row: 1, col: 0 })]
    );

    let d_a1_direct = engine.dependents("Sheet1", "A1").unwrap();
    assert_eq!(d_a1_direct, vec![(0, CellAddr { row: 1, col: 0 })]);

    let d_a1_trans = engine.dependents_transitive("Sheet1", "A1").unwrap();
    assert_eq!(
        d_a1_trans,
        vec![(0, CellAddr { row: 1, col: 0 }), (0, CellAddr { row: 2, col: 0 })]
    );
}

#[test]
fn dirty_dependency_path_explains_why_cell_is_dirty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=A1*2")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=A2+1")
        .unwrap();
    engine.recalculate();

    engine.set_cell_value("Sheet1", "A1", 20.0).unwrap();

    assert!(engine.is_dirty("Sheet1", "A2"));
    assert!(engine.is_dirty("Sheet1", "A3"));

    let path = engine.dirty_dependency_path("Sheet1", "A3").unwrap();
    assert_eq!(
        path,
        vec![
            (0, CellAddr { row: 0, col: 0 }),
            (0, CellAddr { row: 1, col: 0 }),
            (0, CellAddr { row: 2, col: 0 })
        ]
    );

    // Recalc should clear dirty flags and the debug path.
    engine.recalculate();
    assert!(!engine.is_dirty("Sheet1", "A3"));
    assert!(engine.dirty_dependency_path("Sheet1", "A3").is_none());
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(41.0));
}

