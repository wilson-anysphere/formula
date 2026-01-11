use formula_engine::eval::CellAddr;
use formula_engine::{Engine, PrecedentNode, Value};

#[test]
fn large_ranges_do_not_expand_into_per_cell_audit_edges() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A20000", 5.0).unwrap();

    // The key regression: `set_cell_formula` must not expand A1:A20000 into 20k distinct
    // audit-graph precedents. It should be represented as a single range node.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A20000)")
        .unwrap();

    // Direct precedents should stay compact: a single range node rather than 20k cell nodes.
    let precedents = engine.precedents("Sheet1", "B1").unwrap();
    assert_eq!(
        precedents,
        vec![PrecedentNode::Range {
            sheet: 0,
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 19_999, col: 0 },
        }]
    );

    // Precedent tracing expands ranges on-demand, but must cap expansion to avoid returning
    // millions of cells for very large ranges.
    let expanded_1 = engine.precedents_expanded("Sheet1", "B1", 100).unwrap();
    let expanded_2 = engine.precedents_expanded("Sheet1", "B1", 100).unwrap();
    assert_eq!(expanded_1, expanded_2);
    assert_eq!(expanded_1.len(), 100);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

    // Dirty propagation must work through range nodes: changing any cell inside the range should
    // dirty the dependent formula.
    engine.set_cell_value("Sheet1", "A12345", 10.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));

    // Dependents lookup must also work without range expansion.
    let deps = engine.dependents("Sheet1", "A12345").unwrap();
    assert_eq!(
        deps,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }
        }]
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(16.0));
}
