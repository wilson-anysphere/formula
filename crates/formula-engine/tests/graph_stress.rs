use formula_engine::graph::{CellDeps, DependencyGraph, Precedent};
use formula_model::CellId;

#[test]
fn handles_100k_dependents_without_pathological_behavior() {
    let sheet = 1;
    let input = CellId::new(sheet, 0, 0);

    let mut graph = DependencyGraph::new();
    let dependent_count: u32 = 100_000;
    for row in 0..dependent_count {
        let cell = CellId::new(sheet, row, 1);
        graph.update_cell_dependencies(cell, CellDeps::new(vec![Precedent::Cell(input)]));
    }

    assert_eq!(graph.stats().direct_cell_edges as u32, dependent_count);

    graph.mark_dirty(input);
    let order = graph.calc_order_for_dirty().unwrap();
    assert_eq!(order.len() as u32, dependent_count);
    // Spot-check ordering is deterministic (increasing row).
    assert_eq!(order.first().copied(), Some(CellId::new(sheet, 0, 1)));
    assert_eq!(
        order.last().copied(),
        Some(CellId::new(sheet, dependent_count - 1, 1))
    );
}
