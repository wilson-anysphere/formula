use formula_engine::graph::{CellDeps, DependencyGraph, Precedent};
use formula_model::{CellId, CellRef};
use std::collections::HashSet;

fn cell(sheet_id: u32, a1: &str) -> CellId {
    CellId {
        sheet_id,
        cell: CellRef::from_a1(a1).unwrap(),
    }
}

#[test]
fn volatile_cells_propagate_to_dependents() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let b1 = cell(sheet, "B1");
    let c1 = cell(sheet, "C1");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(a1, CellDeps::default().volatile(true));
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.update_cell_dependencies(c1, CellDeps::new(vec![Precedent::Cell(b1)]));

    let volatile = graph.volatile_cells();
    assert_eq!(volatile, HashSet::from([a1, b1, c1]));

    // Without explicitly marking anything dirty, a recalc should still include the volatile closure.
    let order = graph.calc_order_for_dirty().unwrap();
    assert_eq!(order, vec![a1, b1, c1]);
}
