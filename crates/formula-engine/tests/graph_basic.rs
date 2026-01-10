use formula_engine::graph::{CellDeps, DependencyGraph, Precedent};
use formula_model::{CellId, CellRef};

fn cell(sheet_id: u32, a1: &str) -> CellId {
    CellId {
        sheet_id,
        cell: CellRef::from_a1(a1).unwrap(),
    }
}

#[test]
fn simple_dependency_tracking_and_dirty_marking() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let b1 = cell(sheet, "B1");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Cell(a1)]));

    let stats = graph.stats();
    assert_eq!(stats.formula_cells, 1);
    assert_eq!(stats.direct_cell_edges, 1);
    assert_eq!(stats.range_nodes, 0);

    graph.mark_dirty(a1);
    let mut dirty = graph.dirty_cells();
    dirty.sort_by_key(|c| (c.sheet_id, c.cell.row, c.cell.col));
    assert_eq!(dirty, vec![b1]);

    let order = graph.calc_order_for_dirty().expect("no cycle");
    assert_eq!(order, vec![b1]);
}

#[test]
fn dirty_marking_transitive_closure() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let b1 = cell(sheet, "B1");
    let c1 = cell(sheet, "C1");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.update_cell_dependencies(c1, CellDeps::new(vec![Precedent::Cell(b1)]));

    graph.mark_dirty(a1);
    let mut dirty = graph.dirty_cells();
    dirty.sort_by_key(|c| (c.sheet_id, c.cell.row, c.cell.col));
    assert_eq!(dirty, vec![b1, c1]);

    let order = graph.calc_order_for_dirty().expect("no cycle");
    assert_eq!(order, vec![b1, c1]);
}

#[test]
fn correct_calc_ordering_multiple_branches() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let b1 = cell(sheet, "B1");
    let c1 = cell(sheet, "C1");
    let d1 = cell(sheet, "D1");

    let mut graph = DependencyGraph::new();
    // A1 is a formula cell with no precedents.
    graph.update_cell_dependencies(a1, CellDeps::default());
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.update_cell_dependencies(c1, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.update_cell_dependencies(
        d1,
        CellDeps::new(vec![Precedent::Cell(b1), Precedent::Cell(c1)]),
    );

    graph.mark_dirty(a1);
    let order = graph.calc_order_for_dirty().expect("no cycle");
    assert_eq!(order, vec![a1, b1, c1, d1]);
}
