use formula_engine::graph::{CellDeps, DependencyGraph, GraphNode, Precedent, SheetRange};
use formula_model::{CellId, CellRef, Range};

fn cell(sheet_id: u32, a1: &str) -> CellId {
    CellId {
        sheet_id,
        cell: CellRef::from_a1(a1).unwrap(),
    }
}

#[test]
fn detects_direct_cycle_and_reports_path() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let b1 = cell(sheet, "B1");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(a1, CellDeps::new(vec![Precedent::Cell(b1)]));
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.mark_dirty(a1);

    let err = graph.calc_order_for_dirty().expect_err("cycle expected");
    assert_eq!(
        err.path,
        vec![
            GraphNode::Cell(a1),
            GraphNode::Cell(b1),
            GraphNode::Cell(a1)
        ]
    );
}

#[test]
fn detects_cycle_involving_range_node() {
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let self_ref = CellRef::from_a1("A1").unwrap();
    let self_range = SheetRange::new(sheet, Range::new(self_ref, self_ref));

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(a1, CellDeps::new(vec![Precedent::Range(self_range)]));
    graph.mark_dirty(a1);

    let err = graph.calc_order_for_dirty().expect_err("cycle expected");
    assert_eq!(err.path.len(), 3);
    assert_eq!(err.path.first(), Some(&GraphNode::Cell(a1)));
    assert_eq!(err.path.last(), Some(&GraphNode::Cell(a1)));
}
