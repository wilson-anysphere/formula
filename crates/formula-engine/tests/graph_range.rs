use formula_engine::graph::{CellDeps, DependencyGraph, Precedent, SheetRange};
use formula_model::{CellId, CellRef, Range};

fn cell(sheet_id: u32, a1: &str) -> CellId {
    CellId {
        sheet_id,
        cell: CellRef::from_a1(a1).unwrap(),
    }
}

fn sheet_range(sheet_id: u32, a1: &str) -> SheetRange {
    let parts: Vec<&str> = a1.split(':').collect();
    match parts.as_slice() {
        [single] => {
            let c = CellRef::from_a1(single).unwrap();
            SheetRange::new(sheet_id, Range::new(c, c))
        }
        [start, end] => {
            let s = CellRef::from_a1(start).unwrap();
            let e = CellRef::from_a1(end).unwrap();
            SheetRange::new(sheet_id, Range::new(s, e))
        }
        _ => panic!("invalid range"),
    }
}

#[test]
fn range_dependency_is_represented_as_single_range_node() {
    let sheet = 1;
    let b1 = cell(sheet, "B1");
    let range = sheet_range(sheet, "A1:A1000");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Range(range)]));

    let stats = graph.stats();
    assert_eq!(stats.formula_cells, 1);
    assert_eq!(
        stats.direct_cell_edges, 0,
        "range refs must not explode to per-cell edges"
    );
    assert_eq!(stats.range_nodes, 1);
    assert_eq!(stats.range_edges, 1);

    // Dirtying any cell inside the range should mark the dependent cell dirty.
    let a500 = cell(sheet, "A500");
    graph.mark_dirty(a500);
    assert_eq!(graph.calc_order_for_dirty().unwrap(), vec![b1]);
}

#[test]
fn range_dependencies_participate_in_calculation_ordering() {
    // A2 depends on A1, and B1 depends on the range A1:A2. B1 must be ordered after A2.
    let sheet = 1;
    let a1 = cell(sheet, "A1");
    let a2 = cell(sheet, "A2");
    let b1 = cell(sheet, "B1");
    let range = sheet_range(sheet, "A1:A2");

    let mut graph = DependencyGraph::new();
    graph.update_cell_dependencies(a1, CellDeps::default());
    graph.update_cell_dependencies(a2, CellDeps::new(vec![Precedent::Cell(a1)]));
    graph.update_cell_dependencies(b1, CellDeps::new(vec![Precedent::Range(range)]));

    graph.mark_dirty(a1);
    let order = graph.calc_order_for_dirty().unwrap();
    assert_eq!(order, vec![a1, a2, b1]);
}
