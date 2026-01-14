use formula_engine::graph::{CellDeps, DependencyGraph, Precedent};
use formula_model::CellId;
use std::collections::HashSet;

#[test]
fn dirty_marking_over_limit_falls_back_to_full_recalc() {
    let sheet = 1;
    let input = CellId::new(sheet, 0, 0);

    let mut graph = DependencyGraph::with_dirty_mark_limit(10);

    // Large enough fan-out to exceed the limit.
    let reachable_count: u32 = 20;
    let isolated_count: u32 = 5;

    let mut all_formulas: Vec<CellId> = Vec::new();
    for row in 0..reachable_count {
        let cell = CellId::new(sheet, row, 1);
        all_formulas.push(cell);
        graph.update_cell_dependencies(cell, CellDeps::new(vec![Precedent::Cell(input)]));
    }

    // Additional formula cells that are not reachable from `input` via dependency edges.
    for row in 0..isolated_count {
        let cell = CellId::new(sheet, row, 2);
        all_formulas.push(cell);
        graph.update_cell_dependencies(cell, CellDeps::default());
    }

    graph.mark_dirty(input);

    let dirty: HashSet<CellId> = graph.dirty_cells().into_iter().collect();
    assert_eq!(dirty.len(), all_formulas.len());
    for cell in &all_formulas {
        assert!(
            dirty.contains(cell),
            "{cell:?} should be dirty after full recalc fallback"
        );
    }

    // Calculation order/levels should still be computable when everything is dirty.
    let order = graph.calc_order_for_dirty().expect("no cycle");
    assert_eq!(order.len(), all_formulas.len());
    let order_set: HashSet<CellId> = order.into_iter().collect();
    assert_eq!(order_set, dirty);

    let levels = graph.calc_levels_for_dirty().expect("no cycle");
    let flat: Vec<CellId> = levels.into_iter().flatten().collect();
    assert_eq!(flat.len(), all_formulas.len());
    let flat_set: HashSet<CellId> = flat.into_iter().collect();
    assert_eq!(flat_set, dirty);

    // Repeated calls should be stable (and not re-trigger expensive propagation).
    for _ in 0..3 {
        graph.mark_dirty(input);
        let dirty_now: HashSet<CellId> = graph.dirty_cells().into_iter().collect();
        assert_eq!(dirty_now, dirty);
    }
}

#[test]
fn dirty_marking_under_limit_remains_incremental() {
    let sheet = 1;
    let input = CellId::new(sheet, 0, 0);

    let mut graph = DependencyGraph::with_dirty_mark_limit(10);

    // Small enough fan-out to stay under the limit (note the limit includes the starting cell).
    let reachable_count: u32 = 5;
    let isolated_count: u32 = 3;

    let mut reachable: Vec<CellId> = Vec::new();
    let mut isolated: Vec<CellId> = Vec::new();

    for row in 0..reachable_count {
        let cell = CellId::new(sheet, row, 1);
        reachable.push(cell);
        graph.update_cell_dependencies(cell, CellDeps::new(vec![Precedent::Cell(input)]));
    }

    for row in 0..isolated_count {
        let cell = CellId::new(sheet, row, 2);
        isolated.push(cell);
        graph.update_cell_dependencies(cell, CellDeps::default());
    }

    graph.mark_dirty(input);

    let dirty: HashSet<CellId> = graph.dirty_cells().into_iter().collect();
    assert_eq!(dirty.len(), reachable.len());
    for cell in &reachable {
        assert!(dirty.contains(cell));
    }
    for cell in &isolated {
        assert!(!dirty.contains(cell));
    }

    let order = graph.calc_order_for_dirty().expect("no cycle");
    let order_set: HashSet<CellId> = order.into_iter().collect();
    assert_eq!(order_set, dirty);
}
