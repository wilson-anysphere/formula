use std::collections::{HashMap, HashSet, VecDeque};

use crate::engine::CellKey;

/// Compute strongly connected components (Tarjan) for the directed graph described by `edges`.
///
/// - `nodes` are the nodes to include.
/// - `edges` maps `u -> [v...]` and must only include edges between nodes in `nodes`.
pub(crate) fn strongly_connected_components(
    nodes: &[CellKey],
    edges: &HashMap<CellKey, Vec<CellKey>>,
) -> Vec<Vec<CellKey>> {
    let mut index: u32 = 0;
    let mut stack: Vec<CellKey> = Vec::new();
    let mut on_stack: HashSet<CellKey> = HashSet::new();
    let mut indices: HashMap<CellKey, u32> = HashMap::new();
    let mut lowlinks: HashMap<CellKey, u32> = HashMap::new();
    let mut sccs: Vec<Vec<CellKey>> = Vec::new();

    fn strong_connect(
        v: CellKey,
        index: &mut u32,
        stack: &mut Vec<CellKey>,
        on_stack: &mut HashSet<CellKey>,
        indices: &mut HashMap<CellKey, u32>,
        lowlinks: &mut HashMap<CellKey, u32>,
        edges: &HashMap<CellKey, Vec<CellKey>>,
        sccs: &mut Vec<Vec<CellKey>>,
    ) {
        indices.insert(v, *index);
        lowlinks.insert(v, *index);
        *index += 1;
        stack.push(v);
        on_stack.insert(v);

        for &w in edges.get(&v).map(Vec::as_slice).unwrap_or(&[]) {
            if !indices.contains_key(&w) {
                strong_connect(w, index, stack, on_stack, indices, lowlinks, edges, sccs);
                let low_v = lowlinks[&v];
                let low_w = lowlinks[&w];
                if low_w < low_v {
                    lowlinks.insert(v, low_w);
                }
            } else if on_stack.contains(&w) {
                let low_v = lowlinks[&v];
                let idx_w = indices[&w];
                if idx_w < low_v {
                    lowlinks.insert(v, idx_w);
                }
            }
        }

        if lowlinks[&v] == indices[&v] {
            let mut scc = Vec::new();
            loop {
                let w = stack.pop().expect("stack underflow");
                on_stack.remove(&w);
                scc.push(w);
                if w == v {
                    break;
                }
            }
            sccs.push(scc);
        }
    }

    for &v in nodes {
        if !indices.contains_key(&v) {
            strong_connect(
                v,
                &mut index,
                &mut stack,
                &mut on_stack,
                &mut indices,
                &mut lowlinks,
                edges,
                &mut sccs,
            );
        }
    }

    sccs
}

/// Topologically sort SCCs using the SCC condensation graph (Kahn), yielding evaluation order.
///
/// The order is stable across runs and uses the minimum `(sheet, row, col)` key of each SCC as the
/// tiebreaker when multiple SCCs are available.
pub(crate) fn topo_sort_sccs(
    sccs: &[Vec<CellKey>],
    edges: &HashMap<CellKey, Vec<CellKey>>,
) -> Vec<usize> {
    let mut cell_to_scc: HashMap<CellKey, usize> = HashMap::new();
    for (idx, scc) in sccs.iter().enumerate() {
        for &cell in scc {
            cell_to_scc.insert(cell, idx);
        }
    }

    let mut indegree = vec![0usize; sccs.len()];
    let mut scc_edges: Vec<HashSet<usize>> = vec![HashSet::new(); sccs.len()];

    for (&u, vs) in edges {
        let Some(&su) = cell_to_scc.get(&u) else {
            continue;
        };
        for &v in vs {
            let Some(&sv) = cell_to_scc.get(&v) else {
                continue;
            };
            if su != sv && scc_edges[su].insert(sv) {
                indegree[sv] += 1;
            }
        }
    }

    let mut queue: VecDeque<usize> = VecDeque::new();
    let mut zero: Vec<usize> = indegree
        .iter()
        .enumerate()
        .filter_map(|(idx, &deg)| (deg == 0).then_some(idx))
        .collect();
    zero.sort_by_key(|&idx| min_cell_key(&sccs[idx]));
    for idx in zero {
        queue.push_back(idx);
    }

    let mut out = Vec::with_capacity(sccs.len());
    while let Some(scc_idx) = queue.pop_front() {
        out.push(scc_idx);
        let mut next: Vec<usize> = scc_edges[scc_idx].iter().copied().collect();
        next.sort_by_key(|&idx| min_cell_key(&sccs[idx]));
        for v in next {
            indegree[v] = indegree[v].saturating_sub(1);
            if indegree[v] == 0 {
                queue.push_back(v);
            }
        }
    }

    out
}

fn min_cell_key(scc: &[CellKey]) -> (usize, u32, u32) {
    scc.iter()
        .map(|c| (c.sheet, c.addr.row, c.addr.col))
        .min()
        .unwrap_or((usize::MAX, u32::MAX, u32::MAX))
}
