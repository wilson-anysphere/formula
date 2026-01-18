use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet, VecDeque};
use std::fmt;

use crate::engine::CellKey;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum IterativeError {
    AllocationFailure(&'static str),
}

impl fmt::Display for IterativeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            IterativeError::AllocationFailure(ctx) => write!(f, "allocation failed ({ctx})"),
        }
    }
}

impl std::error::Error for IterativeError {}

/// Compute strongly connected components (Tarjan) for the directed graph described by `edges`.
///
/// - `nodes` are the nodes to include.
/// - `edges` maps `u -> [v...]` and must only include edges between nodes in `nodes`.
pub(crate) fn strongly_connected_components(
    nodes: &[CellKey],
    edges: &HashMap<CellKey, Vec<CellKey>>,
) -> Result<Vec<Vec<CellKey>>, IterativeError> {
    #[derive(Debug, Clone, Copy)]
    struct Frame {
        v_idx: usize,
        next_edge: usize,
        parent_idx: Option<usize>,
    }

    let mut cell_to_idx: HashMap<CellKey, usize> = HashMap::new();
    if cell_to_idx.try_reserve(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC cell_to_idx)");
        return Err(IterativeError::AllocationFailure("SCC cell_to_idx"));
    }

    let mut is_duplicate: Vec<bool> = Vec::new();
    if is_duplicate.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC duplicates)");
        return Err(IterativeError::AllocationFailure("SCC duplicates"));
    }
    is_duplicate.resize(nodes.len(), false);
    for (idx, &cell) in nodes.iter().enumerate() {
        match cell_to_idx.entry(cell) {
            Entry::Vacant(entry) => {
                entry.insert(idx);
            }
            Entry::Occupied(_) => {
                debug_assert!(false, "duplicate CellKey in SCC input: {cell:?}");
                is_duplicate[idx] = true;
            }
        }
    }

    let mut next_index: usize = 0;
    let mut stack: Vec<usize> = Vec::new();
    if stack.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC stack)");
        return Err(IterativeError::AllocationFailure("SCC stack"));
    }

    let mut on_stack: Vec<bool> = Vec::new();
    if on_stack.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC on_stack)");
        return Err(IterativeError::AllocationFailure("SCC on_stack"));
    }
    on_stack.resize(nodes.len(), false);

    let mut indices: Vec<Option<usize>> = Vec::new();
    if indices.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC indices)");
        return Err(IterativeError::AllocationFailure("SCC indices"));
    }
    indices.resize(nodes.len(), None);

    let mut lowlinks: Vec<usize> = Vec::new();
    if lowlinks.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC lowlinks)");
        return Err(IterativeError::AllocationFailure("SCC lowlinks"));
    }
    lowlinks.resize(nodes.len(), 0);

    let mut sccs: Vec<Vec<CellKey>> = Vec::new();
    let mut frames: Vec<Frame> = Vec::new();
    if frames.try_reserve_exact(nodes.len()).is_err() {
        debug_assert!(false, "allocation failed (SCC frames)");
        return Err(IterativeError::AllocationFailure("SCC frames"));
    }

    for start_idx in 0..nodes.len() {
        if is_duplicate[start_idx] {
            continue;
        }
        if indices[start_idx].is_some() {
            continue;
        }

        indices[start_idx] = Some(next_index);
        lowlinks[start_idx] = next_index;
        next_index += 1;
        stack.push(start_idx);
        on_stack[start_idx] = true;
        frames.push(Frame {
            v_idx: start_idx,
            next_edge: 0,
            parent_idx: None,
        });

        while let Some(frame) = frames.last_mut() {
            let v_idx = frame.v_idx;
            let v = nodes[v_idx];
            let neighbors = edges.get(&v).map(Vec::as_slice).unwrap_or(&[]);

            if frame.next_edge < neighbors.len() {
                let w = neighbors[frame.next_edge];
                frame.next_edge += 1;

                let Some(&w_idx) = cell_to_idx.get(&w) else {
                    // Be defensive even though `edges` is expected to only reference `nodes`.
                    continue;
                };

                if indices[w_idx].is_none() {
                    indices[w_idx] = Some(next_index);
                    lowlinks[w_idx] = next_index;
                    next_index += 1;
                    stack.push(w_idx);
                    on_stack[w_idx] = true;
                    frames.push(Frame {
                        v_idx: w_idx,
                        next_edge: 0,
                        parent_idx: Some(v_idx),
                    });
                    continue;
                }

                if on_stack[w_idx] {
                    let Some(idx_w) = indices[w_idx] else {
                        debug_assert!(false, "Tarjan indices missing for {w:?}");
                        continue;
                    };
                    if idx_w < lowlinks[v_idx] {
                        lowlinks[v_idx] = idx_w;
                    }
                }
                continue;
            }

            let parent_idx = frame.parent_idx;
            let Some(idx_v) = indices[v_idx] else {
                debug_assert!(false, "Tarjan indices missing for {v:?}");
                frames.pop();
                continue;
            };

            if lowlinks[v_idx] == idx_v {
                let mut count = 0usize;
                for &w_idx in stack.iter().rev() {
                    count += 1;
                    if w_idx == v_idx {
                        break;
                    }
                }

                let mut scc: Vec<CellKey> = Vec::new();
                if scc.try_reserve_exact(count).is_err() {
                    debug_assert!(false, "allocation failed (SCC component)");
                    return Err(IterativeError::AllocationFailure("SCC component"));
                }

                for _ in 0..count {
                    let Some(w_idx) = stack.pop() else {
                        debug_assert!(false, "Tarjan SCC stack underflow for {v:?}");
                        break;
                    };
                    on_stack[w_idx] = false;
                    scc.push(nodes[w_idx]);
                }
                if sccs.try_reserve(1).is_err() {
                    debug_assert!(false, "allocation failed (SCC output)");
                    return Err(IterativeError::AllocationFailure("SCC output"));
                }
                sccs.push(scc);
            }

            frames.pop();

            if let Some(parent_idx) = parent_idx {
                if lowlinks[v_idx] < lowlinks[parent_idx] {
                    lowlinks[parent_idx] = lowlinks[v_idx];
                }
            }
        }
    }

    Ok(sccs)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::eval::CellAddr;

    fn key(sheet: usize, row: u32, col: u32) -> CellKey {
        CellKey {
            sheet,
            addr: CellAddr { row, col },
        }
    }

    #[test]
    fn tarjan_finds_sccs_for_cycle_and_singleton() {
        // Graph:
        // A -> B -> C -> A (cycle)
        // D (isolated)
        let a = key(0, 0, 0);
        let b = key(0, 0, 1);
        let c = key(0, 0, 2);
        let d = key(0, 1, 0);

        let nodes = vec![a, b, c, d];
        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        edges.insert(a, vec![b]);
        edges.insert(b, vec![c]);
        edges.insert(c, vec![a]);
        edges.insert(d, vec![]);

        let mut sccs = strongly_connected_components(&nodes, &edges).expect("SCC should succeed");
        for scc in &mut sccs {
            scc.sort();
        }
        sccs.sort_by_key(|scc| scc.len());

        assert_eq!(sccs.len(), 2);
        assert_eq!(sccs[0], vec![d]);
        assert_eq!(sccs[1], vec![a, b, c]);
    }

    #[test]
    fn topo_sort_sccs_orders_components_by_dependencies() {
        // Condensation:
        // {A,B} -> {C}
        let a = key(0, 0, 0);
        let b = key(0, 0, 1);
        let c = key(0, 0, 2);

        let nodes = vec![a, b, c];
        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        edges.insert(a, vec![b, c]);
        edges.insert(b, vec![a, c]);
        edges.insert(c, vec![]);

        let mut sccs = strongly_connected_components(&nodes, &edges).expect("SCC should succeed");
        for scc in &mut sccs {
            scc.sort();
        }
        let order = topo_sort_sccs(&sccs, &edges).expect("topo sort should succeed");

        assert_eq!(order.len(), sccs.len());

        let mut ordered: Vec<Vec<CellKey>> = Vec::new();
        if ordered.try_reserve_exact(order.len()).is_err() {
            panic!("allocation failed (topo order mapping, len={})", order.len());
        }
        for idx in order {
            ordered.push(sccs[idx].clone());
        }
        let first = &ordered[0];
        let second = &ordered[1];

        assert!(first.contains(&a) && first.contains(&b));
        assert_eq!(second, &vec![c]);
    }

    #[test]
    fn tarjan_is_stack_safe_for_long_chains() {
        let n: u32 = 20_000;
        let mut nodes: Vec<CellKey> = Vec::new();
        if nodes.try_reserve_exact(n as usize).is_err() {
            panic!("allocation failed (tarjan test nodes, n={n})");
        }
        for row in 0..n {
            nodes.push(key(0, row, 0));
        }

        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        if edges.try_reserve(nodes.len()).is_err() {
            panic!("allocation failed (tarjan test edges, nodes={})", nodes.len());
        }
        for i in 0..(n - 1) {
            edges.insert(nodes[i as usize], vec![nodes[(i + 1) as usize]]);
        }
        edges.insert(nodes[(n - 1) as usize], vec![]);

        let sccs = strongly_connected_components(&nodes, &edges).expect("SCC should succeed");
        assert_eq!(sccs.len(), nodes.len());
        assert!(sccs.iter().all(|scc| scc.len() == 1));
    }
}

/// Topologically sort SCCs using the SCC condensation graph (Kahn), yielding evaluation order.
///
/// The order is stable across runs and uses the minimum `(sheet, row, col)` key of each SCC as the
/// tiebreaker when multiple SCCs are available.
pub(crate) fn topo_sort_sccs(
    sccs: &[Vec<CellKey>],
    edges: &HashMap<CellKey, Vec<CellKey>>,
) -> Result<Vec<usize>, IterativeError> {
    let mut min_key_by_scc: Vec<(usize, u32, u32)> = Vec::new();
    if min_key_by_scc.try_reserve_exact(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo min keys)");
        return Err(IterativeError::AllocationFailure("topo min keys"));
    }
    for scc in sccs {
        min_key_by_scc.push(min_cell_key(scc));
    }

    let mut cell_count = 0usize;
    for scc in sccs {
        cell_count = cell_count.saturating_add(scc.len());
    }

    let mut cell_to_scc: HashMap<CellKey, usize> = HashMap::new();
    if cell_to_scc.try_reserve(cell_count).is_err() {
        debug_assert!(false, "allocation failed (topo cell_to_scc)");
        return Err(IterativeError::AllocationFailure("topo cell_to_scc"));
    }
    for (idx, scc) in sccs.iter().enumerate() {
        for &cell in scc {
            cell_to_scc.insert(cell, idx);
        }
    }

    let mut indegree: Vec<usize> = Vec::new();
    if indegree.try_reserve_exact(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo indegree)");
        return Err(IterativeError::AllocationFailure("topo indegree"));
    }
    indegree.resize(sccs.len(), 0);

    let mut scc_edges: Vec<HashSet<usize>> = Vec::new();
    if scc_edges.try_reserve_exact(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo edges)");
        return Err(IterativeError::AllocationFailure("topo edges"));
    }
    for _ in 0..sccs.len() {
        scc_edges.push(HashSet::new());
    }

    for (&u, vs) in edges {
        let Some(&su) = cell_to_scc.get(&u) else {
            continue;
        };
        for &v in vs {
            let Some(&sv) = cell_to_scc.get(&v) else {
                continue;
            };
            if su == sv {
                continue;
            }

            if scc_edges[su].try_reserve(1).is_err() {
                debug_assert!(false, "allocation failed (topo edge set)");
                return Err(IterativeError::AllocationFailure("topo edge set"));
            }

            if scc_edges[su].insert(sv) {
                indegree[sv] = indegree[sv].saturating_add(1);
            }
        }
    }

    let mut queue: VecDeque<usize> = VecDeque::new();
    if queue.try_reserve(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo queue)");
        return Err(IterativeError::AllocationFailure("topo queue"));
    }

    let mut zero: Vec<usize> = Vec::new();
    if zero.try_reserve_exact(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo zero)");
        return Err(IterativeError::AllocationFailure("topo zero"));
    }
    for (idx, &deg) in indegree.iter().enumerate() {
        if deg == 0 {
            zero.push(idx);
        }
    }
    zero.sort_by_key(|&idx| min_key_by_scc[idx]);
    for idx in zero {
        queue.push_back(idx);
    }

    let mut out: Vec<usize> = Vec::new();
    if out.try_reserve_exact(sccs.len()).is_err() {
        debug_assert!(false, "allocation failed (topo output)");
        return Err(IterativeError::AllocationFailure("topo output"));
    }
    while let Some(scc_idx) = queue.pop_front() {
        out.push(scc_idx);
        let mut next: Vec<usize> = Vec::new();
        if next.try_reserve_exact(scc_edges[scc_idx].len()).is_err() {
            debug_assert!(false, "allocation failed (topo next)");
            return Err(IterativeError::AllocationFailure("topo next"));
        }
        for &v in &scc_edges[scc_idx] {
            next.push(v);
        }
        next.sort_by_key(|&idx| min_key_by_scc[idx]);
        for v in next {
            if indegree[v] == 0 {
                debug_assert!(false, "topo_sort_sccs indegree underflow for SCC {v}");
                continue;
            }
            indegree[v] -= 1;
            if indegree[v] == 0 {
                queue.push_back(v);
            }
        }
    }

    Ok(out)
}

fn min_cell_key(scc: &[CellKey]) -> (usize, u32, u32) {
    scc.iter()
        .map(|c| (c.sheet, c.addr.row, c.addr.col))
        .min()
        .unwrap_or((usize::MAX, u32::MAX, u32::MAX))
}
