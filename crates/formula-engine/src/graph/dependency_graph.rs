use formula_model::{CellId, CellRef, Range, WorksheetId};
use rstar::{RTree, RTreeObject, AABB};
use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};
use std::fmt;

pub type RangeId = u32;

pub type SheetId = WorksheetId;

/// A rectangular range pinned to a worksheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SheetRange {
    pub sheet_id: SheetId,
    pub range: Range,
}

impl SheetRange {
    #[must_use]
    pub const fn new(sheet_id: SheetId, range: Range) -> Self {
        Self { sheet_id, range }
    }

    #[must_use]
    pub fn contains(&self, cell: CellId) -> bool {
        cell.sheet_id == self.sheet_id && self.range.contains(cell.cell)
    }

    #[must_use]
    pub fn envelope_i64(self) -> ([i64; 2], [i64; 2]) {
        let min = [self.range.start.row.into(), self.range.start.col.into()];
        let max = [self.range.end.row.into(), self.range.end.col.into()];
        (min, max)
    }

    #[must_use]
    pub fn start(self) -> CellRef {
        self.range.start
    }
}

/// A cell's precedent relationship.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Precedent {
    Cell(CellId),
    Range(SheetRange),
}

/// Metadata describing *how* a dependent cell references a precedent cell.
///
/// This is primarily intended for auditing UX (e.g. "B1 depends on A2 via `A:A`").
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DependentEdgeKind {
    /// The dependent directly referenced the precedent cell (e.g. `=A1`).
    DirectCell,
    /// The dependent referenced a range that contains the precedent cell (e.g. `=SUM(A:A)`).
    Range(SheetRange),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DependentEdge {
    pub dependent: CellId,
    pub kind: DependentEdgeKind,
}

/// Dependency metadata for a single cell.
#[derive(Debug, Clone, Default)]
pub struct CellDeps {
    pub precedents: Vec<Precedent>,
    pub is_volatile: bool,
}

impl CellDeps {
    #[must_use]
    pub fn new(precedents: Vec<Precedent>) -> Self {
        Self {
            precedents,
            is_volatile: false,
        }
    }

    #[must_use]
    pub fn volatile(mut self, is_volatile: bool) -> Self {
        self.is_volatile = is_volatile;
        self
    }
}

#[derive(Debug, Clone, Default)]
struct CellNode {
    precedent_cells: HashSet<CellId>,
    precedent_ranges: HashSet<RangeId>,
}

#[derive(Debug, Clone)]
struct RangeNode {
    range: SheetRange,
    dependents: HashSet<CellId>,
    /// Number of formula cells (i.e., `CellNode`s) currently inside `range`.
    member_formula_cells: usize,
}

#[derive(Debug, Clone)]
pub struct GraphStats {
    pub formula_cells: usize,
    /// Direct cell-to-cell dependency edges (precedent cell -> dependent formula cell).
    pub direct_cell_edges: usize,
    pub range_nodes: usize,
    /// Range node -> dependent formula cell edges.
    pub range_edges: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum GraphNode {
    Cell(CellId),
    Range(RangeId),
}

#[derive(Debug, Clone)]
pub struct CycleError {
    pub path: Vec<GraphNode>,
}

impl fmt::Display for CycleError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "circular reference: ")?;
        for (idx, node) in self.path.iter().enumerate() {
            if idx > 0 {
                write!(f, " -> ")?;
            }
            match node {
                GraphNode::Cell(cell) => write!(f, "S{}!{}", cell.sheet_id, cell.cell.to_a1())?,
                GraphNode::Range(id) => write!(f, "Range({id})")?,
            }
        }
        Ok(())
    }
}

impl std::error::Error for CycleError {}

#[derive(Debug, Clone, Copy)]
struct RangeIndexEntry {
    id: RangeId,
    envelope: AABB<[i64; 2]>,
}

impl PartialEq for RangeIndexEntry {
    fn eq(&self, other: &Self) -> bool {
        self.id == other.id
    }
}

impl Eq for RangeIndexEntry {}

impl RTreeObject for RangeIndexEntry {
    type Envelope = AABB<[i64; 2]>;

    fn envelope(&self) -> Self::Envelope {
        self.envelope
    }
}

#[derive(Debug, Clone, Copy)]
struct CellIndexEntry {
    cell: CellId,
    point: [i64; 2],
}

impl PartialEq for CellIndexEntry {
    fn eq(&self, other: &Self) -> bool {
        self.cell == other.cell
    }
}

impl Eq for CellIndexEntry {}

impl RTreeObject for CellIndexEntry {
    type Envelope = AABB<[i64; 2]>;

    fn envelope(&self) -> Self::Envelope {
        AABB::from_point(self.point)
    }
}

/// An incremental dependency graph with Excel-style range-node optimization.
#[derive(Debug)]
pub struct DependencyGraph {
    /// Formula cells and their precedents.
    cells: HashMap<CellId, CellNode>,

    /// Direct dependents, keyed by precedent cell address.
    cell_dependents: HashMap<CellId, HashSet<CellId>>,

    range_nodes: HashMap<RangeId, RangeNode>,
    range_ids: HashMap<SheetRange, RangeId>,
    next_range_id: RangeId,

    /// Per-sheet R-tree for range-node lookup: point -> containing ranges.
    range_index: HashMap<SheetId, RTree<RangeIndexEntry>>,
    /// Per-sheet R-tree for formula cell lookup: range -> member formula cells.
    cell_index: HashMap<SheetId, RTree<CellIndexEntry>>,

    dirty: HashSet<CellId>,

    volatile_roots: HashSet<CellId>,
    volatile_closure: HashSet<CellId>,
    volatile_closure_valid: bool,

    calc_chain: Vec<CellId>,
    calc_chain_valid: bool,

    /// Maximum number of nodes `mark_dirty` is allowed to visit before it falls back to
    /// a "full recalc" by marking all formula cells dirty (Excel-style behavior).
    dirty_mark_limit: usize,
}

impl DependencyGraph {
    /// Excel-like dependency propagation threshold before falling back to full-recalc mode.
    pub const DEFAULT_DIRTY_MARK_LIMIT: usize = 65_536;

    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a graph with a custom dirty propagation limit.
    ///
    /// This is primarily intended for tests.
    #[must_use]
    pub fn with_dirty_mark_limit(limit: usize) -> Self {
        Self {
            dirty_mark_limit: limit,
            ..Self::default()
        }
    }

    /// Set the dirty propagation limit.
    ///
    /// This is primarily intended for tests.
    pub fn set_dirty_mark_limit(&mut self, limit: usize) {
        self.dirty_mark_limit = limit;
    }

    /// Returns counts useful for asserting the internal representation in tests.
    #[must_use]
    pub fn stats(&self) -> GraphStats {
        GraphStats {
            formula_cells: self.cells.len(),
            direct_cell_edges: self.cell_dependents.values().map(HashSet::len).sum(),
            range_nodes: self.range_nodes.len(),
            range_edges: self.range_nodes.values().map(|n| n.dependents.len()).sum(),
        }
    }

    /// Update the precedents for a formula cell.
    ///
    /// This registers:
    /// - direct cell precedents (cell-to-cell edges)
    /// - range precedents (range nodes for dirty marking + ordering constraints)
    /// - volatile flag
    pub fn update_cell_dependencies(&mut self, cell: CellId, deps: CellDeps) {
        let is_new_formula_cell = !self.cells.contains_key(&cell);

        if is_new_formula_cell {
            self.cells.insert(cell, CellNode::default());
            self.insert_formula_cell_index(cell);
            self.bump_range_member_counts_for_new_formula_cell(cell);
        }

        // Remove existing precedent relationships for this cell.
        if let Some(node) = self.cells.get_mut(&cell) {
            let old_cells: Vec<CellId> = node.precedent_cells.drain().collect();
            for precedent in old_cells {
                if let Some(set) = self.cell_dependents.get_mut(&precedent) {
                    set.remove(&cell);
                    if set.is_empty() {
                        self.cell_dependents.remove(&precedent);
                    }
                }
            }

            let old_ranges: Vec<RangeId> = node.precedent_ranges.drain().collect();
            for range_id in old_ranges {
                self.detach_cell_from_range_node(range_id, cell);
            }
        }

        // Add new precedents.
        let mut new_precedent_cells: HashSet<CellId> = HashSet::new();
        let mut new_precedent_ranges: HashSet<RangeId> = HashSet::new();
        for precedent in deps.precedents {
            match precedent {
                Precedent::Cell(p) => {
                    new_precedent_cells.insert(p);
                    self.cell_dependents.entry(p).or_default().insert(cell);
                }
                Precedent::Range(range) => {
                    let range_id = self.intern_range_node(range);
                    new_precedent_ranges.insert(range_id);
                    self.range_nodes
                        .get_mut(&range_id)
                        .expect("range node must exist")
                        .dependents
                        .insert(cell);
                }
            }
        }
        if let Some(node) = self.cells.get_mut(&cell) {
            node.precedent_cells = new_precedent_cells;
            node.precedent_ranges = new_precedent_ranges;
        }

        // Update volatile root tracking.
        if deps.is_volatile {
            self.volatile_roots.insert(cell);
        } else {
            self.volatile_roots.remove(&cell);
        }

        self.calc_chain_valid = false;
        self.volatile_closure_valid = false;
    }

    /// Update the volatility flag for an existing formula cell without changing its precedents.
    ///
    /// If `cell` is not a tracked formula cell, this is a no-op.
    pub fn set_cell_volatile(&mut self, cell: CellId, is_volatile: bool) {
        if !self.cells.contains_key(&cell) {
            return;
        }

        let changed = if is_volatile {
            self.volatile_roots.insert(cell)
        } else {
            self.volatile_roots.remove(&cell)
        };
        if changed {
            self.volatile_closure_valid = false;
        }
    }

    /// Returns the direct precedents for a formula cell, including range nodes.
    ///
    /// Non-formula cells return an empty list.
    #[must_use]
    pub fn precedents_of(&self, cell: CellId) -> Vec<Precedent> {
        let Some(node) = self.cells.get(&cell) else {
            return Vec::new();
        };

        let mut out: Vec<Precedent> = Vec::with_capacity(
            node.precedent_cells
                .len()
                .saturating_add(node.precedent_ranges.len()),
        );
        out.extend(node.precedent_cells.iter().copied().map(Precedent::Cell));
        out.extend(
            node.precedent_ranges
                .iter()
                .filter_map(|id| self.range_nodes.get(id).map(|n| Precedent::Range(n.range))),
        );

        out.sort_by_key(|p| match p {
            Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
            Precedent::Range(r) => {
                let start = r.range.start;
                let end = r.range.end;
                (1u8, r.sheet_id, start.row, start.col, end.row, end.col)
            }
        });
        out
    }

    /// Returns all direct dependent formula cells for `cell`.
    ///
    /// This includes:
    /// - direct cell references (e.g. `=A1`)
    /// - range references that contain `cell` (e.g. `=SUM(A:A)` depends on `A1`)
    ///
    /// The returned list is deduplicated and sorted deterministically.
    #[must_use]
    pub fn dependents_of(&self, cell: CellId) -> Vec<DependentEdge> {
        let mut best: HashMap<CellId, DependentEdgeKind> = HashMap::new();

        if let Some(dependents) = self.cell_dependents.get(&cell) {
            for &dep in dependents {
                best.insert(dep, DependentEdgeKind::DirectCell);
            }
        }

        for range_id in self.range_nodes_containing_cell(cell) {
            let Some(range_node) = self.range_nodes.get(&range_id) else {
                continue;
            };
            let range_kind = DependentEdgeKind::Range(range_node.range);
            for &dep in &range_node.dependents {
                best.entry(dep)
                    .and_modify(|existing| {
                        if dependent_kind_sort_key(*existing) > dependent_kind_sort_key(range_kind)
                        {
                            *existing = range_kind;
                        }
                    })
                    .or_insert(range_kind);
            }
        }

        let mut out: Vec<DependentEdge> = best
            .into_iter()
            .map(|(dependent, kind)| DependentEdge { dependent, kind })
            .collect();
        out.sort_by_key(|edge| {
            (
                edge.dependent.sheet_id,
                edge.dependent.cell.row,
                edge.dependent.cell.col,
                dependent_kind_sort_key(edge.kind),
            )
        });
        out
    }

    /// Returns all formula cells in `range`, sorted deterministically.
    #[must_use]
    pub fn formula_cells_in_range(&self, range: SheetRange) -> Vec<CellId> {
        let Some(tree) = self.cell_index.get(&range.sheet_id) else {
            return Vec::new();
        };

        let (min, max) = range.envelope_i64();
        let env = AABB::from_corners(min, max);
        let mut out: Vec<CellId> = tree
            .locate_in_envelope_intersecting(&env)
            .map(|entry| entry.cell)
            .collect();
        out.sort_by_key(|cell| (cell.sheet_id, cell.cell.row, cell.cell.col));
        out
    }

    /// Removes a formula cell from the graph.
    ///
    /// Note: other cells may still depend on the removed cell address; those reverse edges remain so that
    /// future edits to that address continue to propagate dirtiness.
    pub fn remove_cell(&mut self, cell: CellId) {
        let Some(node) = self.cells.remove(&cell) else {
            // Not a tracked formula cell.
            self.dirty.remove(&cell);
            self.volatile_roots.remove(&cell);
            self.volatile_closure.remove(&cell);
            return;
        };

        // Remove from indices first so counts stay consistent.
        self.decrement_range_member_counts_for_removed_formula_cell(cell);
        self.remove_formula_cell_index(cell);

        // Detach from precedents.
        for precedent in node.precedent_cells {
            if let Some(set) = self.cell_dependents.get_mut(&precedent) {
                set.remove(&cell);
                if set.is_empty() {
                    self.cell_dependents.remove(&precedent);
                }
            }
        }
        for range_id in node.precedent_ranges {
            self.detach_cell_from_range_node(range_id, cell);
        }

        self.dirty.remove(&cell);
        self.volatile_roots.remove(&cell);
        self.volatile_closure.remove(&cell);
        self.calc_chain_valid = false;
        self.volatile_closure_valid = false;
    }

    /// Marks a cell as dirty and propagates to all transitive dependents.
    ///
    /// If `cell` is not a tracked formula cell it is treated as an input; the cell itself is not added to
    /// the dirty set, but its dependents still are.
    pub fn mark_dirty(&mut self, cell: CellId) {
        // If we already marked all formula cells dirty, further propagation is unnecessary.
        if self.dirty.len() == self.cells.len() {
            return;
        }

        let mut queue = VecDeque::new();
        let mut seen = HashSet::new();

        queue.push_back(cell);
        seen.insert(cell);

        if seen.len() > self.dirty_mark_limit {
            self.mark_all_formula_cells_dirty();
            return;
        }

        while let Some(cur) = queue.pop_front() {
            // If `cur` is a formula cell and was already dirty, we can stop exploring through it:
            // its transitive dependents must already have been marked dirty as well.
            if self.cells.contains_key(&cur) && !self.dirty.insert(cur) {
                continue;
            }

            let mut exceeded_limit = false;

            if let Some(dependents) = self.cell_dependents.get(&cur) {
                for &dep in dependents {
                    if seen.insert(dep) {
                        if seen.len() > self.dirty_mark_limit {
                            exceeded_limit = true;
                            break;
                        }
                        queue.push_back(dep);
                    }
                }
            }

            if exceeded_limit {
                self.mark_all_formula_cells_dirty();
                return;
            }

            for range_id in self.range_nodes_containing_cell(cur) {
                if let Some(range_node) = self.range_nodes.get(&range_id) {
                    for &dep in &range_node.dependents {
                        if seen.insert(dep) {
                            if seen.len() > self.dirty_mark_limit {
                                exceeded_limit = true;
                                break;
                            }
                            queue.push_back(dep);
                        }
                    }
                }
                if exceeded_limit {
                    break;
                }
            }

            if exceeded_limit {
                self.mark_all_formula_cells_dirty();
                return;
            }
        }
    }

    /// Returns the direct dependents of `cell`.
    ///
    /// This includes:
    /// - cell-to-cell dependents (formulas that reference the cell directly)
    /// - range-node dependents (formulas that reference a range that contains the cell)
    ///
    /// The returned list is de-duplicated and sorted deterministically by sheet/row/col.
    #[must_use]
    pub fn direct_dependents(&self, cell: CellId) -> Vec<CellId> {
        let mut vec: Vec<CellId> = Vec::new();

        if let Some(dependents) = self.cell_dependents.get(&cell) {
            vec.extend(dependents.iter().copied());
        }

        for range_id in self.range_nodes_containing_cell(cell) {
            if let Some(range_node) = self.range_nodes.get(&range_id) {
                vec.extend(range_node.dependents.iter().copied());
            }
        }

        vec.sort_by_key(|c| (c.sheet_id, c.cell.row, c.cell.col));
        vec.dedup();
        vec
    }

    #[must_use]
    pub fn dirty_cells(&self) -> Vec<CellId> {
        self.dirty.iter().copied().collect()
    }

    /// Returns the calculation order for the current dirty set (plus any volatile closure), in a
    /// topological order consistent with Excel semantics.
    pub fn calc_order_for_dirty(&mut self) -> Result<Vec<CellId>, CycleError> {
        self.rebuild_calc_chain()?;
        self.rebuild_volatile_closure_if_needed();

        let mut out = Vec::new();
        out.reserve(self.dirty.len().saturating_add(self.volatile_closure.len()));

        for &cell in &self.calc_chain {
            if self.dirty.contains(&cell) || self.volatile_closure.contains(&cell) {
                out.push(cell);
            }
        }

        Ok(out)
    }

    /// Returns the calculation schedule for the current dirty set (plus volatile closure),
    /// grouped into independent dependency levels.
    ///
    /// Each inner `Vec<CellId>` contains formula cells that can be evaluated in parallel because
    /// all of their precedents (within the evaluated subset) appear in earlier levels.
    ///
    /// This is similar to [`calc_order_for_dirty`], but preserves the level structure needed for
    /// multi-threaded recalculation. Range nodes are treated as ordering constraints and are
    /// processed internally (they do not appear in the returned schedule).
    pub fn calc_levels_for_dirty(&mut self) -> Result<Vec<Vec<CellId>>, CycleError> {
        // Ensure the full graph has no cycles. If there is a cycle anywhere in the workbook,
        // Excel's calculation chain becomes undefined; we surface the cycle as an error and
        // avoid producing a partial schedule.
        self.rebuild_calc_chain()?;
        self.rebuild_volatile_closure_if_needed();

        let mut eval_cells: HashSet<CellId> = HashSet::new();
        eval_cells.extend(self.dirty.iter().copied());
        eval_cells.extend(self.volatile_closure.iter().copied());

        if eval_cells.is_empty() {
            return Ok(Vec::new());
        }

        // Relevant range nodes are those referenced as precedents by any cell we will evaluate.
        let mut relevant_ranges: HashSet<RangeId> = HashSet::new();
        for cell in eval_cells.iter().copied() {
            if let Some(node) = self.cells.get(&cell) {
                relevant_ranges.extend(node.precedent_ranges.iter().copied());
            }
        }

        // In-degree for formula cells restricted to the evaluated subset.
        let mut cell_in: HashMap<CellId, usize> = HashMap::with_capacity(eval_cells.len());
        let mut ready_cells: BTreeSet<ReadyNode> = BTreeSet::new();
        for cell in eval_cells.iter().copied() {
            let Some(node) = self.cells.get(&cell) else {
                continue;
            };
            let direct_formula_precedents = node
                .precedent_cells
                .iter()
                .filter(|p| eval_cells.contains(p))
                .count();
            let deg = direct_formula_precedents + node.precedent_ranges.len();
            cell_in.insert(cell, deg);
            if deg == 0 {
                ready_cells.insert(ReadyNode::cell(cell));
            }
        }

        // In-degree for range nodes = number of evaluated formula cells inside the range.
        let mut range_in: HashMap<RangeId, usize> = relevant_ranges
            .iter()
            .copied()
            .map(|id| (id, 0usize))
            .collect();
        for cell in eval_cells.iter().copied() {
            for range_id in self.range_nodes_containing_cell(cell) {
                if let Some(deg) = range_in.get_mut(&range_id) {
                    *deg = deg.saturating_add(1);
                }
            }
        }

        let mut ready_ranges: BTreeSet<ReadyNode> = BTreeSet::new();
        for (&range_id, &deg) in &range_in {
            if deg == 0 {
                ready_ranges.insert(self.ready_range_node(range_id));
            }
        }

        let mut out: Vec<Vec<CellId>> = Vec::new();
        let mut processed = 0usize;

        while processed < eval_cells.len() {
            // Range nodes are ordering-only: process any that are ready before selecting the next
            // batch of evaluable formula cells.
            while let Some(node) = ready_ranges.pop_first() {
                let range_id = match node.kind {
                    ReadyNodeKind::Range(id) => id,
                    ReadyNodeKind::Cell(_) => unreachable!("ready_ranges only stores range nodes"),
                };

                range_in.remove(&range_id);

                if let Some(range_node) = self.range_nodes.get(&range_id) {
                    for &dep in &range_node.dependents {
                        if !eval_cells.contains(&dep) {
                            continue;
                        }
                        if let Some(deg) = cell_in.get_mut(&dep) {
                            *deg = deg.saturating_sub(1);
                            if *deg == 0 {
                                ready_cells.insert(ReadyNode::cell(dep));
                            }
                        }
                    }
                }
            }

            if ready_cells.is_empty() {
                // The full graph is acyclic (checked above), so this should not happen. Still, be
                // defensive in case of internal inconsistency.
                break;
            }

            let mut level: Vec<CellId> = Vec::with_capacity(ready_cells.len());
            while let Some(node) = ready_cells.pop_first() {
                let cell = match node.kind {
                    ReadyNodeKind::Cell(cell) => cell,
                    ReadyNodeKind::Range(_) => unreachable!("ready_cells only stores cell nodes"),
                };
                // `cell_in` only tracks cells we still need to schedule; skip if already processed.
                if cell_in.remove(&cell).is_none() {
                    continue;
                }
                level.push(cell);
            }

            if level.is_empty() {
                break;
            }

            processed += level.len();
            out.push(level.clone());

            let mut next_ready_cells: BTreeSet<ReadyNode> = BTreeSet::new();
            let mut next_ready_ranges: BTreeSet<ReadyNode> = BTreeSet::new();

            for cell in level {
                // Direct dependents.
                if let Some(dependents) = self.cell_dependents.get(&cell) {
                    for &dep in dependents {
                        if let Some(deg) = cell_in.get_mut(&dep) {
                            *deg = deg.saturating_sub(1);
                            if *deg == 0 {
                                next_ready_cells.insert(ReadyNode::cell(dep));
                            }
                        }
                    }
                }

                // Range membership edges.
                for range_id in self.range_nodes_containing_cell(cell) {
                    if let Some(deg) = range_in.get_mut(&range_id) {
                        *deg = deg.saturating_sub(1);
                        if *deg == 0 {
                            next_ready_ranges.insert(self.ready_range_node(range_id));
                        }
                    }
                }
            }

            ready_cells = next_ready_cells;
            ready_ranges = next_ready_ranges;
        }

        if processed != eval_cells.len() {
            return Err(CycleError { path: Vec::new() });
        }

        Ok(out)
    }

    /// Clears the explicit dirty set. Volatile cells remain effectively dirty on every call to
    /// [`calc_order_for_dirty`].
    pub fn clear_dirty(&mut self) {
        self.dirty.clear();
    }

    #[must_use]
    pub fn volatile_cells(&mut self) -> HashSet<CellId> {
        self.rebuild_volatile_closure_if_needed();
        self.volatile_closure.clone()
    }

    fn rebuild_volatile_closure_if_needed(&mut self) {
        if self.volatile_closure_valid {
            return;
        }

        let mut visited = HashSet::new();
        let mut queue = VecDeque::new();

        for &root in &self.volatile_roots {
            visited.insert(root);
            queue.push_back(root);
        }

        while let Some(cur) = queue.pop_front() {
            if let Some(dependents) = self.cell_dependents.get(&cur) {
                for &dep in dependents {
                    if visited.insert(dep) {
                        queue.push_back(dep);
                    }
                }
            }

            for range_id in self.range_nodes_containing_cell(cur) {
                if let Some(range_node) = self.range_nodes.get(&range_id) {
                    for &dep in &range_node.dependents {
                        if visited.insert(dep) {
                            queue.push_back(dep);
                        }
                    }
                }
            }
        }

        self.volatile_closure = visited;
        self.volatile_closure_valid = true;
    }

    fn rebuild_calc_chain(&mut self) -> Result<(), CycleError> {
        if self.calc_chain_valid {
            return Ok(());
        }

        // In-degree for formula cell nodes.
        let mut cell_in: HashMap<CellId, usize> = HashMap::with_capacity(self.cells.len());
        for (&cell, node) in &self.cells {
            let direct_formula_precedents = node
                .precedent_cells
                .iter()
                .filter(|p| self.cells.contains_key(p))
                .count();
            let deg = direct_formula_precedents + node.precedent_ranges.len();
            cell_in.insert(cell, deg);
        }

        // In-degree for range nodes = number of formula cells currently in the range.
        let mut range_in: HashMap<RangeId, usize> = HashMap::with_capacity(self.range_nodes.len());
        for (&id, node) in &self.range_nodes {
            range_in.insert(id, node.member_formula_cells);
        }

        let total_nodes = cell_in.len() + range_in.len();

        let mut ready = BTreeSet::new();
        for (&cell, &deg) in &cell_in {
            if deg == 0 {
                ready.insert(ReadyNode::cell(cell));
            }
        }
        for (&range_id, &deg) in &range_in {
            if deg == 0 {
                ready.insert(self.ready_range_node(range_id));
            }
        }

        let mut chain = Vec::with_capacity(self.cells.len());
        let mut processed_nodes = 0usize;

        while let Some(node) = ready.pop_first() {
            processed_nodes += 1;
            match node.kind {
                ReadyNodeKind::Cell(cell) => {
                    chain.push(cell);

                    // Direct cell dependents.
                    if let Some(dependents) = self.cell_dependents.get(&cell) {
                        for &dep in dependents {
                            if let Some(deg) = cell_in.get_mut(&dep) {
                                *deg = deg.saturating_sub(1);
                                if *deg == 0 {
                                    ready.insert(ReadyNode::cell(dep));
                                }
                            }
                        }
                    }

                    // Range membership edges: this cell contributes to any containing range node's in-degree.
                    for range_id in self.range_nodes_containing_cell(cell) {
                        if let Some(deg) = range_in.get_mut(&range_id) {
                            *deg = deg.saturating_sub(1);
                            if *deg == 0 {
                                ready.insert(self.ready_range_node(range_id));
                            }
                        }
                    }
                }
                ReadyNodeKind::Range(range_id) => {
                    if let Some(range_node) = self.range_nodes.get(&range_id) {
                        for &dep in &range_node.dependents {
                            if let Some(deg) = cell_in.get_mut(&dep) {
                                *deg = deg.saturating_sub(1);
                                if *deg == 0 {
                                    ready.insert(ReadyNode::cell(dep));
                                }
                            }
                        }
                    }
                }
            }
        }

        if processed_nodes != total_nodes {
            let remaining_cells: HashSet<CellId> = cell_in
                .into_iter()
                .filter_map(|(c, d)| (d > 0).then_some(c))
                .collect();
            let remaining_ranges: HashSet<RangeId> = range_in
                .into_iter()
                .filter_map(|(r, d)| (d > 0).then_some(r))
                .collect();
            let cycle_path = self
                .find_cycle(&remaining_cells, &remaining_ranges)
                .unwrap_or_else(|| vec![]);
            return Err(CycleError { path: cycle_path });
        }

        self.calc_chain = chain;
        self.calc_chain_valid = true;
        Ok(())
    }

    fn ready_range_node(&self, range_id: RangeId) -> ReadyNode {
        let range = &self
            .range_nodes
            .get(&range_id)
            .expect("range must exist")
            .range;
        let start = range.start();
        ReadyNode::range(range_id, range.sheet_id, start.row, start.col)
    }

    fn find_cycle(
        &self,
        remaining_cells: &HashSet<CellId>,
        remaining_ranges: &HashSet<RangeId>,
    ) -> Option<Vec<GraphNode>> {
        #[derive(Clone, Copy, PartialEq, Eq)]
        enum Color {
            White,
            Gray,
            Black,
        }

        let mut color: HashMap<GraphNode, Color> = HashMap::new();
        for &cell in remaining_cells {
            color.insert(GraphNode::Cell(cell), Color::White);
        }
        for &range_id in remaining_ranges {
            color.insert(GraphNode::Range(range_id), Color::White);
        }

        let mut stack: Vec<GraphNode> = Vec::new();
        let mut pos_in_stack: HashMap<GraphNode, usize> = HashMap::new();

        let mut nodes: Vec<GraphNode> = color.keys().copied().collect();
        nodes.sort_by(|a, b| {
            self.graph_node_sort_key(*a)
                .cmp(&self.graph_node_sort_key(*b))
        });

        for start in nodes {
            if color.get(&start) != Some(&Color::White) {
                continue;
            }

            let mut frames: Vec<Frame> = Vec::new();
            frames.push(Frame {
                node: start,
                neighbors: self.remaining_neighbors(start, remaining_cells, remaining_ranges),
                idx: 0,
            });
            stack.push(start);
            pos_in_stack.insert(start, stack.len() - 1);
            color.insert(start, Color::Gray);

            while let Some(frame) = frames.last_mut() {
                if frame.idx >= frame.neighbors.len() {
                    color.insert(frame.node, Color::Black);
                    pos_in_stack.remove(&frame.node);
                    stack.pop();
                    frames.pop();
                    continue;
                }

                let next = frame.neighbors[frame.idx];
                frame.idx += 1;

                match color.get(&next).copied().unwrap_or(Color::Black) {
                    Color::White => {
                        color.insert(next, Color::Gray);
                        stack.push(next);
                        pos_in_stack.insert(next, stack.len() - 1);
                        frames.push(Frame {
                            node: next,
                            neighbors: self.remaining_neighbors(
                                next,
                                remaining_cells,
                                remaining_ranges,
                            ),
                            idx: 0,
                        });
                    }
                    Color::Gray => {
                        let start_idx = *pos_in_stack.get(&next).unwrap_or(&0);
                        let mut cycle: Vec<GraphNode> = stack[start_idx..].to_vec();
                        cycle.push(next);
                        return Some(cycle);
                    }
                    Color::Black => {}
                }
            }
        }

        None
    }

    fn remaining_neighbors(
        &self,
        node: GraphNode,
        remaining_cells: &HashSet<CellId>,
        remaining_ranges: &HashSet<RangeId>,
    ) -> Vec<GraphNode> {
        let mut out = Vec::new();
        match node {
            GraphNode::Cell(cell) => {
                // Direct dependents.
                if let Some(dependents) = self.cell_dependents.get(&cell) {
                    for &dep in dependents {
                        if remaining_cells.contains(&dep) {
                            out.push(GraphNode::Cell(dep));
                        }
                    }
                }

                // Membership edges.
                for range_id in self.range_nodes_containing_cell(cell) {
                    if remaining_ranges.contains(&range_id) {
                        out.push(GraphNode::Range(range_id));
                    }
                }
            }
            GraphNode::Range(range_id) => {
                if let Some(range_node) = self.range_nodes.get(&range_id) {
                    for &dep in &range_node.dependents {
                        if remaining_cells.contains(&dep) {
                            out.push(GraphNode::Cell(dep));
                        }
                    }
                }
            }
        }

        out.sort_by(|a, b| {
            self.graph_node_sort_key(*a)
                .cmp(&self.graph_node_sort_key(*b))
        });
        out
    }

    fn graph_node_sort_key(&self, node: GraphNode) -> (u32, u32, u32, u32, u32) {
        match node {
            GraphNode::Cell(cell) => (0, cell.sheet_id, cell.cell.row, cell.cell.col, 0),
            GraphNode::Range(range_id) => {
                let range = &self.range_nodes.get(&range_id).expect("range exists").range;
                let start = range.start();
                (1, range.sheet_id, start.row, start.col, range_id)
            }
        }
    }

    fn intern_range_node(&mut self, range: SheetRange) -> RangeId {
        if let Some(&id) = self.range_ids.get(&range) {
            return id;
        }

        let id = self.next_range_id;
        self.next_range_id = self.next_range_id.wrapping_add(1);

        let member_formula_cells = self.count_formula_cells_in_range(range);

        self.range_nodes.insert(
            id,
            RangeNode {
                range,
                dependents: HashSet::new(),
                member_formula_cells,
            },
        );
        self.range_ids.insert(range, id);
        self.insert_range_index(range, id);

        id
    }

    fn detach_cell_from_range_node(&mut self, range_id: RangeId, cell: CellId) {
        let Some(node) = self.range_nodes.get_mut(&range_id) else {
            return;
        };
        node.dependents.remove(&cell);
        if node.dependents.is_empty() {
            let range = node.range;
            self.range_nodes.remove(&range_id);
            self.range_ids.remove(&range);
            self.remove_range_index(range, range_id);
            self.calc_chain_valid = false;
            self.volatile_closure_valid = false;
        }
    }

    fn insert_range_index(&mut self, range: SheetRange, id: RangeId) {
        let (min, max) = range.envelope_i64();
        let entry = RangeIndexEntry {
            id,
            envelope: AABB::from_corners(min, max),
        };
        self.range_index
            .entry(range.sheet_id)
            .or_default()
            .insert(entry);
    }

    fn remove_range_index(&mut self, range: SheetRange, id: RangeId) {
        if let Some(tree) = self.range_index.get_mut(&range.sheet_id) {
            let (min, max) = range.envelope_i64();
            let entry = RangeIndexEntry {
                id,
                envelope: AABB::from_corners(min, max),
            };
            tree.remove(&entry);
            if tree.size() == 0 {
                self.range_index.remove(&range.sheet_id);
            }
        }
    }

    fn insert_formula_cell_index(&mut self, cell: CellId) {
        let entry = CellIndexEntry {
            cell,
            point: cell_to_point(cell),
        };
        self.cell_index
            .entry(cell.sheet_id)
            .or_default()
            .insert(entry);
    }

    fn remove_formula_cell_index(&mut self, cell: CellId) {
        if let Some(tree) = self.cell_index.get_mut(&cell.sheet_id) {
            let entry = CellIndexEntry {
                cell,
                point: cell_to_point(cell),
            };
            tree.remove(&entry);
            if tree.size() == 0 {
                self.cell_index.remove(&cell.sheet_id);
            }
        }
    }

    fn range_nodes_containing_cell(&self, cell: CellId) -> Vec<RangeId> {
        let Some(tree) = self.range_index.get(&cell.sheet_id) else {
            return Vec::new();
        };
        let point = cell_to_point(cell);
        let env = AABB::from_point(point);
        tree.locate_in_envelope_intersecting(&env)
            .map(|entry| entry.id)
            .collect()
    }

    fn count_formula_cells_in_range(&self, range: SheetRange) -> usize {
        let Some(tree) = self.cell_index.get(&range.sheet_id) else {
            return 0;
        };
        let (min, max) = range.envelope_i64();
        let env = AABB::from_corners(min, max);
        tree.locate_in_envelope_intersecting(&env).count()
    }

    fn bump_range_member_counts_for_new_formula_cell(&mut self, cell: CellId) {
        for range_id in self.range_nodes_containing_cell(cell) {
            if let Some(range_node) = self.range_nodes.get_mut(&range_id) {
                range_node.member_formula_cells = range_node.member_formula_cells.saturating_add(1);
            }
        }
    }

    fn decrement_range_member_counts_for_removed_formula_cell(&mut self, cell: CellId) {
        for range_id in self.range_nodes_containing_cell(cell) {
            if let Some(range_node) = self.range_nodes.get_mut(&range_id) {
                range_node.member_formula_cells = range_node.member_formula_cells.saturating_sub(1);
            }
        }
    }

    fn mark_all_formula_cells_dirty(&mut self) {
        if self.dirty.len() == self.cells.len() {
            return;
        }

        self.dirty
            .reserve(self.cells.len().saturating_sub(self.dirty.len()));
        self.dirty.extend(self.cells.keys().copied());
    }
}

impl Default for DependencyGraph {
    fn default() -> Self {
        Self {
            cells: HashMap::new(),
            cell_dependents: HashMap::new(),
            range_nodes: HashMap::new(),
            range_ids: HashMap::new(),
            next_range_id: 0,
            range_index: HashMap::new(),
            cell_index: HashMap::new(),
            dirty: HashSet::new(),
            volatile_roots: HashSet::new(),
            volatile_closure: HashSet::new(),
            volatile_closure_valid: false,
            calc_chain: Vec::new(),
            calc_chain_valid: false,
            dirty_mark_limit: Self::DEFAULT_DIRTY_MARK_LIMIT,
        }
    }
}

fn dependent_kind_sort_key(kind: DependentEdgeKind) -> (u8, u32, u32, u32, u32, u32) {
    match kind {
        DependentEdgeKind::DirectCell => (0, 0, 0, 0, 0, 0),
        DependentEdgeKind::Range(range) => {
            let start = range.range.start;
            let end = range.range.end;
            (1, range.sheet_id, start.row, start.col, end.row, end.col)
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReadyNodeKind {
    Cell(CellId),
    Range(RangeId),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReadyNode {
    sheet: SheetId,
    row: u32,
    col: u32,
    kind: ReadyNodeKind,
}

impl ReadyNode {
    fn cell(cell: CellId) -> Self {
        Self {
            sheet: cell.sheet_id,
            row: cell.cell.row,
            col: cell.cell.col,
            kind: ReadyNodeKind::Cell(cell),
        }
    }

    fn range(range_id: RangeId, sheet: SheetId, row: u32, col: u32) -> Self {
        Self {
            sheet,
            row,
            col,
            kind: ReadyNodeKind::Range(range_id),
        }
    }
}

impl Ord for ReadyNode {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        (self.sheet, self.row, self.col, self.kind_sort_key()).cmp(&(
            other.sheet,
            other.row,
            other.col,
            other.kind_sort_key(),
        ))
    }
}

impl PartialOrd for ReadyNode {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl ReadyNode {
    fn kind_sort_key(&self) -> (u8, u32) {
        match self.kind {
            ReadyNodeKind::Cell(_) => (0, 0),
            ReadyNodeKind::Range(id) => (1, id),
        }
    }
}

#[derive(Debug)]
struct Frame {
    node: GraphNode,
    neighbors: Vec<GraphNode>,
    idx: usize,
}

fn cell_to_point(cell: CellId) -> [i64; 2] {
    [cell.cell.row.into(), cell.cell.col.into()]
}
