use super::ast::Expr;
use super::grid::{Grid, GridMut};
use super::value::{CellCoord, Value};
use super::{BytecodeCache, Program, Vm};
use ahash::{AHashMap, AHashSet};
use crate::date::ExcelDateSystem;
use crate::locale::ValueLocaleConfig;
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use rayon::prelude::*;
use std::sync::Arc;

#[derive(Clone)]
pub struct FormulaCell {
    pub coord: CellCoord,
    pub expr: Expr,
}

pub struct CellNode {
    pub coord: CellCoord,
    pub program: Arc<Program>,
    pub deps: Vec<CellCoord>,
}

pub struct CalcGraph {
    pub nodes: Vec<CellNode>,
    pub levels: Vec<Vec<usize>>,
    pub max_level_width: usize,
}

impl CalcGraph {
    pub fn build(cells: Vec<FormulaCell>, cache: &BytecodeCache) -> Self {
        let mut nodes = Vec::with_capacity(cells.len());
        let mut index: AHashMap<(i32, i32), usize> = AHashMap::with_capacity(cells.len());

        for (i, cell) in cells.into_iter().enumerate() {
            let program = cache.get_or_compile(&cell.expr);
            let deps = collect_deps(&cell.expr, cell.coord);
            nodes.push(CellNode {
                coord: cell.coord,
                program,
                deps,
            });
            index.insert((cell.coord.row, cell.coord.col), i);
        }

        let mut dependents: Vec<Vec<usize>> = vec![Vec::new(); nodes.len()];
        let mut indegree: Vec<usize> = vec![0; nodes.len()];

        for (i, node) in nodes.iter().enumerate() {
            for dep in &node.deps {
                if let Some(&j) = index.get(&(dep.row, dep.col)) {
                    dependents[j].push(i);
                    indegree[i] += 1;
                }
            }
        }

        let mut levels: Vec<Vec<usize>> = Vec::new();
        let mut current: Vec<usize> = indegree
            .iter()
            .enumerate()
            .filter_map(|(i, &deg)| if deg == 0 { Some(i) } else { None })
            .collect();

        let mut remaining = nodes.len();
        while !current.is_empty() {
            remaining -= current.len();
            levels.push(current);
            let mut next: Vec<usize> = Vec::new();
            for &n in levels.last().unwrap() {
                for &m in &dependents[n] {
                    indegree[m] -= 1;
                    if indegree[m] == 0 {
                        next.push(m);
                    }
                }
            }
            current = next;
        }

        if remaining != 0 {
            // Cycle detected; fall back to a single sequential level.
            levels.clear();
            levels.push((0..nodes.len()).collect());
        }

        let max_level_width = levels.iter().map(|l| l.len()).max().unwrap_or(0);

        Self {
            nodes,
            levels,
            max_level_width,
        }
    }
}

pub struct RecalcEngine {
    cache: BytecodeCache,
}

impl Default for RecalcEngine {
    fn default() -> Self {
        Self::new()
    }
}

impl RecalcEngine {
    pub fn new() -> Self {
        Self {
            cache: BytecodeCache::new(),
        }
    }

    pub fn cache(&self) -> &BytecodeCache {
        &self.cache
    }

    pub fn build_graph(&self, cells: Vec<FormulaCell>) -> CalcGraph {
        CalcGraph::build(cells, &self.cache)
    }

    pub fn recalc(&self, graph: &CalcGraph, grid: &mut dyn GridMut) {
        let locale = crate::LocaleConfig::en_us();
        let mut results: Vec<Value> = Vec::with_capacity(graph.max_level_width);
        let now_utc = chrono::Utc::now();
        let date_system = ExcelDateSystem::EXCEL_1900;
        let value_locale = ValueLocaleConfig::en_us();

        for level in &graph.levels {
            results.clear();
            results.resize(level.len(), Value::Empty);
            {
                let g: &dyn Grid = &*grid;
                #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
                {
                    results
                        .par_iter_mut()
                        .zip(level.par_iter())
                        .for_each_init(
                            || {
                                (
                                    Vm::with_capacity(32),
                                    super::runtime::set_thread_eval_context(
                                        date_system,
                                        value_locale,
                                        now_utc.clone(),
                                    ),
                                )
                            },
                            |(vm, _guard), (out, &idx)| {
                                let node = &graph.nodes[idx];
                                *out = vm.eval(&node.program, g, node.coord, &locale);
                            },
                        );
                }
                #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
                {
                    let mut vm = Vm::with_capacity(32);
                    let _guard = super::runtime::set_thread_eval_context(
                        date_system,
                        value_locale,
                        now_utc.clone(),
                    );
                    for (out, &idx) in results.iter_mut().zip(level.iter()) {
                        let node = &graph.nodes[idx];
                        *out = vm.eval(&node.program, g, node.coord, &locale);
                    }
                }
            }

            for (i, &idx) in level.iter().enumerate() {
                let v = std::mem::replace(&mut results[i], Value::Empty);
                grid.set_value(graph.nodes[idx].coord, v);
            }
        }
    }
}

fn collect_deps(expr: &Expr, base: CellCoord) -> Vec<CellCoord> {
    let mut out: AHashSet<(i32, i32)> = AHashSet::new();
    collect_deps_inner(expr, base, &mut out);
    out.into_iter()
        .map(|(row, col)| CellCoord { row, col })
        .collect()
}

fn collect_deps_inner(expr: &Expr, base: CellCoord, out: &mut AHashSet<(i32, i32)>) {
    match expr {
        Expr::CellRef(r) => {
            let c = r.resolve(base);
            out.insert((c.row, c.col));
        }
        Expr::RangeRef(r) => {
            // Expand ranges into individual cells. This is intentionally simple and will be
            // replaced by range nodes in the full engine.
            let rr = r.resolve(base);
            for col in rr.col_start..=rr.col_end {
                for row in rr.row_start..=rr.row_end {
                    out.insert((row, col));
                }
            }
        }
        Expr::Literal(_) => {}
        Expr::Unary { expr, .. } => collect_deps_inner(expr, base, out),
        Expr::Binary { left, right, .. } => {
            collect_deps_inner(left, base, out);
            collect_deps_inner(right, base, out);
        }
        Expr::FuncCall { args, .. } => {
            for arg in args {
                collect_deps_inner(arg, base, out);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    fn assert_recalc_sets_eval_context_for_vms() {
        let origin = CellCoord::new(0, 0);
        let expr = crate::bytecode::parse_formula("=\"1.234,56\"+0", origin).expect("parse");

        let cache = BytecodeCache::new();
        let program = cache.get_or_compile(&expr);
        let empty_grid = crate::bytecode::ColumnarGrid::new(1, 1);
        let locale = crate::LocaleConfig::en_us();
        let now_utc = chrono::Utc.with_ymd_and_hms(2024, 1, 1, 0, 0, 0).unwrap();

        // Set a non-default locale on this thread so we can verify `RecalcEngine::recalc` uses its
        // own deterministic eval context rather than inheriting ambient thread-local state.
        let _outer_guard = crate::bytecode::runtime::set_thread_eval_context(
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::de_de(),
            now_utc.clone(),
        );

        let mut vm = Vm::with_capacity(32);
        let de_de_value = vm.eval(&program, &empty_grid, origin, &locale);
        let en_us_value = vm.eval_with_coercion_context(
            &program,
            &empty_grid,
            origin,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::en_us(),
            now_utc.clone(),
        );
        assert_ne!(
            de_de_value, en_us_value,
            "expected locale-dependent coercion to differ for the chosen input"
        );

        let engine = RecalcEngine::new();
        let graph = engine.build_graph(vec![FormulaCell {
            coord: origin,
            expr: expr.clone(),
        }]);
        let mut grid = crate::bytecode::ColumnarGrid::new(1, 1);

        engine.recalc(&graph, &mut grid);
        let recalc_value = grid.get_value(origin);
        assert_eq!(
            recalc_value, en_us_value,
            "recalc should evaluate using the engine's deterministic context"
        );

        // Ensure `recalc` restores the thread-local context after it finishes.
        let after_value = vm.eval(&program, &empty_grid, origin, &locale);
        assert_eq!(
            after_value, de_de_value,
            "recalc should restore any prior thread-local eval context"
        );
    }

    #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
    #[test]
    fn recalc_sets_eval_context_for_rayon_workers() {
        use std::sync::mpsc;
        use std::time::Duration;

        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(1)
            .build()
            .expect("build thread pool");
        let (tx, rx) = mpsc::channel::<std::thread::Result<()>>();
        pool.spawn(move || {
            let result = std::panic::catch_unwind(assert_recalc_sets_eval_context_for_vms);
            tx.send(result).ok();
        });

        match rx.recv_timeout(Duration::from_secs(5)).expect("recalc task") {
            Ok(()) => {}
            Err(panic) => std::panic::resume_unwind(panic),
        }
    }

    #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
    #[test]
    fn recalc_sets_eval_context_serial() {
        assert_recalc_sets_eval_context_for_vms();
    }
}
