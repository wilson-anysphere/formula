use super::ast::{BinaryOp, Expr, Function, UnaryOp};
use super::grid::Grid;
use super::value::{
    Array as ArrayValue, CellCoord, ErrorKind, MultiRangeRef, RangeRef, Ref, ResolvedRange,
    SheetId, SheetRangeRef, Value,
};
use crate::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::error::ExcelError;
use crate::eval::split_external_sheet_key;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::lookup;
use crate::functions::math::criteria::Criteria as EngineCriteria;
use crate::functions::wildcard::WildcardPattern;
use crate::locale::ValueLocaleConfig;
use crate::simd::{self, CmpOp, NumericCriteria};
use crate::value::{
    cmp_case_insensitive, format_number_general_with_options, parse_number,
    ErrorKind as EngineErrorKind, RecordValue, Value as EngineValue,
};
use chrono::{DateTime, Datelike, Timelike, Utc};
use smallvec::SmallVec;
use std::borrow::Cow;
use std::cell::{Cell, RefCell};
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

thread_local! {
    static BYTECODE_DATE_SYSTEM: Cell<ExcelDateSystem> = Cell::new(ExcelDateSystem::EXCEL_1900);
    static BYTECODE_VALUE_LOCALE: Cell<ValueLocaleConfig> = Cell::new(ValueLocaleConfig::en_us());
    static BYTECODE_NOW_UTC: RefCell<DateTime<Utc>> = RefCell::new(Utc::now());
    static BYTECODE_RECALC_ID: Cell<u64> = Cell::new(0);
    static BYTECODE_CURRENT_SHEET_ID: Cell<u64> = Cell::new(0);
    static BYTECODE_RNG_COUNTER: Cell<u64> = Cell::new(0);
}

#[derive(Clone, Debug)]
struct ResolvedSheetRange {
    sheet: SheetId,
    range: ResolvedRange,
    /// Index of the originating area in the source [`MultiRangeRef`].
    ///
    /// This is used to preserve AST-like error precedence for functions that iterate ranges
    /// row-major: reference unions are evaluated area-by-area (in sorted order), and within an
    /// area the first error is the one with the smallest `(row, col)` coordinate.
    area_idx: usize,
}

fn intersect_resolved_ranges(a: ResolvedRange, b: ResolvedRange) -> Option<ResolvedRange> {
    let row_start = a.row_start.max(b.row_start);
    let row_end = a.row_end.min(b.row_end);
    if row_start > row_end {
        return None;
    }

    let col_start = a.col_start.max(b.col_start);
    let col_end = a.col_end.min(b.col_end);
    if col_start > col_end {
        return None;
    }

    Some(ResolvedRange {
        row_start,
        row_end,
        col_start,
        col_end,
    })
}

fn subtract_resolved_range(a: ResolvedRange, b: ResolvedRange) -> Vec<ResolvedRange> {
    let Some(i) = intersect_resolved_ranges(a, b) else {
        return vec![a];
    };

    // Full coverage: subtraction yields an empty set.
    if i.row_start == a.row_start
        && i.row_end == a.row_end
        && i.col_start == a.col_start
        && i.col_end == a.col_end
    {
        return Vec::new();
    }

    let mut out = Vec::new();

    // Emit disjoint pieces in row-major order (top -> middle left/right -> bottom).
    if a.row_start < i.row_start {
        out.push(ResolvedRange {
            row_start: a.row_start,
            row_end: i.row_start - 1,
            col_start: a.col_start,
            col_end: a.col_end,
        });
    }

    if a.col_start < i.col_start {
        out.push(ResolvedRange {
            row_start: i.row_start,
            row_end: i.row_end,
            col_start: a.col_start,
            col_end: i.col_start - 1,
        });
    }

    if i.col_end < a.col_end {
        out.push(ResolvedRange {
            row_start: i.row_start,
            row_end: i.row_end,
            col_start: i.col_end + 1,
            col_end: a.col_end,
        });
    }

    if i.row_end < a.row_end {
        out.push(ResolvedRange {
            row_start: i.row_end + 1,
            row_end: a.row_end,
            col_start: a.col_start,
            col_end: a.col_end,
        });
    }

    out
}

fn cmp_sheet_ids_in_tab_order(grid: &dyn Grid, a: &SheetId, b: &SheetId) -> Ordering {
    match (a, b) {
        (SheetId::Local(a_id), SheetId::Local(b_id)) => {
            let a_idx = grid.sheet_order_index(*a_id).unwrap_or(*a_id);
            let b_idx = grid.sheet_order_index(*b_id).unwrap_or(*b_id);
            a_idx.cmp(&b_idx).then_with(|| a_id.cmp(b_id))
        }
        (SheetId::Local(_), SheetId::External(_)) => Ordering::Less,
        (SheetId::External(_), SheetId::Local(_)) => Ordering::Greater,
        (SheetId::External(a_key), SheetId::External(b_key)) => {
            // Preserve external-workbook tab order when available.
            match (
                split_external_sheet_key(a_key),
                split_external_sheet_key(b_key),
            ) {
                (Some((a_wb, a_sheet)), Some((b_wb, b_sheet))) if a_wb == b_wb => {
                    match grid.external_sheet_order(a_wb) {
                        Some(order) => {
                            let mut a_idx: Option<usize> = None;
                            let mut b_idx: Option<usize> = None;
                            for (idx, name) in order.iter().enumerate() {
                                if a_idx.is_none()
                                    && formula_model::sheet_name_eq_case_insensitive(name, a_sheet)
                                {
                                    a_idx = Some(idx);
                                }
                                if b_idx.is_none()
                                    && formula_model::sheet_name_eq_case_insensitive(name, b_sheet)
                                {
                                    b_idx = Some(idx);
                                }
                                if a_idx.is_some() && b_idx.is_some() {
                                    break;
                                }
                            }
                            match (a_idx, b_idx) {
                                (Some(a_idx), Some(b_idx)) => {
                                    a_idx.cmp(&b_idx).then_with(|| a_key.cmp(b_key))
                                }
                                _ => a_key.cmp(b_key),
                            }
                        }
                        None => a_key.cmp(b_key),
                    }
                }
                _ => a_key.cmp(b_key),
            }
        }
    }
}

/// Convert a [`MultiRangeRef`] into a sequence of disjoint rectangular areas.
///
/// Excel reference unions behave like set union: overlapping cells should only be visited once.
/// The bytecode runtime represents unions as multi-range values, so we normalize them here by
/// subtracting overlaps in a deterministic order that matches the AST evaluator.
fn multirange_unique_areas(
    r: &MultiRangeRef,
    grid: &dyn Grid,
    base: CellCoord,
) -> Vec<ResolvedSheetRange> {
    let mut out = Vec::new();
    let mut seen_by_sheet: HashMap<SheetId, Vec<ResolvedRange>> = HashMap::new();

    // Match `Evaluator::eval_arg`: ensure a stable ordering for deterministic behavior (and to
    // align error precedence with the AST backend).
    let mut areas: Vec<(SheetRangeRef, ResolvedRange)> = r
        .areas
        .iter()
        .cloned()
        .map(|area| {
            let resolved = area.range.resolve(base);
            (area, resolved)
        })
        .collect();
    areas.sort_by(|(a_area, a_range), (b_area, b_range)| {
        cmp_sheet_ids_in_tab_order(grid, &a_area.sheet, &b_area.sheet)
            .then_with(|| a_range.row_start.cmp(&b_range.row_start))
            .then_with(|| a_range.col_start.cmp(&b_range.col_start))
            .then_with(|| a_range.row_end.cmp(&b_range.row_end))
            .then_with(|| a_range.col_end.cmp(&b_range.col_end))
    });

    for (area_idx, (area, resolved)) in areas.into_iter().enumerate() {
        let sheet = area.sheet.clone();

        let seen = seen_by_sheet.entry(sheet.clone()).or_default();

        let mut pieces = vec![resolved];
        for prev in seen.iter().copied() {
            let mut next = Vec::new();
            for piece in pieces {
                next.extend(subtract_resolved_range(piece, prev));
            }
            pieces = next;
            if pieces.is_empty() {
                break;
            }
        }

        seen.extend(pieces.iter().copied());
        out.extend(pieces.into_iter().map(|range| ResolvedSheetRange {
            sheet: sheet.clone(),
            range,
            area_idx,
        }));
    }

    out
}

/// Row-span threshold for treating a reference as "huge" and preferring sparse iteration.
///
/// The engine's bytecode column-cache builder also uses this threshold to avoid allocating large
/// `Vec<f64>` buffers for ranges like `A:A` on sparse sheets. When the cache is skipped, the
/// bytecode runtime can still compute aggregates correctly by iterating only stored cells.
pub(crate) const BYTECODE_SPARSE_RANGE_ROW_THRESHOLD: i32 = 262_144;

/// Maximum number of cells the bytecode runtime is willing to materialize into an in-memory array.
///
/// This is a safety guard to avoid allocating enormous intermediate buffers for expressions that
/// would otherwise require dense materialization (e.g. `A:A+1` on very large sheets). When the
/// limit is exceeded, the runtime surfaces `#SPILL!` instead of attempting the allocation.
pub(crate) const BYTECODE_MAX_RANGE_CELLS: usize = MAX_MATERIALIZED_ARRAY_CELLS;

#[inline]
fn range_should_iterate_sparse(range: ResolvedRange) -> bool {
    let rows = range.rows();
    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD {
        return true;
    }
    let cols = range.cols();
    if rows <= 0 || cols <= 0 {
        return false;
    }
    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);
    cells > BYTECODE_MAX_RANGE_CELLS as i64
}

// Array aggregates (e.g. SUM({1,2,3,...})) are executed inside the bytecode runtime even when no
// cell/range references are involved. These can be hot for large array literals or array-producing
// functions that are bytecode-eligible.
//
// Use the same stack-buffer chunking strategy as the AST evaluator to avoid allocating full-length
// `Vec<f64>` buffers.
const SIMD_AGGREGATE_BLOCK: usize = 1024;
const SIMD_ARRAY_MIN_LEN: usize = 32;

pub(crate) struct BytecodeEvalContextGuard {
    prev_date_system: ExcelDateSystem,
    prev_value_locale: ValueLocaleConfig,
    prev_now_utc: DateTime<Utc>,
    prev_recalc_id: u64,
}

impl Drop for BytecodeEvalContextGuard {
    fn drop(&mut self) {
        BYTECODE_DATE_SYSTEM.with(|cell| cell.set(self.prev_date_system));
        BYTECODE_VALUE_LOCALE.with(|cell| cell.set(self.prev_value_locale));
        BYTECODE_NOW_UTC.with(|cell| {
            cell.replace(self.prev_now_utc.clone());
        });
        BYTECODE_RECALC_ID.with(|cell| cell.set(self.prev_recalc_id));
    }
}

pub(crate) fn set_thread_eval_context(
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    recalc_id: u64,
) -> BytecodeEvalContextGuard {
    let prev_date_system = BYTECODE_DATE_SYSTEM.with(|cell| cell.replace(date_system));
    let prev_value_locale = BYTECODE_VALUE_LOCALE.with(|cell| cell.replace(value_locale));
    let prev_now_utc = BYTECODE_NOW_UTC.with(|cell| cell.replace(now_utc));
    let prev_recalc_id = BYTECODE_RECALC_ID.with(|cell| cell.replace(recalc_id));

    BytecodeEvalContextGuard {
        prev_date_system,
        prev_value_locale,
        prev_now_utc,
        prev_recalc_id,
    }
}

fn thread_date_system() -> ExcelDateSystem {
    BYTECODE_DATE_SYSTEM.with(|cell| cell.get())
}

fn thread_value_locale() -> ValueLocaleConfig {
    BYTECODE_VALUE_LOCALE.with(|cell| cell.get())
}

fn thread_number_locale() -> crate::value::NumberLocale {
    let separators = thread_value_locale().separators;
    crate::value::NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep))
}

fn thread_now_utc() -> DateTime<Utc> {
    BYTECODE_NOW_UTC.with(|cell| cell.borrow().clone())
}

fn thread_recalc_id() -> u64 {
    BYTECODE_RECALC_ID.with(|cell| cell.get())
}

fn thread_current_sheet_id() -> u64 {
    BYTECODE_CURRENT_SHEET_ID.with(|cell| cell.get())
}

pub(crate) fn set_thread_current_sheet_id(sheet_id: usize) {
    BYTECODE_CURRENT_SHEET_ID.with(|cell| cell.set(sheet_id as u64));
}

pub(crate) fn reset_thread_rng_counter() {
    BYTECODE_RNG_COUNTER.with(|cell| cell.set(0));
}

fn next_rng_draw() -> u64 {
    BYTECODE_RNG_COUNTER.with(|cell| {
        let draw = cell.get();
        cell.set(draw.wrapping_add(1));
        draw
    })
}

fn parse_value_from_text(s: &str) -> Result<f64, ErrorKind> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return Ok(0.0);
    }

    crate::coercion::datetime::parse_value_text(
        trimmed,
        thread_value_locale(),
        thread_now_utc(),
        thread_date_system(),
    )
    .map_err(|e| match e {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    })
}

pub fn eval_ast(
    expr: &Expr,
    grid: &dyn Grid,
    sheet_id: usize,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    // Match `Vm::eval`: the RNG draw counter is reset per top-level evaluation, and the
    // sheet context is set so deterministic volatile functions (e.g. RAND) can incorporate it.
    set_thread_current_sheet_id(sheet_id);
    reset_thread_rng_counter();

    let mut lexical_scopes: Vec<HashMap<Arc<str>, Value>> = Vec::new();
    // Match `Vm::eval`: top-level range references should deref dynamically (spill) instead of
    // remaining as a reference value.
    let v = eval_ast_inner(
        expr,
        grid,
        sheet_id,
        base,
        locale,
        &mut lexical_scopes,
        false,
    );
    deref_value_dynamic(v, grid, base)
}

fn eval_ast_inner(
    expr: &Expr,
    grid: &dyn Grid,
    sheet_id: usize,
    base: CellCoord,
    locale: &crate::LocaleConfig,
    lexical_scopes: &mut Vec<HashMap<Arc<str>, Value>>,
    allow_range: bool,
) -> Value {
    match expr {
        Expr::Literal(v) => match v {
            // `Value::Missing` is used during lowering as a placeholder for syntactically blank
            // arguments (e.g. `ADDRESS(1,1,,FALSE)`), but it must not be allowed to propagate as a
            // general runtime value (e.g. via `IF(FALSE,1,)`).
            //
            // Treat literal `Missing` as a normal blank value during expression evaluation.
            // Call sites that are evaluating *direct* function arguments should preserve Missing so
            // functions can distinguish omitted arguments from blank cell values.
            Value::Missing => Value::Empty,
            other => other.clone(),
        },
        Expr::CellRef(r) => {
            if allow_range {
                Value::Range(RangeRef::new(*r, *r))
            } else {
                grid.get_value(r.resolve(base))
            }
        }
        Expr::RangeRef(r) => Value::Range(*r),
        Expr::MultiRangeRef(r) => Value::MultiRange(r.clone()),
        Expr::SpillRange(inner) => {
            // The spill-range operator (`expr#`) evaluates its operand in a "reference context"
            // (i.e. it must preserve references rather than implicitly intersecting them).
            let v = eval_ast_inner(inner, grid, sheet_id, base, locale, lexical_scopes, true);
            apply_spill_range(v, grid, sheet_id, base)
        }
        Expr::NameRef(name) => {
            for scope in lexical_scopes.iter().rev() {
                if let Some(v) = scope.get(name) {
                    let v = v.clone();
                    // LET binding values are evaluated in "argument mode" (may preserve references).
                    // When a reference value is used in a scalar context, apply implicit intersection
                    // for single-cell references to match VM semantics.
                    if !allow_range && matches!(&v, Value::Range(r) if r.start == r.end) {
                        return apply_implicit_intersection(v, grid, base);
                    }
                    return v;
                }
            }
            Value::Error(ErrorKind::Name)
        }
        Expr::Unary { op, expr } => {
            let v = match op {
                UnaryOp::ImplicitIntersection => {
                    eval_ast_inner(expr, grid, sheet_id, base, locale, lexical_scopes, true)
                }
                UnaryOp::Plus | UnaryOp::Neg => {
                    eval_ast_inner(expr, grid, sheet_id, base, locale, lexical_scopes, false)
                }
            };
            match op {
                UnaryOp::ImplicitIntersection => apply_implicit_intersection(v, grid, base),
                _ => apply_unary(*op, v, grid, base),
            }
        }
        Expr::Binary { op, left, right } => {
            // Reference-algebra operators (union/intersection) must evaluate operands in a
            // reference context so `A1` behaves like a single-cell range.
            let allow_range = matches!(op, BinaryOp::Union | BinaryOp::Intersect);
            let l = eval_ast_inner(
                left,
                grid,
                sheet_id,
                base,
                locale,
                lexical_scopes,
                allow_range,
            );
            let r = eval_ast_inner(
                right,
                grid,
                sheet_id,
                base,
                locale,
                lexical_scopes,
                allow_range,
            );
            apply_binary(*op, l, r, grid, sheet_id, base)
        }
        Expr::FuncCall { func, args } => {
            if matches!(func, Function::Let) {
                if args.len() < 3 || args.len() % 2 == 0 {
                    return Value::Error(ErrorKind::Value);
                }

                let last = args.len() - 1;
                lexical_scopes.push(HashMap::new());
                for pair in args[..last].chunks_exact(2) {
                    let Expr::NameRef(name) = &pair[0] else {
                        lexical_scopes.pop();
                        return Value::Error(ErrorKind::Value);
                    };
                    // LET binding values are evaluated in "argument mode" (may preserve references).
                    let value = eval_ast_inner(
                        &pair[1],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        true,
                    );
                    lexical_scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(name.clone(), value);
                }
                let result = eval_ast_inner(
                    &args[last],
                    grid,
                    sheet_id,
                    base,
                    locale,
                    lexical_scopes,
                    allow_range,
                );
                lexical_scopes.pop();
                return result;
            }

            // Some logical/error/select functions are lazy in Excel: avoid evaluating unused
            // branches/fallbacks. This mirrors the bytecode VM which compiles these functions into
            // explicit control flow.
            match func {
                Function::If => {
                    if args.len() < 2 || args.len() > 3 {
                        return Value::Error(ErrorKind::Value);
                    }
                    let cond_val = eval_ast_inner(
                        &args[0],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        false,
                    );
                    // Match the bytecode VM: dereference single-cell ranges before coercing to bool
                    // so expressions like `IF(CHOOSE(1, A1, FALSE), ...)` behave like `IF(A1, ...)`.
                    let cond_val = deref_value_dynamic(cond_val, grid, base);
                    let cond = match coerce_to_bool(&cond_val) {
                        Ok(b) => b,
                        Err(e) => return Value::Error(e),
                    };
                    if cond {
                        return eval_ast_inner(
                            &args[1],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                    }
                    if args.len() == 3 {
                        return eval_ast_inner(
                            &args[2],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                    }
                    // Engine behavior: missing false branch defaults to FALSE (not blank).
                    return Value::Bool(false);
                }
                Function::Choose => {
                    if args.len() < 2 || args.len() > 255 {
                        return Value::Error(ErrorKind::Value);
                    }
                    let idx_val = eval_ast_inner(
                        &args[0],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        false,
                    );
                    let idx = match coerce_to_i64(&idx_val) {
                        Ok(i) => i,
                        Err(e) => return Value::Error(e),
                    };
                    // CHOOSE is 1-indexed.
                    if idx < 1 {
                        return Value::Error(ErrorKind::Value);
                    }
                    let idx_usize = match usize::try_from(idx) {
                        Ok(i) => i,
                        Err(_) => return Value::Error(ErrorKind::Value),
                    };
                    if idx_usize >= args.len() {
                        return Value::Error(ErrorKind::Value);
                    }
                    let choice_expr = &args[idx_usize];
                    // CHOOSE's value arguments may need to preserve references depending on
                    // surrounding context (e.g. `SUM(CHOOSE(1, A1, B1))`), so propagate
                    // `allow_range` to the selected branch.
                    return eval_ast_inner(
                        choice_expr,
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        allow_range,
                    );
                }
                Function::Ifs => {
                    if args.len() % 2 != 0 {
                        return Value::Error(ErrorKind::Value);
                    }
                    if args.len() < 2 {
                        return Value::Error(ErrorKind::Value);
                    }

                    for pair in args.chunks_exact(2) {
                        let cond_val = eval_ast_inner(
                            &pair[0],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                        let cond_val = deref_value_dynamic(cond_val, grid, base);
                        let cond = match coerce_to_bool(&cond_val) {
                            Ok(b) => b,
                            Err(e) => return Value::Error(e),
                        };
                        if cond {
                            return eval_ast_inner(
                                &pair[1],
                                grid,
                                sheet_id,
                                base,
                                locale,
                                lexical_scopes,
                                false,
                            );
                        }
                    }
                    return Value::Error(ErrorKind::NA);
                }
                Function::IfError => {
                    if args.len() != 2 {
                        return Value::Error(ErrorKind::Value);
                    }
                    let first = eval_ast_inner(
                        &args[0],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        false,
                    );
                    if matches!(first, Value::Error(_)) {
                        return eval_ast_inner(
                            &args[1],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                    }
                    return first;
                }
                Function::IfNa => {
                    if args.len() != 2 {
                        return Value::Error(ErrorKind::Value);
                    }
                    let first = eval_ast_inner(
                        &args[0],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        false,
                    );
                    if matches!(first, Value::Error(ErrorKind::NA)) {
                        return eval_ast_inner(
                            &args[1],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                    }
                    return first;
                }
                Function::Switch => {
                    if args.len() < 3 {
                        return Value::Error(ErrorKind::Value);
                    }

                    let expr_val = eval_ast_inner(
                        &args[0],
                        grid,
                        sheet_id,
                        base,
                        locale,
                        lexical_scopes,
                        false,
                    );
                    if let Value::Error(e) = expr_val {
                        return Value::Error(e);
                    }

                    let has_default = (args.len() - 1) % 2 != 0;
                    let pairs_end = if has_default {
                        args.len() - 1
                    } else {
                        args.len()
                    };
                    let pairs = &args[1..pairs_end];
                    let default = if has_default {
                        Some(&args[args.len() - 1])
                    } else {
                        None
                    };

                    if pairs.len() < 2 || pairs.len() % 2 != 0 {
                        return Value::Error(ErrorKind::Value);
                    }

                    for pair in pairs.chunks_exact(2) {
                        let case_val = eval_ast_inner(
                            &pair[0],
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                        let matches_val = apply_binary(
                            BinaryOp::Eq,
                            expr_val.clone(),
                            case_val,
                            grid,
                            sheet_id,
                            base,
                        );
                        let matches = match coerce_to_bool(&matches_val) {
                            Ok(b) => b,
                            Err(e) => return Value::Error(e),
                        };
                        if matches {
                            return eval_ast_inner(
                                &pair[1],
                                grid,
                                sheet_id,
                                base,
                                locale,
                                lexical_scopes,
                                false,
                            );
                        }
                    }

                    if let Some(default_expr) = default {
                        return eval_ast_inner(
                            default_expr,
                            grid,
                            sheet_id,
                            base,
                            locale,
                            lexical_scopes,
                            false,
                        );
                    }
                    return Value::Error(ErrorKind::NA);
                }
                _ => {}
            }

            // Evaluate arguments first (AST evaluation).
            let mut evaluated: SmallVec<[Value; 8]> = SmallVec::with_capacity(args.len());
            for (arg_idx, arg) in args.iter().enumerate() {
                // Preserve `Missing` for *direct* blank arguments so functions can distinguish a
                // syntactically omitted argument from a blank cell value.
                if matches!(arg, Expr::Literal(Value::Missing)) {
                    evaluated.push(Value::Missing);
                    continue;
                }

                // See `Compiler::compile_func_arg` for the rationale: certain functions treat a
                // single-cell reference passed directly as an argument as a range/reference value,
                // not a scalar.
                let allow_range = match func {
                    // AND/OR accept range arguments (which ignore text-like values), but direct
                    // *cell* references are treated as scalar arguments (so text-like values
                    // produce #VALUE!).
                    Function::And | Function::Or => true,
                    // XOR uses reference semantics for direct cell references (like ranges),
                    // matching the evaluator.
                    Function::Xor => true,
                    Function::Sum
                    | Function::Average
                    | Function::Min
                    | Function::Max
                    | Function::Count
                    | Function::CountA
                    | Function::CountBlank => true,
                    Function::CountIf => arg_idx == 0,
                    Function::SumIf | Function::AverageIf => arg_idx == 0 || arg_idx == 2,
                    Function::SumIfs
                    | Function::AverageIfs
                    | Function::MinIfs
                    | Function::MaxIfs => arg_idx == 0 || arg_idx % 2 == 1,
                    Function::CountIfs => arg_idx % 2 == 0,
                    Function::SumProduct => true,
                    Function::VLookup | Function::HLookup | Function::Match => arg_idx == 1,
                    Function::XMatch => arg_idx == 1,
                    Function::XLookup => arg_idx == 1 || arg_idx == 2,
                    Function::Offset => arg_idx == 0,
                    Function::Row | Function::Column | Function::Rows | Function::Columns => true,
                    _ => false,
                };

                let allow_range = if matches!(func, Function::And | Function::Or)
                    && matches!(arg, Expr::CellRef(_))
                {
                    false
                } else {
                    allow_range
                };

                evaluated.push(eval_ast_inner(
                    arg,
                    grid,
                    sheet_id,
                    base,
                    locale,
                    lexical_scopes,
                    allow_range,
                ));
            }
            call_function(func, &evaluated, grid, base, locale)
        }
        // The bytecode AST evaluator is a lightweight reference implementation for the bytecode
        // module's internal parser/tests. Higher-order constructs (LAMBDA/call) are evaluated
        // through the bytecode VM.
        Expr::Lambda { .. } | Expr::Call { .. } => Value::Error(ErrorKind::Name),
    }
}

pub(crate) fn apply_spill_range(
    v: Value,
    grid: &dyn Grid,
    sheet_id: usize,
    base: CellCoord,
) -> Value {
    match v {
        Value::Error(e) => Value::Error(e),
        Value::Range(r) => {
            let start = r.start.resolve(base);
            let end = r.end.resolve(base);
            if start.row != end.row || start.col != end.col {
                return Value::Error(ErrorKind::Value);
            }
            if !grid.in_bounds(start) {
                return Value::Error(ErrorKind::Ref);
            }

            let addr = crate::eval::CellAddr {
                row: start.row as u32,
                col: start.col as u32,
            };
            let Some(origin) = grid.spill_origin(&SheetId::Local(sheet_id), addr) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some((spill_start, spill_end)) =
                grid.spill_range(&SheetId::Local(sheet_id), origin)
            else {
                return Value::Error(ErrorKind::Ref);
            };

            Value::Range(RangeRef::new(
                Ref::new(spill_start.row as i32, spill_start.col as i32, true, true),
                Ref::new(spill_end.row as i32, spill_end.col as i32, true, true),
            ))
        }
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => {
                let start = only.range.start.resolve(base);
                let end = only.range.end.resolve(base);
                if start.row != end.row || start.col != end.col {
                    return Value::Error(ErrorKind::Value);
                }
                if !grid.in_bounds_on_sheet(&only.sheet, start) {
                    return Value::Error(ErrorKind::Ref);
                }

                let addr = crate::eval::CellAddr {
                    row: start.row as u32,
                    col: start.col as u32,
                };
                let Some(origin) = grid.spill_origin(&only.sheet, addr) else {
                    return Value::Error(ErrorKind::Ref);
                };
                let Some((spill_start, spill_end)) = grid.spill_range(&only.sheet, origin) else {
                    return Value::Error(ErrorKind::Ref);
                };

                let range = RangeRef::new(
                    Ref::new(spill_start.row as i32, spill_start.col as i32, true, true),
                    Ref::new(spill_end.row as i32, spill_end.col as i32, true, true),
                );
                if only.sheet == SheetId::Local(sheet_id) {
                    Value::Range(range)
                } else {
                    Value::MultiRange(MultiRangeRef::new(
                        vec![SheetRangeRef::new(only.sheet.clone(), range)].into(),
                    ))
                }
            }
            _ => Value::Error(ErrorKind::Value),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn coerce_to_number(v: &Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Empty | Value::Missing => Ok(0.0),
        Value::Text(s) => parse_value_from_text(s),
        Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Error(e) => Err(*e),
        // Dynamic arrays / range-as-scalar: treat as a spill attempt (engine semantics).
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Spill),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Err(ErrorKind::Ref),
            [_] => Err(ErrorKind::Spill),
            _ => Err(ErrorKind::Value),
        },
    }
}

pub(crate) fn coerce_to_bool(v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Bool(b) => Ok(*b),
        Value::Number(n) => Ok(*n != 0.0),
        Value::Empty | Value::Missing => Ok(false),
        Value::Text(s) => {
            let trimmed = s.trim();
            if trimmed.is_empty() {
                return Ok(false);
            }
            if trimmed.eq_ignore_ascii_case("TRUE") {
                return Ok(true);
            }
            if trimmed.eq_ignore_ascii_case("FALSE") {
                return Ok(false);
            }
            // Match evaluator semantics: if the text isn't a boolean literal, coerce it via the
            // same value parser used for numeric/date coercion.
            let n = parse_value_from_text(trimmed)?;
            Ok(n != 0.0)
        }
        Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Range(_) => Err(ErrorKind::Spill),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Err(ErrorKind::Ref),
            [_] => Err(ErrorKind::Spill),
            _ => Err(ErrorKind::Value),
        },
    }
}

fn excel_error_kind(e: ExcelError) -> ErrorKind {
    match e {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn excel_result_number(res: Result<f64, ExcelError>) -> Value {
    match res {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

fn excel_result_serial(res: Result<i32, ExcelError>) -> Value {
    match res {
        Ok(n) => Value::Number(n as f64),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

fn coerce_to_finite_number(v: &Value) -> Result<f64, ErrorKind> {
    let n = coerce_to_number(v)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

fn coerce_number_to_i32_trunc(n: f64) -> Result<i32, ErrorKind> {
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i32)
}

fn coerce_to_i32_trunc(v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(v)?;
    coerce_number_to_i32_trunc(n)
}

fn datevalue_from_value(value: &Value) -> Result<i32, ErrorKind> {
    match value {
        Value::Text(s) => crate::functions::date_time::datevalue(
            s,
            thread_value_locale(),
            thread_now_utc(),
            thread_date_system(),
        )
        .map_err(excel_error_kind),
        _ => {
            let n = coerce_to_finite_number(value)?;
            let serial = n.floor();
            if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
                return Err(ErrorKind::Num);
            }
            Ok(serial as i32)
        }
    }
}

fn datevalue_from_value_validated(value: &Value) -> Result<i32, ErrorKind> {
    let serial = datevalue_from_value(value)?;
    serial_to_ymd(serial, thread_date_system()).map_err(excel_error_kind)?;
    Ok(serial)
}

fn basis_from_optional_arg(arg: Option<&Value>) -> Result<i32, ErrorKind> {
    let Some(arg) = arg else {
        return Ok(0);
    };
    if matches!(arg, Value::Empty | Value::Missing) {
        return Ok(0);
    }

    let n = coerce_to_finite_number(arg)?;
    let basis = coerce_number_to_i32_trunc(n)?;
    if !(0..=4).contains(&basis) {
        return Err(ErrorKind::Num);
    }
    Ok(basis)
}

fn frequency_from_value(arg: &Value) -> Result<i32, ErrorKind> {
    let frequency = coerce_to_i32_trunc(arg)?;
    match frequency {
        1 | 2 | 4 => Ok(frequency),
        _ => Err(ErrorKind::Num),
    }
}

fn coerce_to_bool_finite(v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            Ok(*n != 0.0)
        }
        other => coerce_to_bool(other),
    }
}

fn matches_numeric_criteria(v: f64, criteria: NumericCriteria) -> bool {
    match criteria.op {
        CmpOp::Eq => v == criteria.rhs,
        CmpOp::Ne => v != criteria.rhs,
        CmpOp::Lt => v < criteria.rhs,
        CmpOp::Le => v <= criteria.rhs,
        CmpOp::Gt => v > criteria.rhs,
        CmpOp::Ge => v >= criteria.rhs,
    }
}

fn coerce_countif_value_to_number(v: &Value) -> Option<f64> {
    match v {
        Value::Number(n) => Some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Empty | Value::Missing => Some(0.0),
        Value::Text(s) => parse_number(s, thread_number_locale()).ok(),
        Value::Entity(v) => parse_number(v.display.as_str(), thread_number_locale()).ok(),
        Value::Record(v) => parse_number(v.display.as_str(), thread_number_locale()).ok(),
        Value::Lambda(_) => None,
        Value::Error(_) | Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => None,
    }
}

pub fn apply_implicit_intersection(v: Value, grid: &dyn Grid, base: CellCoord) -> Value {
    match v {
        Value::Error(e) => Value::Error(e),
        Value::Range(r) => {
            let range = r.resolve(base);
            if !range_in_bounds(grid, range) {
                return Value::Error(ErrorKind::Ref);
            }

            // Single-cell ranges return that cell.
            if range.row_start == range.row_end && range.col_start == range.col_end {
                let coord = CellCoord {
                    row: range.row_start,
                    col: range.col_start,
                };
                grid.record_reference(grid.sheet_id(), coord, coord);
                return grid.get_value(coord);
            }

            // 1D ranges intersect on the matching row/column.
            if range.col_start == range.col_end {
                if base.row >= range.row_start && base.row <= range.row_end {
                    let coord = CellCoord {
                        row: base.row,
                        col: range.col_start,
                    };
                    grid.record_reference(grid.sheet_id(), coord, coord);
                    return grid.get_value(coord);
                }
                return Value::Error(ErrorKind::Value);
            }

            if range.row_start == range.row_end {
                if base.col >= range.col_start && base.col <= range.col_end {
                    let coord = CellCoord {
                        row: range.row_start,
                        col: base.col,
                    };
                    grid.record_reference(grid.sheet_id(), coord, coord);
                    return grid.get_value(coord);
                }
                return Value::Error(ErrorKind::Value);
            }

            // 2D ranges intersect only if the current cell is within the rectangle.
            if base.row >= range.row_start
                && base.row <= range.row_end
                && base.col >= range.col_start
                && base.col <= range.col_end
            {
                grid.record_reference(grid.sheet_id(), base, base);
                return grid.get_value(base);
            }

            Value::Error(ErrorKind::Value)
        }
        Value::MultiRange(r) => {
            // Excel's implicit intersection on a multi-area reference is ambiguous; we approximate
            // by succeeding only when exactly one area intersects.
            let mut hit: Option<Value> = None;
            for area in r.areas.iter() {
                let v = apply_implicit_intersection_sheet_range(area, grid, base);
                if matches!(v, Value::Error(ErrorKind::Value)) {
                    continue;
                }
                if hit.is_some() {
                    return Value::Error(ErrorKind::Value);
                }
                hit = Some(v);
            }
            hit.unwrap_or(Value::Error(ErrorKind::Value))
        }
        other => other,
    }
}

fn apply_implicit_intersection_sheet_range(
    area: &SheetRangeRef,
    grid: &dyn Grid,
    base: CellCoord,
) -> Value {
    let range = area.range.resolve(base);
    if !range_in_bounds_on_sheet(grid, &area.sheet, range) {
        return Value::Error(ErrorKind::Ref);
    }

    // Single-cell ranges return that cell.
    if range.row_start == range.row_end && range.col_start == range.col_end {
        let coord = CellCoord {
            row: range.row_start,
            col: range.col_start,
        };
        grid.record_reference_on_sheet(&area.sheet, coord, coord);
        return grid.get_value_on_sheet(&area.sheet, coord);
    }

    // 1D ranges intersect on the matching row/column.
    if range.col_start == range.col_end {
        if base.row >= range.row_start && base.row <= range.row_end {
            let coord = CellCoord {
                row: base.row,
                col: range.col_start,
            };
            grid.record_reference_on_sheet(&area.sheet, coord, coord);
            return grid.get_value_on_sheet(&area.sheet, coord);
        }
        return Value::Error(ErrorKind::Value);
    }

    if range.row_start == range.row_end {
        if base.col >= range.col_start && base.col <= range.col_end {
            let coord = CellCoord {
                row: range.row_start,
                col: base.col,
            };
            grid.record_reference_on_sheet(&area.sheet, coord, coord);
            return grid.get_value_on_sheet(&area.sheet, coord);
        }
        return Value::Error(ErrorKind::Value);
    }

    // 2D ranges intersect only if the current cell is within the rectangle.
    if base.row >= range.row_start
        && base.row <= range.row_end
        && base.col >= range.col_start
        && base.col <= range.col_end
    {
        grid.record_reference_on_sheet(&area.sheet, base, base);
        return grid.get_value_on_sheet(&area.sheet, base);
    }

    Value::Error(ErrorKind::Value)
}

fn numeric_unary(op: UnaryOp, v: &Value) -> Value {
    let n = match coerce_to_number(v) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    match op {
        UnaryOp::Plus => Value::Number(n),
        UnaryOp::Neg => Value::Number(-n),
        UnaryOp::ImplicitIntersection => {
            unreachable!("implicit intersection requires Grid + base context")
        }
    }
}

fn numeric_binary(op: BinaryOp, left: &Value, right: &Value) -> Value {
    let ln = match coerce_to_number(left) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let rn = match coerce_to_number(right) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    match op {
        BinaryOp::Add => Value::Number(ln + rn),
        BinaryOp::Sub => Value::Number(ln - rn),
        BinaryOp::Mul => Value::Number(ln * rn),
        BinaryOp::Div => {
            if rn == 0.0 {
                Value::Error(ErrorKind::Div0)
            } else {
                Value::Number(ln / rn)
            }
        }
        BinaryOp::Pow => match crate::functions::math::power(ln, rn) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => {
            let total = match arr.rows.checked_mul(arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for v in arr.iter() {
                out.push(f(v));
            }
            Value::Array(ArrayValue::new(arr.rows, arr.cols, out))
        }
        other => f(other),
    }
}

fn elementwise_binary(left: &Value, right: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    match (left, right) {
        (Value::Array(left_arr), Value::Array(right_arr)) => {
            let out_rows = if left_arr.rows == right_arr.rows {
                left_arr.rows
            } else if left_arr.rows == 1 {
                right_arr.rows
            } else if right_arr.rows == 1 {
                left_arr.rows
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let out_cols = if left_arr.cols == right_arr.cols {
                left_arr.cols
            } else if left_arr.cols == 1 {
                right_arr.cols
            } else if right_arr.cols == 1 {
                left_arr.cols
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let total = match out_rows.checked_mul(out_cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }

            for row in 0..out_rows {
                let l_row = if left_arr.rows == 1 { 0 } else { row };
                let r_row = if right_arr.rows == 1 { 0 } else { row };
                for col in 0..out_cols {
                    let l_col = if left_arr.cols == 1 { 0 } else { col };
                    let r_col = if right_arr.cols == 1 { 0 } else { col };
                    let l = left_arr.get(l_row, l_col).unwrap_or(&Value::Empty);
                    let r = right_arr.get(r_row, r_col).unwrap_or(&Value::Empty);
                    out.push(f(l, r));
                }
            }

            Value::Array(ArrayValue::new(out_rows, out_cols, out))
        }
        (Value::Array(left_arr), right_scalar) => {
            let total = match left_arr.rows.checked_mul(left_arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for v in left_arr.iter() {
                out.push(f(v, right_scalar));
            }
            Value::Array(ArrayValue::new(left_arr.rows, left_arr.cols, out))
        }
        (left_scalar, Value::Array(right_arr)) => {
            let total = match right_arr.rows.checked_mul(right_arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for v in right_arr.iter() {
                out.push(f(left_scalar, v));
            }
            Value::Array(ArrayValue::new(right_arr.rows, right_arr.cols, out))
        }
        (left_scalar, right_scalar) => f(left_scalar, right_scalar),
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct LiftShape {
    rows: usize,
    cols: usize,
}

impl LiftShape {
    fn is_1x1(self) -> bool {
        self.rows == 1 && self.cols == 1
    }
}

fn lift_shape(value: &Value) -> Option<LiftShape> {
    match value {
        Value::Array(arr) => Some(LiftShape {
            rows: arr.rows,
            cols: arr.cols,
        }),
        _ => None,
    }
}

/// Excel-style array-lifting shape inference used by many scalar functions:
/// - Scalars broadcast over arrays.
/// - 1x1 arrays broadcast over larger arrays.
/// - Other array shapes must match exactly.
fn lift_dominant_shape(values: &[&Value]) -> Result<Option<LiftShape>, ErrorKind> {
    let mut dominant: Option<LiftShape> = None;
    let mut saw_array = false;

    for value in values {
        let Some(shape) = lift_shape(value) else {
            continue;
        };
        saw_array = true;

        if shape.is_1x1() {
            continue;
        }

        match dominant {
            None => dominant = Some(shape),
            Some(existing) if existing == shape => {}
            Some(_) => return Err(ErrorKind::Value),
        }
    }

    if dominant.is_some() {
        return Ok(dominant);
    }

    if saw_array {
        return Ok(Some(LiftShape { rows: 1, cols: 1 }));
    }

    Ok(None)
}

fn lift_broadcast_compatible(value: &Value, target: LiftShape) -> bool {
    match value {
        Value::Array(arr) => {
            (arr.rows == target.rows && arr.cols == target.cols) || (arr.rows == 1 && arr.cols == 1)
        }
        _ => true,
    }
}

fn lift_element_at<'a>(value: &'a Value, target: LiftShape, idx: usize) -> &'a Value {
    match value {
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                return arr.values.get(0).unwrap_or(&Value::Empty);
            }
            debug_assert_eq!(arr.rows, target.rows);
            debug_assert_eq!(arr.cols, target.cols);
            arr.values.get(idx).unwrap_or(&Value::Empty)
        }
        other => other,
    }
}

fn lift2(a: &Value, b: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    let Some(shape) = (match lift_dominant_shape(&[a, b]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    }) else {
        return f(a, b);
    };

    if !lift_broadcast_compatible(a, shape) || !lift_broadcast_compatible(b, shape) {
        return Value::Error(ErrorKind::Value);
    }

    let len = match shape.rows.checked_mul(shape.cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if len > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(len).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for idx in 0..len {
        let av = lift_element_at(a, shape, idx);
        let bv = lift_element_at(b, shape, idx);
        out.push(f(av, bv));
    }
    Value::Array(ArrayValue::new(shape.rows, shape.cols, out))
}

fn deref_range_dynamic(grid: &dyn Grid, range: ResolvedRange) -> Value {
    if !range_in_bounds(grid, range) {
        return Value::Error(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range.rows() == 1 && range.cols() == 1 {
        return grid.get_value(CellCoord {
            row: range.row_start,
            col: range.col_start,
        });
    }

    let rows = match usize::try_from(range.rows()) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Spill),
    };
    let cols = match usize::try_from(range.cols()) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Spill),
    };
    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            values.push(grid.get_value(CellCoord { row, col }));
        }
    }
    Value::Array(ArrayValue::new(rows, cols, values))
}

fn deref_range_dynamic_on_sheet(grid: &dyn Grid, sheet: &SheetId, range: ResolvedRange) -> Value {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Value::Error(ErrorKind::Ref);
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range.rows() == 1 && range.cols() == 1 {
        return grid.get_value_on_sheet(
            sheet,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
        );
    }

    let rows = match usize::try_from(range.rows()) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Spill),
    };
    let cols = match usize::try_from(range.cols()) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Spill),
    };
    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            values.push(grid.get_value_on_sheet(sheet, CellCoord { row, col }));
        }
    }
    Value::Array(ArrayValue::new(rows, cols, values))
}

pub(crate) fn deref_value_dynamic(v: Value, grid: &dyn Grid, base: CellCoord) -> Value {
    match v {
        Value::Range(r) => deref_range_dynamic(grid, r.resolve(base)),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => deref_range_dynamic_on_sheet(grid, &only.sheet, only.range.resolve(base)),
            // Discontiguous unions cannot be represented as a single rectangular spill.
            _ => Value::Error(ErrorKind::Value),
        },
        other => other,
    }
}

pub fn apply_unary(op: UnaryOp, v: Value, grid: &dyn Grid, base: CellCoord) -> Value {
    let v = deref_value_dynamic(v, grid, base);
    elementwise_unary(&v, |elem| numeric_unary(op, elem))
}

pub fn apply_binary(
    op: BinaryOp,
    left: Value,
    right: Value,
    grid: &dyn Grid,
    sheet_id: usize,
    base: CellCoord,
) -> Value {
    match op {
        BinaryOp::Union => reference_union(left, right, sheet_id, base),
        BinaryOp::Intersect => reference_intersect(left, right, sheet_id, base),
        BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div | BinaryOp::Pow => {
            let left = deref_value_dynamic(left, grid, base);
            let right = deref_value_dynamic(right, grid, base);
            elementwise_binary(&left, &right, |a, b| numeric_binary(op, a, b))
        }
        BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
            let left = deref_value_dynamic(left, grid, base);
            let right = deref_value_dynamic(right, grid, base);
            elementwise_binary(&left, &right, |a, b| excel_compare(a, b, op))
        }
    }
}

fn value_into_reference_areas(value: Value, sheet_id: usize) -> Result<Vec<SheetRangeRef>, Value> {
    match value {
        Value::Range(r) => Ok(vec![SheetRangeRef::new(SheetId::Local(sheet_id), r)]),
        Value::MultiRange(r) => Ok(r.areas.iter().cloned().collect()),
        Value::Error(e) => Err(Value::Error(e)),
        _ => Err(Value::Error(ErrorKind::Value)),
    }
}

fn sort_reference_areas(areas: &mut [SheetRangeRef], base: CellCoord) {
    areas.sort_by(|a, b| {
        a.sheet.cmp(&b.sheet).then_with(|| {
            let ra = a.range.resolve(base);
            let rb = b.range.resolve(base);
            (ra.row_start, ra.col_start, ra.row_end, ra.col_end).cmp(&(
                rb.row_start,
                rb.col_start,
                rb.row_end,
                rb.col_end,
            ))
        })
    });
}

fn reference_union(left: Value, right: Value, sheet_id: usize, base: CellCoord) -> Value {
    let mut left = match value_into_reference_areas(left, sheet_id) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let right = match value_into_reference_areas(right, sheet_id) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let Some(first) = left.first() else {
        return Value::Error(ErrorKind::Ref);
    };
    let expected_sheet = first.sheet.clone();
    if left.iter().any(|r| r.sheet != expected_sheet)
        || right.iter().any(|r| r.sheet != expected_sheet)
    {
        return Value::Error(ErrorKind::Ref);
    }

    left.extend(right);
    sort_reference_areas(&mut left, base);

    match left.as_slice() {
        [] => Value::Error(ErrorKind::Ref),
        [only] if matches!(&only.sheet, SheetId::Local(id) if *id == sheet_id) => {
            Value::Range(only.range)
        }
        [only] => Value::MultiRange(MultiRangeRef::new(vec![only.clone()].into())),
        _ => Value::MultiRange(MultiRangeRef::new(left.into())),
    }
}

fn reference_intersect(left: Value, right: Value, sheet_id: usize, base: CellCoord) -> Value {
    let left = match value_into_reference_areas(left, sheet_id) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let right = match value_into_reference_areas(right, sheet_id) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let Some(first) = left.first() else {
        return Value::Error(ErrorKind::Ref);
    };
    let expected_sheet = first.sheet.clone();
    if left.iter().any(|r| r.sheet != expected_sheet)
        || right.iter().any(|r| r.sheet != expected_sheet)
    {
        return Value::Error(ErrorKind::Ref);
    }

    let mut out: Vec<SheetRangeRef> = Vec::new();
    for a in &left {
        let ra = a.range.resolve(base);
        for b in &right {
            let rb = b.range.resolve(base);
            let Some(intersection) = intersect_ranges(ra, rb) else {
                continue;
            };
            let start = Ref::new(intersection.row_start, intersection.col_start, true, true);
            let end = Ref::new(intersection.row_end, intersection.col_end, true, true);
            out.push(SheetRangeRef::new(
                expected_sheet.clone(),
                RangeRef::new(start, end),
            ));
        }
    }

    if out.is_empty() {
        return Value::Error(ErrorKind::Null);
    }
    sort_reference_areas(&mut out, base);

    match out.as_slice() {
        [only] if matches!(&only.sheet, SheetId::Local(id) if *id == sheet_id) => {
            Value::Range(only.range)
        }
        [only] => Value::MultiRange(MultiRangeRef::new(vec![only.clone()].into())),
        _ => Value::MultiRange(MultiRangeRef::new(out.into())),
    }
}

#[inline]
fn intersect_ranges(a: ResolvedRange, b: ResolvedRange) -> Option<ResolvedRange> {
    let row_start = a.row_start.max(b.row_start);
    let row_end = a.row_end.min(b.row_end);
    if row_start > row_end {
        return None;
    }
    let col_start = a.col_start.max(b.col_start);
    let col_end = a.col_end.min(b.col_end);
    if col_start > col_end {
        return None;
    }
    Some(ResolvedRange {
        row_start,
        row_end,
        col_start,
        col_end,
    })
}

fn excel_compare(left: &Value, right: &Value, op: BinaryOp) -> Value {
    let ord = match excel_order(left, right) {
        Ok(ord) => ord,
        Err(e) => return Value::Error(e),
    };

    let result = match op {
        BinaryOp::Eq => ord == Ordering::Equal,
        BinaryOp::Ne => ord != Ordering::Equal,
        BinaryOp::Lt => ord == Ordering::Less,
        BinaryOp::Le => ord != Ordering::Greater,
        BinaryOp::Gt => ord == Ordering::Greater,
        BinaryOp::Ge => ord != Ordering::Less,
        _ => return Value::Error(ErrorKind::Value),
    };

    Value::Bool(result)
}

fn excel_order(left: &Value, right: &Value) -> Result<Ordering, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(*e);
    }
    if let Value::Error(e) = right {
        return Err(*e);
    }

    fn normalize_rich(value: &Value) -> Result<Value, ErrorKind> {
        match value {
            Value::Entity(v) => Ok(Value::Text(Arc::from(v.display.as_str()))),
            Value::Record(v) => {
                if let Some(display_field) = v.display_field.as_deref() {
                    if let Some(field_value) = v.get_field_case_insensitive(display_field) {
                        let s = field_value.coerce_to_string().map_err(ErrorKind::from)?;
                        return Ok(Value::Text(Arc::from(s.as_str())));
                    }
                }
                Ok(Value::Text(Arc::from(v.display.as_str())))
            }
            other => Ok(other.clone()),
        }
    }

    let left = normalize_rich(left)?;
    let right = normalize_rich(right)?;
    if matches!(
        left,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_)
    ) || matches!(
        right,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_)
    ) {
        return Err(ErrorKind::Value);
    }

    fn text_like_str(value: &Value) -> Option<&str> {
        match value {
            Value::Text(s) => Some(s.as_ref()),
            _ => None,
        }
    }

    // Blank coerces to the other type for comparisons.
    let (l, r) = match (&left, &right) {
        (Value::Empty | Value::Missing, Value::Number(_)) => (Value::Number(0.0), right.clone()),
        (Value::Number(_), Value::Empty | Value::Missing) => (left.clone(), Value::Number(0.0)),
        (Value::Empty | Value::Missing, Value::Bool(_)) => (Value::Bool(false), right.clone()),
        (Value::Bool(_), Value::Empty | Value::Missing) => (left.clone(), Value::Bool(false)),
        (Value::Empty | Value::Missing, other) if text_like_str(other).is_some() => {
            (Value::Text(Arc::from("")), right.clone())
        }
        (other, Value::Empty | Value::Missing) if text_like_str(other).is_some() => {
            (left.clone(), Value::Text(Arc::from("")))
        }
        _ => (left.clone(), right.clone()),
    };

    Ok(match (l, r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(&b).unwrap_or(Ordering::Equal),
        (a, b) if text_like_str(&a).is_some() && text_like_str(&b).is_some() => {
            cmp_case_insensitive(text_like_str(&a).unwrap(), text_like_str(&b).unwrap())
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(&b),
        // Type precedence (approximate Excel): numbers < text < booleans.
        (Value::Number(_), b) if text_like_str(&b).is_some() || matches!(b, Value::Bool(_)) => {
            Ordering::Less
        }
        (a, Value::Bool(_)) if text_like_str(&a).is_some() => Ordering::Less,
        (a, Value::Number(_)) if text_like_str(&a).is_some() => Ordering::Greater,
        (Value::Bool(_), b) if matches!(b, Value::Number(_)) || text_like_str(&b).is_some() => {
            Ordering::Greater
        }
        // Blank should have been coerced above.
        (Value::Empty | Value::Missing, Value::Empty | Value::Missing) => Ordering::Equal,
        (Value::Empty | Value::Missing, _) => Ordering::Less,
        (_, Value::Empty | Value::Missing) => Ordering::Greater,
        // Errors are handled above.
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        // Arrays/ranges/lambdas are rejected above.
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Range(_), _)
        | (_, Value::Range(_))
        | (Value::MultiRange(_), _)
        | (_, Value::MultiRange(_))
        | (Value::Lambda(_), _)
        | (_, Value::Lambda(_)) => Ordering::Equal,
        _ => Ordering::Equal,
    })
}

pub fn call_function(
    func: &Function,
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    match func {
        Function::FieldAccess => fn_fieldaccess(args, grid, base),
        // LET is lowered to bytecode locals by the compiler; it should not be invoked via the
        // generic function-call path (its "name" arguments are not evaluated values).
        Function::Let => Value::Error(ErrorKind::Value),
        // ISOMITTED requires access to the lambda invocation frame, so it is compiled into direct
        // local loads by the bytecode compiler. It should never be invoked via `CallFunc`.
        Function::IsOmitted => Value::Error(ErrorKind::Value),
        Function::True => fn_true(args),
        Function::False => fn_false(args),
        Function::If => fn_if(args),
        Function::Choose => fn_choose(args),
        Function::Ifs => fn_ifs(args),
        Function::And => fn_and(args, grid, base),
        Function::Or => fn_or(args, grid, base),
        Function::Xor => fn_xor(args, grid, base),
        Function::IfError => fn_iferror(args),
        Function::IfNa => fn_ifna(args),
        Function::IsError => fn_iserror(args, grid, base),
        Function::IsNa => fn_isna(args, grid, base),
        Function::Na => fn_na(args),
        Function::Switch => fn_switch(args, grid, base),
        Function::Sum => fn_sum(args, grid, base),
        Function::SumIf => fn_sumif(args, grid, base, locale),
        Function::SumIfs => fn_sumifs(args, grid, base, locale),
        Function::Average => fn_average(args, grid, base),
        Function::AverageIf => fn_averageif(args, grid, base, locale),
        Function::AverageIfs => fn_averageifs(args, grid, base, locale),
        Function::Min => fn_min(args, grid, base),
        Function::MinIfs => fn_minifs(args, grid, base, locale),
        Function::Max => fn_max(args, grid, base),
        Function::MaxIfs => fn_maxifs(args, grid, base, locale),
        Function::Count => fn_count(args, grid, base),
        Function::CountA => fn_counta(args, grid, base),
        Function::CountBlank => fn_countblank(args, grid, base),
        Function::CountIf => fn_countif(args, grid, base, locale),
        Function::CountIfs => fn_countifs(args, grid, base, locale),
        Function::SumProduct => fn_sumproduct(args, grid, base),
        Function::VLookup => fn_vlookup(args, grid, base),
        Function::HLookup => fn_hlookup(args, grid, base),
        Function::Match => fn_match(args, grid, base),
        Function::Abs => fn_abs(args, grid, base),
        Function::Int => fn_int(args, grid, base),
        Function::Round => fn_round(args, grid, base),
        Function::RoundUp => fn_roundup(args, grid, base),
        Function::RoundDown => fn_rounddown(args, grid, base),
        Function::Mod => fn_mod(args, grid, base),
        Function::Sign => fn_sign(args, grid, base),
        Function::Db => fn_db(args),
        Function::Vdb => fn_vdb(args),
        Function::CoupDayBs => fn_coupdaybs(args),
        Function::CoupDays => fn_coupdays(args),
        Function::CoupDaysNc => fn_coupdaysnc(args),
        Function::CoupNcd => fn_coupncd(args),
        Function::CoupNum => fn_coupnum(args),
        Function::CoupPcd => fn_couppcd(args),
        Function::Price => fn_price(args),
        Function::Yield => fn_yield(args),
        Function::Duration => fn_duration(args),
        Function::MDuration => fn_mduration(args),
        Function::Accrint => fn_accrint(args),
        Function::Accrintm => fn_accrintm(args),
        Function::Disc => fn_disc(args),
        Function::PriceDisc => fn_pricedisc(args),
        Function::YieldDisc => fn_yielddisc(args),
        Function::Intrate => fn_intrate(args),
        Function::Received => fn_received(args),
        Function::PriceMat => fn_pricemat(args),
        Function::YieldMat => fn_yieldmat(args),
        Function::TbillEq => fn_tbilleq(args),
        Function::TbillPrice => fn_tbillprice(args),
        Function::TbillYield => fn_tbillyield(args),
        Function::OddFPrice => fn_oddfprice(args),
        Function::OddFYield => fn_oddfyield(args),
        Function::OddLPrice => fn_oddlprice(args),
        Function::OddLYield => fn_oddlyield(args),
        Function::ConcatOp => fn_concat_op(args, grid, base),
        Function::Concat => fn_concat(args, grid, base),
        Function::Concatenate => fn_concatenate(args, grid, base),
        Function::Rand => fn_rand(args, base),
        Function::RandBetween => fn_randbetween(args, base),
        Function::Not => fn_not(args, grid, base),
        Function::IsBlank => fn_isblank(args, grid, base),
        Function::IsNumber => fn_isnumber(args, grid, base),
        Function::IsText => fn_istext(args, grid, base),
        Function::IsLogical => fn_islogical(args, grid, base),
        Function::IsErr => fn_iserr(args, grid, base),
        Function::Type => fn_type(args, grid, base),
        Function::ErrorType => fn_error_type(args, grid, base),
        Function::N => fn_n(args, grid, base),
        Function::T => fn_t(args, grid, base),
        Function::Now => fn_now(args),
        Function::Today => fn_today(args),
        Function::Row => fn_row(args, grid, base),
        Function::Column => fn_column(args, grid, base),
        Function::Rows => fn_rows(args, base),
        Function::Columns => fn_columns(args, base),
        Function::Address => fn_address(args, grid, base),
        Function::Offset => fn_offset(args, grid, base),
        Function::Indirect => fn_indirect(args, grid, base),
        Function::XMatch => fn_xmatch(args, grid, base),
        Function::XLookup => fn_xlookup(args, grid, base),
        Function::Unknown(_) => Value::Error(ErrorKind::Name),
    }
}

fn fn_true(args: &[Value]) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    Value::Bool(true)
}

fn fn_false(args: &[Value]) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    Value::Bool(false)
}

fn fn_fieldaccess(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }

    // Match the AST evaluator's behavior (see `functions::builtins_rich_values::_FIELDACCESS`):
    // - The base expression is evaluated with reference semantics and then dereferenced into either
    //   a scalar (single-cell) or array (multi-cell) value.
    // - The field argument is evaluated as a scalar (with implicit intersection for references).
    let base_val = deref_value_dynamic(args[0].clone(), grid, base);

    let field_val = match args[1] {
        // `_FIELDACCESS` is an internal lowering builtin; callers may still provide a reference as
        // the field name via direct calls, so apply implicit intersection like normal scalar args.
        Value::Range(_) | Value::MultiRange(_) => {
            apply_implicit_intersection(args[1].clone(), grid, base)
        }
        _ => args[1].clone(),
    };

    let field = match &field_val {
        Value::Error(e) => return Value::Error(*e),
        Value::Text(s) => s.to_string(),
        other => match coerce_to_string(other) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        },
    };
    // Preserve the key exactly as written (including leading/trailing whitespace) to match the
    // AST evaluator's `.["..."]` semantics.
    //
    // This lets selectors like `A1.[" Price "]` address keys that include whitespace.
    let field_key = field.as_str();
    if field_key.trim().is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    match base_val {
        Value::Error(e) => Value::Error(e),
        Value::Array(arr) => {
            let mut out = Vec::with_capacity(arr.values.len());
            for elem in arr.iter() {
                out.push(fieldaccess_scalar(elem, field_key));
            }
            Value::Array(ArrayValue::new(arr.rows, arr.cols, out))
        }
        other => fieldaccess_scalar(&other, field_key),
    }
}

fn fieldaccess_scalar(base: &Value, field: &str) -> Value {
    match base {
        Value::Error(e) => Value::Error(*e),
        Value::Entity(entity) => match entity.get_field_case_insensitive(field) {
            Some(v) => engine_value_to_bytecode(v),
            None => Value::Error(ErrorKind::Field),
        },
        Value::Record(record) => match record.get_field_case_insensitive(field) {
            Some(v) => engine_value_to_bytecode(v),
            None => Value::Error(ErrorKind::Field),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn engine_value_to_bytecode(value: EngineValue) -> Value {
    match value {
        EngineValue::Number(n) => Value::Number(n),
        EngineValue::Bool(b) => Value::Bool(b),
        EngineValue::Text(s) => Value::Text(Arc::from(s)),
        EngineValue::Entity(e) => Value::Entity(Arc::new(e)),
        EngineValue::Record(r) => Value::Record(Arc::new(r)),
        EngineValue::Blank => Value::Empty,
        EngineValue::Error(e) => Value::Error(e.into()),
        EngineValue::Array(arr) => {
            let total = match arr.rows.checked_mul(arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for v in arr.values {
                values.push(engine_value_to_bytecode(v));
            }
            Value::Array(ArrayValue::new(arr.rows, arr.cols, values))
        }
        // Reference-producing values are not expected inside rich-value field payloads; treat them
        // as scalar type errors when surfaced.
        EngineValue::Reference(_) | EngineValue::ReferenceUnion(_) => {
            Value::Error(ErrorKind::Value)
        }
        // Lambdas cannot be represented in the bytecode runtime value model; match the engine's
        // cell-value conversion by surfacing `#CALC!`.
        EngineValue::Lambda(_) => Value::Error(ErrorKind::Calc),
        EngineValue::Spill { .. } => Value::Error(ErrorKind::Spill),
    }
}

fn fn_rand(args: &[Value], base: CellCoord) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let bits = volatile_rand_u64(base) >> 11; // 53 bits.
    Value::Number((bits as f64) / ((1u64 << 53) as f64))
}

fn fn_randbetween(args: &[Value], base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let bottom = match coerce_to_number(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let top = match coerce_to_number(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if !bottom.is_finite() || !top.is_finite() {
        return Value::Error(ErrorKind::Num);
    }

    let low_f = bottom.ceil();
    let high_f = top.floor();
    if low_f < (i64::MIN as f64)
        || low_f > (i64::MAX as f64)
        || high_f < (i64::MIN as f64)
        || high_f > (i64::MAX as f64)
    {
        return Value::Error(ErrorKind::Num);
    }

    let low = low_f as i64;
    let high = high_f as i64;
    if low > high {
        return Value::Error(ErrorKind::Num);
    }

    let span = match high.checked_sub(low).and_then(|d| d.checked_add(1)) {
        Some(v) if v > 0 => v as u64,
        _ => return Value::Error(ErrorKind::Num),
    };

    let offset = volatile_rand_u64_below(span, base) as i64;
    Value::Number((low + offset) as f64)
}

fn volatile_rand_u64_below(span: u64, base: CellCoord) -> u64 {
    if span <= 1 {
        return 0;
    }

    let zone = (u64::MAX / span) * span;
    loop {
        let v = volatile_rand_u64(base);
        if v < zone {
            return v % span;
        }
    }
}

fn volatile_rand_u64(base: CellCoord) -> u64 {
    let draw = next_rng_draw();

    let mut seed = thread_recalc_id();
    seed ^= thread_current_sheet_id().wrapping_mul(0x9e3779b97f4a7c15);
    seed ^= (base.row as u64).wrapping_mul(0xbf58476d1ce4e5b9);
    seed ^= (base.col as u64).wrapping_mul(0x94d049bb133111eb);
    seed ^= draw.wrapping_mul(0x3c79ac492ba7b653);
    splitmix64(seed)
}

fn splitmix64(mut state: u64) -> u64 {
    // A simple, fast mixer with good statistical properties (used as a deterministic
    // PRNG building block). The transform is bijective over u64, making it a good fit
    // for per-cell deterministic RNG.
    state = state.wrapping_add(0x9e3779b97f4a7c15);
    state = (state ^ (state >> 30)).wrapping_mul(0xbf58476d1ce4e5b9);
    state = (state ^ (state >> 27)).wrapping_mul(0x94d049bb133111eb);
    state ^ (state >> 31)
}

fn fn_today(args: &[Value]) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let now = thread_now_utc();
    let date = now.date_naive();
    match ymd_to_serial(
        ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
        thread_date_system(),
    ) {
        Ok(serial) => Value::Number(serial as f64),
        Err(_) => Value::Error(ErrorKind::Num),
    }
}

fn fn_now(args: &[Value]) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let now = thread_now_utc();
    let date = now.date_naive();
    let base = match ymd_to_serial(
        ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
        thread_date_system(),
    ) {
        Ok(serial) => serial as f64,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let seconds = now.time().num_seconds_from_midnight() as f64
        + (now.time().nanosecond() as f64 / 1_000_000_000.0);
    Value::Number(base + seconds / 86_400.0)
}

fn fn_db(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let cost = match coerce_to_number(&args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match coerce_to_number(&args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match coerce_to_number(&args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let period = match coerce_to_number(&args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let month = if args.len() == 5 {
        match coerce_to_number(&args[4]) {
            Ok(n) => Some(n),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match crate::functions::financial::db(cost, salvage, life, period, month) {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(match e {
            crate::error::ExcelError::Div0 => ErrorKind::Div0,
            crate::error::ExcelError::Value => ErrorKind::Value,
            crate::error::ExcelError::Num => ErrorKind::Num,
        }),
    }
}

fn fn_vdb(args: &[Value]) -> Value {
    if args.len() < 5 || args.len() > 7 {
        return Value::Error(ErrorKind::Value);
    }

    let cost = match coerce_to_number(&args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match coerce_to_number(&args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match coerce_to_number(&args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let start = match coerce_to_number(&args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let end = match coerce_to_number(&args[4]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let factor = match args.get(5) {
        None => None,
        Some(v) => match coerce_to_number(v) {
            Ok(n) => Some(n),
            Err(e) => return Value::Error(e),
        },
    };
    let no_switch = match args.get(6) {
        None => None,
        Some(v) => match coerce_to_number(v) {
            Ok(n) => Some(n),
            Err(e) => return Value::Error(e),
        },
    };

    match crate::functions::financial::vdb(cost, salvage, life, start, end, factor, no_switch) {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(match e {
            crate::error::ExcelError::Div0 => ErrorKind::Div0,
            crate::error::ExcelError::Value => ErrorKind::Value,
            crate::error::ExcelError::Num => ErrorKind::Num,
        }),
    }
}

// ---------------------------------------------------------------------
// Securities / bond / day-count financial functions
// ---------------------------------------------------------------------

fn fn_coupdaybs(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::coupdaybs(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_coupdays(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::coupdays(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_coupdaysnc(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::coupdaysnc(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_coupncd(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_serial(crate::functions::financial::coupncd(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_coupnum(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::coupnum(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_couppcd(args: &[Value]) -> Value {
    if args.len() != 3 && args.len() != 4 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(3) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_serial(crate::functions::financial::couppcd(
        settlement,
        maturity,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_price(args: &[Value]) -> Value {
    if args.len() != 6 && args.len() != 7 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(6) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::price(
        settlement,
        maturity,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_yield(args: &[Value]) -> Value {
    if args.len() != 6 && args.len() != 7 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(6) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::yield_rate(
        settlement,
        maturity,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_duration(args: &[Value]) -> Value {
    if args.len() != 5 && args.len() != 6 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(5) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::duration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_mduration(args: &[Value]) -> Value {
    if args.len() != 5 && args.len() != 6 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(5) {
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::mduration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_accrintm(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let issue = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let par = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match args.get(4) {
        Some(v) if matches!(v, Value::Empty) => 0,
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(crate::functions::financial::accrintm(
        issue,
        settlement,
        rate,
        par,
        basis,
        thread_date_system(),
    ))
}

fn fn_accrint(args: &[Value]) -> Value {
    if args.len() < 6 || args.len() > 8 {
        return Value::Error(ErrorKind::Value);
    }

    let issue = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_interest = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let par = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(6) {
        Some(v) if matches!(v, Value::Empty) => 0,
        Some(v) => match coerce_to_i32_trunc(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    let calc_method = match args.get(7) {
        Some(v) if matches!(v, Value::Empty) => false,
        Some(v) => match coerce_to_bool_finite(v) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => false,
    };

    excel_result_number(crate::functions::financial::accrint(
        issue,
        first_interest,
        settlement,
        rate,
        par,
        frequency,
        basis,
        calc_method,
        thread_date_system(),
    ))
}

fn fn_disc(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::disc(
        settlement,
        maturity,
        pr,
        redemption,
        basis,
        thread_date_system(),
    ))
}

fn fn_pricedisc(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::pricedisc(
        settlement,
        maturity,
        discount,
        redemption,
        basis,
        thread_date_system(),
    ))
}

fn fn_yielddisc(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::yielddisc(
        settlement,
        maturity,
        pr,
        redemption,
        basis,
        thread_date_system(),
    ))
}

fn fn_intrate(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let investment = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::intrate(
        settlement,
        maturity,
        investment,
        redemption,
        basis,
        thread_date_system(),
    ))
}

fn fn_received(args: &[Value]) -> Value {
    if args.len() != 4 && args.len() != 5 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let investment = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::received(
        settlement,
        maturity,
        investment,
        discount,
        basis,
        thread_date_system(),
    ))
}

fn fn_pricemat(args: &[Value]) -> Value {
    if args.len() != 5 && args.len() != 6 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::pricemat(
        settlement,
        maturity,
        issue,
        rate,
        yld,
        basis,
        thread_date_system(),
    ))
}

fn fn_yieldmat(args: &[Value]) -> Value {
    if args.len() != 5 && args.len() != 6 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::yieldmat(
        settlement,
        maturity,
        issue,
        rate,
        pr,
        basis,
        thread_date_system(),
    ))
}

fn fn_tbillprice(args: &[Value]) -> Value {
    if args.len() != 3 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::tbillprice(
        settlement, maturity, discount,
    ))
}

fn fn_tbillyield(args: &[Value]) -> Value {
    if args.len() != 3 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::tbillyield(
        settlement, maturity, pr,
    ))
}

fn fn_tbilleq(args: &[Value]) -> Value {
    if args.len() != 3 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::tbilleq(
        settlement, maturity, discount,
    ))
}

fn fn_oddfprice(args: &[Value]) -> Value {
    if args.len() != 8 && args.len() != 9 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value_validated(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value_validated(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value_validated(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_from_value_validated(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(&args[7]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_oddfyield(args: &[Value]) -> Value {
    if args.len() != 8 && args.len() != 9 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value_validated(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value_validated(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value_validated(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_from_value_validated(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(&args[7]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_oddlprice(args: &[Value]) -> Value {
    if args.len() != 7 && args.len() != 8 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value_validated(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value_validated(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_from_value_validated(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(&args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_oddlyield(args: &[Value]) -> Value {
    if args.len() != 7 && args.len() != 8 {
        return Value::Error(ErrorKind::Value);
    }

    let settlement = match datevalue_from_value_validated(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value_validated(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_from_value_validated(&args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(&args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(&args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(&args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(&args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(crate::functions::financial::oddlyield(
        settlement,
        maturity,
        last_interest,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        thread_date_system(),
    ))
}

fn fn_abs(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    let f = |elem: &Value| match coerce_to_number(elem) {
        Ok(n) => Value::Number(n.abs()),
        Err(e) => Value::Error(e),
    };
    match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            let v = deref_value_dynamic(args[0].clone(), grid, base);
            elementwise_unary(&v, f)
        }
        other => elementwise_unary(other, f),
    }
}

fn fn_int(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    let f = |elem: &Value| match coerce_to_number(elem) {
        Ok(n) => Value::Number(n.floor()),
        Err(e) => Value::Error(e),
    };
    match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            let v = deref_value_dynamic(args[0].clone(), grid, base);
            elementwise_unary(&v, f)
        }
        other => elementwise_unary(other, f),
    }
}

fn coerce_to_i64(v: &Value) -> Result<i64, ErrorKind> {
    let n = coerce_to_number(v)?;
    Ok(n.trunc() as i64)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RoundMode {
    Nearest,
    Down,
    Up,
}

fn round_with_mode(n: f64, digits: i32, mode: RoundMode) -> f64 {
    let factor = 10f64.powi(digits.saturating_abs());
    if !factor.is_finite() || factor == 0.0 {
        return n;
    }

    let scaled = if digits >= 0 { n * factor } else { n / factor };
    let rounded = match mode {
        RoundMode::Down => scaled.trunc(),
        RoundMode::Up => {
            if scaled.is_sign_negative() {
                scaled.trunc() - if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            } else {
                scaled.trunc() + if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            }
        }
        RoundMode::Nearest => {
            // Excel rounds halves away from zero.
            let frac = scaled.fract().abs();
            let base = scaled.trunc();
            if frac < 0.5 {
                base
            } else {
                base + scaled.signum()
            }
        }
    };

    if digits >= 0 {
        rounded / factor
    } else {
        rounded * factor
    }
}

fn fn_round_impl(args: &[Value], mode: RoundMode, grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let number = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(deref_value_dynamic(args[0].clone(), grid, base))
        }
        _ => None,
    };
    let digits = match &args[1] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(deref_value_dynamic(args[1].clone(), grid, base))
        }
        _ => None,
    };
    let number = number.as_ref().unwrap_or(&args[0]);
    let digits = digits.as_ref().unwrap_or(&args[1]);

    lift2(number, digits, |number, digits| {
        let number = match coerce_to_number(number) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let digits = match coerce_to_i64(digits) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        Value::Number(round_with_mode(number, digits as i32, mode))
    })
}

fn fn_round(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    fn_round_impl(args, RoundMode::Nearest, grid, base)
}

fn fn_roundup(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    fn_round_impl(args, RoundMode::Up, grid, base)
}

fn fn_rounddown(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    fn_round_impl(args, RoundMode::Down, grid, base)
}

fn fn_mod(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let n = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(deref_value_dynamic(args[0].clone(), grid, base))
        }
        _ => None,
    };
    let d = match &args[1] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(deref_value_dynamic(args[1].clone(), grid, base))
        }
        _ => None,
    };
    let n = n.as_ref().unwrap_or(&args[0]);
    let d = d.as_ref().unwrap_or(&args[1]);

    lift2(n, d, |n, d| {
        let n = match coerce_to_number(n) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let d = match coerce_to_number(d) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if d == 0.0 {
            return Value::Error(ErrorKind::Div0);
        }
        Value::Number(n - d * (n / d).floor())
    })
}

fn fn_sign(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    let f = |elem: &Value| {
        let number = match coerce_to_number(elem) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if !number.is_finite() {
            return Value::Error(ErrorKind::Num);
        }
        if number > 0.0 {
            Value::Number(1.0)
        } else if number < 0.0 {
            Value::Number(-1.0)
        } else {
            Value::Number(0.0)
        }
    };

    match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            let v = deref_value_dynamic(args[0].clone(), grid, base);
            elementwise_unary(&v, f)
        }
        other => elementwise_unary(other, f),
    }
}

fn fn_not(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| match coerce_to_bool(v) {
        Ok(b) => Value::Bool(!b),
        Err(e) => Value::Error(e),
    })
}

fn fn_if(args: &[Value]) -> Value {
    if args.len() < 2 || args.len() > 3 {
        return Value::Error(ErrorKind::Value);
    }
    let cond = match coerce_to_bool(&args[0]) {
        Ok(b) => b,
        Err(e) => return Value::Error(e),
    };
    if cond {
        args[1].clone()
    } else if args.len() >= 3 {
        args[2].clone()
    } else {
        // Engine behavior: missing false branch defaults to FALSE (not blank).
        Value::Bool(false)
    }
}

fn fn_ifs(args: &[Value]) -> Value {
    if args.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }
    if args.len() < 2 {
        return Value::Error(ErrorKind::Value);
    }
    for pair in args.chunks_exact(2) {
        let cond = match coerce_to_bool(&pair[0]) {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        };
        if cond {
            return pair[1].clone();
        }
    }
    Value::Error(ErrorKind::NA)
}

fn fn_switch(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 3 {
        return Value::Error(ErrorKind::Value);
    }
    let expr_val = args[0].clone();
    if let Value::Error(e) = expr_val {
        return Value::Error(e);
    }

    let has_default = (args.len() - 1) % 2 != 0;
    let pairs_end = if has_default {
        args.len() - 1
    } else {
        args.len()
    };
    let pairs = &args[1..pairs_end];
    let default = if has_default {
        Some(&args[args.len() - 1])
    } else {
        None
    };

    if pairs.len() < 2 || pairs.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    for pair in pairs.chunks_exact(2) {
        let left = deref_value_dynamic(expr_val.clone(), grid, base);
        let right = deref_value_dynamic(pair[0].clone(), grid, base);
        let matches = excel_compare(&left, &right, BinaryOp::Eq);
        match matches {
            Value::Bool(true) => return pair[1].clone(),
            Value::Bool(false) => continue,
            Value::Error(e) => return Value::Error(e),
            _ => return Value::Error(ErrorKind::Value),
        }
    }

    if let Some(default_val) = default {
        return default_val.clone();
    }
    Value::Error(ErrorKind::NA)
}

fn fn_choose(args: &[Value]) -> Value {
    // CHOOSE(index, value1, [value2], ...)
    if args.len() < 2 || args.len() > 255 {
        return Value::Error(ErrorKind::Value);
    }

    let index_value = &args[0];
    if let Value::Error(e) = index_value {
        return Value::Error(*e);
    }

    let idx = match coerce_to_i64(index_value) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if idx < 1 {
        return Value::Error(ErrorKind::Value);
    }

    let choice_idx = match usize::try_from(idx - 1) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Value),
    };
    let Some(v) = args.get(choice_idx + 1) else {
        return Value::Error(ErrorKind::Value);
    };
    v.clone()
}

fn fn_iferror(args: &[Value]) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(args[0], Value::Error(_)) {
        args[1].clone()
    } else {
        args[0].clone()
    }
}

fn fn_ifna(args: &[Value]) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(args[0], Value::Error(ErrorKind::NA)) {
        args[1].clone()
    } else {
        args[0].clone()
    }
}

fn fn_iserror(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    match &args[0] {
        // Multi-sheet spans are lowered as `Value::MultiRange` in bytecode. The AST evaluator
        // treats these as a reference-union argument which evaluates to `#VALUE!`, and then
        // ISERROR returns TRUE for that error value.
        Value::MultiRange(r) if r.areas.len() != 1 => Value::Bool(true),
        other => map_arg(other, grid, base, |v| {
            Value::Bool(matches!(v, Value::Error(_)))
        }),
    }
}

fn fn_isna(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(v, Value::Error(ErrorKind::NA)))
    })
}

fn fn_na(args: &[Value]) -> Value {
    if !args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    Value::Error(ErrorKind::NA)
}

fn fn_and(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut all_true = true;
    let mut any = false;

    for arg in args {
        let err = match arg {
            Value::Range(r) => and_range(grid, r.resolve(base), &mut all_true, &mut any),
            Value::MultiRange(r) => and_multi_range(grid, r, base, &mut all_true, &mut any),
            Value::Array(a) => and_array(a, &mut all_true, &mut any),
            other => and_scalar(other, &mut all_true, &mut any),
        };
        if let Some(e) = err {
            return Value::Error(e);
        }
    }

    if !any {
        Value::Bool(true)
    } else {
        Value::Bool(all_true)
    }
}

fn fn_or(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut any_true = false;
    let mut any = false;

    for arg in args {
        let err = match arg {
            Value::Range(r) => or_range(grid, r.resolve(base), &mut any_true, &mut any),
            Value::MultiRange(r) => or_multi_range(grid, r, base, &mut any_true, &mut any),
            Value::Array(a) => or_array(a, &mut any_true, &mut any),
            other => or_scalar(other, &mut any_true, &mut any),
        };
        if let Some(e) = err {
            return Value::Error(e);
        }
    }

    if !any {
        Value::Bool(false)
    } else {
        Value::Bool(any_true)
    }
}

fn fn_xor(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut acc = false;

    for arg in args {
        let err = match arg {
            Value::Range(r) => xor_range(grid, r.resolve(base), &mut acc),
            Value::MultiRange(r) => xor_multi_range(grid, r, base, &mut acc),
            Value::Array(a) => xor_array(a, &mut acc),
            other => xor_scalar(other, &mut acc),
        };
        if let Some(e) = err {
            return Value::Error(e);
        }
    }

    Value::Bool(acc)
}

fn and_scalar(v: &Value, all_true: &mut bool, any: &mut bool) -> Option<ErrorKind> {
    match v {
        Value::Error(e) => Some(*e),
        Value::Number(n) => {
            *any = true;
            if *n == 0.0 {
                *all_true = false;
            }
            None
        }
        Value::Bool(b) => {
            *any = true;
            if !*b {
                *all_true = false;
            }
            None
        }
        Value::Empty | Value::Missing => None,
        Value::Text(_) | Value::Entity(_) | Value::Record(_) | Value::Lambda(_) => {
            Some(ErrorKind::Value)
        }
        // Ranges/arrays are handled by the caller.
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => Some(ErrorKind::Spill),
    }
}

fn xor_scalar(v: &Value, acc: &mut bool) -> Option<ErrorKind> {
    match v {
        Value::Error(e) => Some(*e),
        Value::Number(n) => {
            *acc ^= *n != 0.0;
            None
        }
        Value::Bool(b) => {
            *acc ^= *b;
            None
        }
        Value::Empty | Value::Missing => None,
        // Scalar text arguments coerce like NOT().
        Value::Text(_) | Value::Entity(_) | Value::Record(_) => match coerce_to_bool(v) {
            Ok(b) => {
                *acc ^= b;
                None
            }
            Err(e) => Some(e),
        },
        Value::Lambda(_) => Some(ErrorKind::Value),
        // Ranges/arrays are handled by the caller.
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => Some(ErrorKind::Spill),
    }
}

fn or_scalar(v: &Value, any_true: &mut bool, any: &mut bool) -> Option<ErrorKind> {
    match v {
        Value::Error(e) => Some(*e),
        Value::Number(n) => {
            *any = true;
            if *n != 0.0 {
                *any_true = true;
            }
            None
        }
        Value::Bool(b) => {
            *any = true;
            if *b {
                *any_true = true;
            }
            None
        }
        Value::Empty | Value::Missing => None,
        Value::Text(_) | Value::Entity(_) | Value::Record(_) | Value::Lambda(_) => {
            Some(ErrorKind::Value)
        }
        // Ranges/arrays are handled by the caller.
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => Some(ErrorKind::Spill),
    }
}

fn and_array(a: &ArrayValue, all_true: &mut bool, any: &mut bool) -> Option<ErrorKind> {
    for v in a.iter() {
        match v {
            Value::Error(e) => return Some(*e),
            Value::Number(n) => {
                *any = true;
                if *n == 0.0 {
                    *all_true = false;
                }
            }
            Value::Bool(b) => {
                *any = true;
                if !*b {
                    *all_true = false;
                }
            }
            // Text and blanks in arrays are ignored (same as references).
            Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Lambda(_)
            | Value::Empty
            | Value::Missing => {}
            // Arrays should be scalar values; ignore any nested arrays/references rather than
            // treating them as implicit spills.
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {}
        }
    }
    None
}

fn or_array(a: &ArrayValue, any_true: &mut bool, any: &mut bool) -> Option<ErrorKind> {
    for v in a.iter() {
        match v {
            Value::Error(e) => return Some(*e),
            Value::Number(n) => {
                *any = true;
                if *n != 0.0 {
                    *any_true = true;
                }
            }
            Value::Bool(b) => {
                *any = true;
                if *b {
                    *any_true = true;
                }
            }
            // Text and blanks in arrays are ignored (same as references).
            Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Lambda(_)
            | Value::Empty
            | Value::Missing => {}
            // Arrays should be scalar values; ignore any nested arrays/references rather than
            // treating them as implicit spills.
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {}
        }
    }
    None
}

fn xor_array(a: &ArrayValue, acc: &mut bool) -> Option<ErrorKind> {
    for v in a.iter() {
        match v {
            Value::Error(e) => return Some(*e),
            // Historical behavior: the engine used NaN as a blank sentinel for dense numeric arrays.
            // Keep ignoring NaNs so older array materializations still behave like blanks.
            Value::Number(n) if n.is_nan() => {}
            Value::Number(n) => *acc ^= *n != 0.0,
            Value::Bool(b) => *acc ^= *b,
            // Text and blanks in arrays are ignored (same as references).
            Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Lambda(_)
            | Value::Empty
            | Value::Missing => {}
            // Arrays should be scalar values; ignore any nested arrays/references rather than
            // treating them as implicit spills.
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {}
        }
    }
    None
}

fn and_range(
    grid: &dyn Grid,
    range: ResolvedRange,
    all_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Some(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => {
                        *any = true;
                        if n == 0.0 {
                            *all_true = false;
                        }
                    }
                    Value::Bool(b) => {
                        *any = true;
                        if !b {
                            *all_true = false;
                        }
                    }
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Some(err);
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value(CellCoord { row, col }) {
                Value::Error(e) => return Some(e),
                Value::Number(n) => {
                    *any = true;
                    if n == 0.0 {
                        *all_true = false;
                    }
                }
                Value::Bool(b) => {
                    *any = true;
                    if !b {
                        *all_true = false;
                    }
                }
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    None
}

fn or_range(
    grid: &dyn Grid,
    range: ResolvedRange,
    any_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Some(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => {
                        *any = true;
                        if n != 0.0 {
                            *any_true = true;
                        }
                    }
                    Value::Bool(b) => {
                        *any = true;
                        if b {
                            *any_true = true;
                        }
                    }
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Some(err);
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value(CellCoord { row, col }) {
                Value::Error(e) => return Some(e),
                Value::Number(n) => {
                    *any = true;
                    if n != 0.0 {
                        *any_true = true;
                    }
                }
                Value::Bool(b) => {
                    *any = true;
                    if b {
                        *any_true = true;
                    }
                }
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Lambda(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_) => {}
            }
        }
    }
    None
}

fn and_multi_range(
    grid: &dyn Grid,
    range: &super::value::MultiRangeRef,
    base: CellCoord,
    all_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    let mut areas: Vec<(SheetRangeRef, ResolvedRange)> = range
        .areas
        .iter()
        .cloned()
        .map(|area| {
            let resolved = area.range.resolve(base);
            (area, resolved)
        })
        .collect();
    areas.sort_by(|(a_area, a_range), (b_area, b_range)| {
        cmp_sheet_ids_in_tab_order(grid, &a_area.sheet, &b_area.sheet)
            .then_with(|| a_range.row_start.cmp(&b_range.row_start))
            .then_with(|| a_range.col_start.cmp(&b_range.col_start))
            .then_with(|| a_range.row_end.cmp(&b_range.row_end))
            .then_with(|| a_range.col_end.cmp(&b_range.col_end))
    });

    for (area, resolved) in areas {
        if let Some(e) = and_range_on_sheet(grid, &area.sheet, resolved, all_true, any) {
            return Some(e);
        }
    }
    None
}

fn and_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
    all_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Some(ErrorKind::Ref);
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => {
                        *any = true;
                        if n == 0.0 {
                            *all_true = false;
                        }
                    }
                    Value::Bool(b) => {
                        *any = true;
                        if !b {
                            *all_true = false;
                        }
                    }
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Some(err);
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Error(e) => return Some(e),
                Value::Number(n) => {
                    *any = true;
                    if n == 0.0 {
                        *all_true = false;
                    }
                }
                Value::Bool(b) => {
                    *any = true;
                    if !b {
                        *all_true = false;
                    }
                }
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Lambda(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_) => {}
            }
        }
    }
    None
}

fn or_multi_range(
    grid: &dyn Grid,
    range: &super::value::MultiRangeRef,
    base: CellCoord,
    any_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    let mut areas: Vec<(SheetRangeRef, ResolvedRange)> = range
        .areas
        .iter()
        .cloned()
        .map(|area| {
            let resolved = area.range.resolve(base);
            (area, resolved)
        })
        .collect();
    areas.sort_by(|(a_area, a_range), (b_area, b_range)| {
        cmp_sheet_ids_in_tab_order(grid, &a_area.sheet, &b_area.sheet)
            .then_with(|| a_range.row_start.cmp(&b_range.row_start))
            .then_with(|| a_range.col_start.cmp(&b_range.col_start))
            .then_with(|| a_range.row_end.cmp(&b_range.row_end))
            .then_with(|| a_range.col_end.cmp(&b_range.col_end))
    });

    for (area, resolved) in areas {
        if let Some(e) = or_range_on_sheet(grid, &area.sheet, resolved, any_true, any) {
            return Some(e);
        }
    }
    None
}

fn or_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
    any_true: &mut bool,
    any: &mut bool,
) -> Option<ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Some(ErrorKind::Ref);
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => {
                        *any = true;
                        if n != 0.0 {
                            *any_true = true;
                        }
                    }
                    Value::Bool(b) => {
                        *any = true;
                        if b {
                            *any_true = true;
                        }
                    }
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Some(err);
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Error(e) => return Some(e),
                Value::Number(n) => {
                    *any = true;
                    if n != 0.0 {
                        *any_true = true;
                    }
                }
                Value::Bool(b) => {
                    *any = true;
                    if b {
                        *any_true = true;
                    }
                }
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Lambda(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_) => {}
            }
        }
    }
    None
}

fn map_value<F>(value: &Value, f: F) -> Value
where
    F: Fn(&Value) -> Value + Copy,
{
    match value {
        Value::Array(arr) => {
            let total = match arr.rows.checked_mul(arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for v in arr.iter() {
                out.push(f(v));
            }
            Value::Array(ArrayValue::new(arr.rows, arr.cols, out))
        }
        other => f(other),
    }
}

fn map_arg<F>(arg: &Value, grid: &dyn Grid, base: CellCoord, f: F) -> Value
where
    F: Fn(&Value) -> Value + Copy,
{
    match arg {
        Value::Range(r) => {
            let resolved = r.resolve(base);
            if !range_in_bounds(grid, resolved) {
                return map_value(&Value::Error(ErrorKind::Ref), f);
            }

            // This function materializes/iterates the range, so record the referenced rectangle
            // once for dynamic dependency tracing (avoid per-cell events).
            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: resolved.row_start,
                    col: resolved.col_start,
                },
                CellCoord {
                    row: resolved.row_end,
                    col: resolved.col_end,
                },
            );

            if resolved.rows() == 1 && resolved.cols() == 1 {
                let v = grid.get_value(CellCoord {
                    row: resolved.row_start,
                    col: resolved.col_start,
                });
                return map_value(&v, f);
            }

            let rows = match usize::try_from(resolved.rows()) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Num),
            };
            let cols = match usize::try_from(resolved.cols()) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Num),
            };
            let total = match rows.checked_mul(cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for row in resolved.row_start..=resolved.row_end {
                for col in resolved.col_start..=resolved.col_end {
                    let v = grid.get_value(CellCoord { row, col });
                    values.push(f(&v));
                }
            }
            Value::Array(ArrayValue::new(rows, cols, values))
        }
        Value::MultiRange(r) => {
            // Multi-range values generally represent 3D sheet spans. For single-sheet spans,
            // treat this like a normal range reference; for multi-sheet spans, we cannot spill
            // a single rectangular array, so surface `#VALUE!` (matching the AST evaluator's
            // reference-union behavior).
            let [area] = r.areas.as_ref() else {
                return Value::Error(ErrorKind::Value);
            };

            let resolved = area.range.resolve(base);
            if !range_in_bounds_on_sheet(grid, &area.sheet, resolved) {
                return map_value(&Value::Error(ErrorKind::Ref), f);
            }

            // Record the referenced rectangle once for dynamic dependency tracing.
            grid.record_reference_on_sheet(
                &area.sheet,
                CellCoord {
                    row: resolved.row_start,
                    col: resolved.col_start,
                },
                CellCoord {
                    row: resolved.row_end,
                    col: resolved.col_end,
                },
            );

            if resolved.rows() == 1 && resolved.cols() == 1 {
                let v = grid.get_value_on_sheet(
                    &area.sheet,
                    CellCoord {
                        row: resolved.row_start,
                        col: resolved.col_start,
                    },
                );
                return map_value(&v, f);
            }

            let rows = match usize::try_from(resolved.rows()) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Num),
            };
            let cols = match usize::try_from(resolved.cols()) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Num),
            };
            let total = match rows.checked_mul(cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for row in resolved.row_start..=resolved.row_end {
                for col in resolved.col_start..=resolved.col_end {
                    let v = grid.get_value_on_sheet(&area.sheet, CellCoord { row, col });
                    values.push(f(&v));
                }
            }

            Value::Array(ArrayValue::new(rows, cols, values))
        }
        other => map_value(other, f),
    }
}

fn fn_isblank(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(v, Value::Empty | Value::Missing))
    })
}

fn fn_isnumber(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(v, Value::Number(n) if n.is_finite()))
    })
}

fn fn_istext(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(
            v,
            Value::Text(_) | Value::Entity(_) | Value::Record(_)
        ))
    })
}

fn fn_islogical(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(v, Value::Bool(_)))
    })
}

fn fn_iserr(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| {
        Value::Bool(matches!(v, Value::Error(e) if *e != ErrorKind::NA))
    })
}

fn type_code_for_scalar(value: &Value) -> i32 {
    match value {
        Value::Number(_) | Value::Empty | Value::Missing => 1,
        Value::Text(_) | Value::Entity(_) | Value::Record(_) => 2,
        Value::Bool(_) => 4,
        Value::Error(_) | Value::Lambda(_) => 16,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => 64,
    }
}

fn fn_type(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }

    let code = match &args[0] {
        Value::Range(r) => {
            let resolved = r.resolve(base);
            if resolved.rows() == 1 && resolved.cols() == 1 {
                let coord = CellCoord {
                    row: resolved.row_start,
                    col: resolved.col_start,
                };
                if grid.in_bounds(coord) {
                    grid.record_reference(grid.sheet_id(), coord, coord);
                }
                let v = grid.get_value(coord);
                type_code_for_scalar(&v)
            } else {
                64
            }
        }
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => return Value::Error(ErrorKind::Ref),
            [only] => {
                let resolved = only.range.resolve(base);
                if resolved.rows() == 1 && resolved.cols() == 1 {
                    let coord = CellCoord {
                        row: resolved.row_start,
                        col: resolved.col_start,
                    };
                    if grid.in_bounds_on_sheet(&only.sheet, coord) {
                        grid.record_reference_on_sheet(&only.sheet, coord, coord);
                    }
                    let v = grid.get_value_on_sheet(&only.sheet, coord);
                    type_code_for_scalar(&v)
                } else {
                    64
                }
            }
            // Discontiguous unions (e.g. 3D sheet spans) are not valid TYPE arguments.
            _ => return Value::Error(ErrorKind::Value),
        },
        other => type_code_for_scalar(other),
    };

    Value::Number(code as f64)
}

fn error_type_code(kind: ErrorKind) -> i32 {
    // Keep bytecode ERROR.TYPE semantics aligned with the AST evaluator's mapping.
    // `bytecode::value::ErrorKind` mirrors `crate::value::ErrorKind`, so use the canonical
    // `code()` mapping from the main error kind.
    i32::from(EngineErrorKind::from(kind).code())
}

fn fn_error_type(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| match v {
        Value::Error(e) => Value::Number(error_type_code(*e) as f64),
        _ => Value::Error(ErrorKind::NA),
    })
}

fn fn_n(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    map_arg(&args[0], grid, base, |v| match v {
        Value::Error(e) => Value::Error(*e),
        Value::Number(n) => Value::Number(*n),
        Value::Bool(b) => Value::Number(if *b { 1.0 } else { 0.0 }),
        Value::Empty | Value::Missing | Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
            Value::Number(0.0)
        }
        Value::Lambda(_) | Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {
            Value::Error(ErrorKind::Value)
        }
    })
}

fn fn_t(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }

    fn record_display_text_like(record: &RecordValue) -> String {
        fn text_like_value(value: &EngineValue) -> Option<String> {
            match value {
                EngineValue::Text(s) => Some(s.clone()),
                EngineValue::Entity(entity) => Some(entity.display.clone()),
                EngineValue::Record(record) => Some(record_display_text_like(record)),
                _ => None,
            }
        }

        if let Some(display_field) = record.display_field.as_deref() {
            if let Some(value) = record.get_field_case_insensitive(display_field) {
                if let Some(text) = text_like_value(&value) {
                    return text;
                }
            }
        }
        record.display.clone()
    }

    map_arg(&args[0], grid, base, |v| match v {
        Value::Error(e) => Value::Error(*e),
        Value::Text(s) => Value::Text(s.clone()),
        Value::Entity(ent) => Value::Text(Arc::from(ent.display.as_str())),
        Value::Record(rec) => {
            Value::Text(Arc::from(record_display_text_like(rec.as_ref()).as_str()))
        }
        Value::Number(_) | Value::Bool(_) | Value::Empty | Value::Missing | Value::Lambda(_) => {
            Value::Text(Arc::from(""))
        }
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => Value::Text(Arc::from("")),
    })
}

fn fn_row(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    fn eval_row_for_range(grid: &dyn Grid, sheet: &SheetId, reference: ResolvedRange) -> Value {
        if !range_in_bounds_on_sheet(grid, sheet, reference) {
            return Value::Error(ErrorKind::Ref);
        }

        if reference.rows() == 1 && reference.cols() == 1 {
            return Value::Number((reference.row_start + 1) as f64);
        }

        let rows = match usize::try_from(reference.rows()) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Num),
        };
        let cols = match usize::try_from(reference.cols()) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Num),
        };

        let (sheet_rows, sheet_cols) = grid.bounds_on_sheet(sheet);
        let spans_all_cols =
            reference.col_start == 0 && reference.col_end == sheet_cols.saturating_sub(1);
        let spans_all_rows =
            reference.row_start == 0 && reference.row_end == sheet_rows.saturating_sub(1);

        if spans_all_cols || spans_all_rows {
            if rows > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(rows).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for row in reference.row_start..=reference.row_end {
                values.push(Value::Number((row + 1) as f64));
            }
            if rows == 1 {
                return values.first().cloned().unwrap_or(Value::Empty);
            }
            return Value::Array(ArrayValue::new(rows, 1, values));
        }

        let total = match rows.checked_mul(cols) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Spill),
        };
        if total > MAX_MATERIALIZED_ARRAY_CELLS {
            return Value::Error(ErrorKind::Spill);
        }
        let mut values = Vec::new();
        if values.try_reserve_exact(total).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        for row in reference.row_start..=reference.row_end {
            let n = Value::Number((row + 1) as f64);
            for _ in reference.col_start..=reference.col_end {
                values.push(n.clone());
            }
        }
        Value::Array(ArrayValue::new(rows, cols, values))
    }

    match args {
        [] => Value::Number((base.row + 1) as f64),
        [Value::Range(r)] => {
            eval_row_for_range(grid, &SheetId::Local(grid.sheet_id()), r.resolve(base))
        }
        [Value::MultiRange(r)] => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => eval_row_for_range(grid, &only.sheet, only.range.resolve(base)),
            _ => Value::Error(ErrorKind::Value),
        },
        [Value::Error(e)] => Value::Error(*e),
        [_] => Value::Error(ErrorKind::Value),
        _ => Value::Error(ErrorKind::Value),
    }
}

fn fn_column(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    fn eval_column_for_range(grid: &dyn Grid, sheet: &SheetId, reference: ResolvedRange) -> Value {
        if !range_in_bounds_on_sheet(grid, sheet, reference) {
            return Value::Error(ErrorKind::Ref);
        }

        if reference.rows() == 1 && reference.cols() == 1 {
            return Value::Number((reference.col_start + 1) as f64);
        }

        let rows = match usize::try_from(reference.rows()) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Num),
        };
        let cols = match usize::try_from(reference.cols()) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Num),
        };

        let (sheet_rows, sheet_cols) = grid.bounds_on_sheet(sheet);
        let spans_all_cols =
            reference.col_start == 0 && reference.col_end == sheet_cols.saturating_sub(1);
        let spans_all_rows =
            reference.row_start == 0 && reference.row_end == sheet_rows.saturating_sub(1);

        if spans_all_cols || spans_all_rows {
            if cols > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(cols).is_err() {
                return Value::Error(ErrorKind::Num);
            }
            for col in reference.col_start..=reference.col_end {
                values.push(Value::Number((col + 1) as f64));
            }
            if cols == 1 {
                return values.first().cloned().unwrap_or(Value::Empty);
            }
            return Value::Array(ArrayValue::new(1, cols, values));
        }

        let total = match rows.checked_mul(cols) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Spill),
        };
        if total > MAX_MATERIALIZED_ARRAY_CELLS {
            return Value::Error(ErrorKind::Spill);
        }
        let mut values = Vec::new();
        if values.try_reserve_exact(total).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        let mut row_values = Vec::new();
        if row_values.try_reserve_exact(cols).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        for col in reference.col_start..=reference.col_end {
            row_values.push(Value::Number((col + 1) as f64));
        }
        for _ in 0..rows {
            values.extend(row_values.iter().cloned());
        }

        Value::Array(ArrayValue::new(rows, cols, values))
    }

    match args {
        [] => Value::Number((base.col + 1) as f64),
        [Value::Range(r)] => {
            eval_column_for_range(grid, &SheetId::Local(grid.sheet_id()), r.resolve(base))
        }
        [Value::MultiRange(r)] => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => eval_column_for_range(grid, &only.sheet, only.range.resolve(base)),
            _ => Value::Error(ErrorKind::Value),
        },
        [Value::Error(e)] => Value::Error(*e),
        [_] => Value::Error(ErrorKind::Value),
        _ => Value::Error(ErrorKind::Value),
    }
}

fn fn_rows(args: &[Value], base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    match &args[0] {
        Value::Range(r) => Value::Number(r.resolve(base).rows() as f64),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => Value::Number(only.range.resolve(base).rows() as f64),
            _ => Value::Error(ErrorKind::Value),
        },
        Value::Array(a) => Value::Number(a.rows as f64),
        Value::Error(e) => Value::Error(*e),
        _ => Value::Error(ErrorKind::Value),
    }
}

fn fn_columns(args: &[Value], base: CellCoord) -> Value {
    if args.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }
    match &args[0] {
        Value::Range(r) => Value::Number(r.resolve(base).cols() as f64),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => Value::Number(only.range.resolve(base).cols() as f64),
            _ => Value::Error(ErrorKind::Value),
        },
        Value::Array(a) => Value::Number(a.cols as f64),
        Value::Error(e) => Value::Error(*e),
        _ => Value::Error(ErrorKind::Value),
    }
}

fn fn_address(args: &[Value], grid: &dyn Grid, _base: CellCoord) -> Value {
    if !(2..=5).contains(&args.len()) {
        return Value::Error(ErrorKind::Value);
    }

    let row_num = match coerce_to_i64(&args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let col_num = match coerce_to_i64(&args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (sheet_rows, sheet_cols) = grid.bounds();
    let sheet_rows = i64::from(sheet_rows);
    let sheet_cols = i64::from(sheet_cols);

    if row_num < 1 || row_num > sheet_rows {
        return Value::Error(ErrorKind::Value);
    }
    if col_num < 1 || col_num > sheet_cols {
        return Value::Error(ErrorKind::Value);
    }

    let abs_num = if args.len() >= 3 && !matches!(args[2], Value::Missing) {
        match coerce_to_i64(&args[2]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    let (col_abs, row_abs) = match abs_num {
        1 => (true, true),
        2 => (false, true),
        3 => (true, false),
        4 => (false, false),
        _ => return Value::Error(ErrorKind::Value),
    };

    let a1 = if args.len() >= 4 && !matches!(args[3], Value::Missing) {
        match coerce_to_bool(&args[3]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    let sheet_prefix = if args.len() >= 5 && !matches!(args[4], Value::Missing) {
        match coerce_to_string(&args[4]) {
            Ok(raw) => {
                if raw.is_empty() {
                    None
                } else {
                    Some(format!("{}!", quote_sheet_name(&raw)))
                }
            }
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    let addr = if a1 {
        format_a1_address(row_num as u32, col_num as u32, row_abs, col_abs)
    } else {
        format_r1c1_address(row_num, col_num, row_abs, col_abs)
    };

    if let Some(prefix) = sheet_prefix {
        Value::Text(Arc::from(format!("{prefix}{addr}")))
    } else {
        Value::Text(Arc::from(addr))
    }
}

fn fn_offset(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 3 || args.len() > 5 {
        return Value::Error(ErrorKind::Value);
    }

    // Coerce scalar numeric arguments, matching the AST evaluator's `eval_scalar_arg` behavior:
    // ranges are implicitly intersected, while array/lambda values are rejected.
    fn scalar_i64_arg(v: &Value, grid: &dyn Grid, base: CellCoord) -> Result<i64, ErrorKind> {
        let v = match v {
            Value::Range(_) | Value::MultiRange(_) => {
                apply_implicit_intersection(v.clone(), grid, base)
            }
            _ => v.clone(),
        };
        match v {
            Value::Error(e) => Err(e),
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_) => {
                Err(ErrorKind::Value)
            }
            other => coerce_to_i64(&other),
        }
    }

    let current_sheet = thread_current_sheet_id() as usize;

    let (sheet, base_range) = match &args[0] {
        Value::Range(r) => (SheetId::Local(current_sheet), r.resolve(base)),
        Value::MultiRange(r) => match r.areas.as_ref() {
            [] => return Value::Error(ErrorKind::Ref),
            [only] => (only.sheet.clone(), only.range.resolve(base)),
            _ => return Value::Error(ErrorKind::Value),
        },
        Value::Error(e) => return Value::Error(*e),
        _ => return Value::Error(ErrorKind::Value),
    };

    if !range_in_bounds_on_sheet(grid, &sheet, base_range) {
        return Value::Error(ErrorKind::Ref);
    }

    let default_height = i64::from(base_range.rows());
    let default_width = i64::from(base_range.cols());

    let rows = match scalar_i64_arg(&args[1], grid, base) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = match scalar_i64_arg(&args[2], grid, base) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let height = if args.len() >= 4 && !matches!(args[3], Value::Missing) {
        match scalar_i64_arg(&args[3], grid, base) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        default_height
    };
    let width = if args.len() >= 5 && !matches!(args[4], Value::Missing) {
        match scalar_i64_arg(&args[4], grid, base) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        default_width
    };

    if height < 1 || width < 1 {
        return Value::Error(ErrorKind::Ref);
    }

    let start_row = i64::from(base_range.row_start).saturating_add(rows);
    let start_col = i64::from(base_range.col_start).saturating_add(cols);
    let end_row = match start_row.checked_add(height.saturating_sub(1)) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Ref),
    };
    let end_col = match start_col.checked_add(width.saturating_sub(1)) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Ref),
    };

    let (sheet_rows, sheet_cols) = grid.bounds_on_sheet(&sheet);
    let sheet_rows = i64::from(sheet_rows);
    let sheet_cols = i64::from(sheet_cols);
    if sheet_rows <= 0 || sheet_cols <= 0 {
        return Value::Error(ErrorKind::Ref);
    }
    if start_row < 0
        || start_col < 0
        || end_row < 0
        || end_col < 0
        || start_row >= sheet_rows
        || end_row >= sheet_rows
        || start_col >= sheet_cols
        || end_col >= sheet_cols
    {
        return Value::Error(ErrorKind::Ref);
    }

    let (Ok(start_row), Ok(start_col), Ok(end_row), Ok(end_col)) = (
        i32::try_from(start_row),
        i32::try_from(start_col),
        i32::try_from(end_row),
        i32::try_from(end_col),
    ) else {
        return Value::Error(ErrorKind::Ref);
    };

    let range = RangeRef::new(
        Ref::new(start_row, start_col, true, true),
        Ref::new(end_row, end_col, true, true),
    );
    match sheet {
        SheetId::Local(sheet_id) if sheet_id == current_sheet => Value::Range(range),
        other_sheet => Value::MultiRange(MultiRangeRef::new(
            vec![SheetRangeRef::new(other_sheet, range)].into(),
        )),
    }
}

fn fn_indirect(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() || args.len() > 2 {
        return Value::Error(ErrorKind::Value);
    }

    let text_val = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            apply_implicit_intersection(args[0].clone(), grid, base)
        }
        _ => args[0].clone(),
    };
    if let Value::Error(e) = text_val {
        return Value::Error(e);
    }
    if matches!(
        text_val,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_)
    ) {
        return Value::Error(ErrorKind::Value);
    }
    let text = match coerce_to_string(&text_val) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let a1 = if args.len() >= 2 && !matches!(args[1], Value::Missing) {
        let a1_val = match &args[1] {
            Value::Range(_) | Value::MultiRange(_) => {
                apply_implicit_intersection(args[1].clone(), grid, base)
            }
            _ => args[1].clone(),
        };
        if let Value::Error(e) = a1_val {
            return Value::Error(e);
        }
        if matches!(
            a1_val,
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_)
        ) {
            return Value::Error(ErrorKind::Value);
        }
        match coerce_to_bool(&a1_val) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    let ref_text = text.trim();
    if ref_text.is_empty() {
        return Value::Error(ErrorKind::Ref);
    }

    let parsed = match crate::parse_formula(
        ref_text,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: if a1 {
                crate::ReferenceStyle::A1
            } else {
                crate::ReferenceStyle::R1C1
            },
            normalize_relative_to: None,
        },
    ) {
        Ok(ast) => ast,
        Err(_) => return Value::Error(ErrorKind::Ref),
    };

    let Ok(origin_row) = u32::try_from(base.row) else {
        return Value::Error(ErrorKind::Ref);
    };
    let Ok(origin_col) = u32::try_from(base.col) else {
        return Value::Error(ErrorKind::Ref);
    };
    let origin_ast = crate::CellAddr::new(origin_row, origin_col);
    let origin_eval = crate::eval::CellAddr {
        row: origin_row,
        col: origin_col,
    };
    let lowered = crate::eval::lower_ast(&parsed, if a1 { None } else { Some(origin_ast) });

    let current_sheet = thread_current_sheet_id() as usize;
    fn resolve_sheet(
        grid: &dyn Grid,
        sheet: &crate::eval::SheetReference<String>,
    ) -> Option<SheetId> {
        match sheet {
            crate::eval::SheetReference::Current => {
                Some(SheetId::Local(thread_current_sheet_id() as usize))
            }
            crate::eval::SheetReference::Sheet(name) => {
                // Resolve local sheets only. Avoid interpreting bracketed external sheet keys like
                // `"[Book.xlsx]Sheet1"` as local sheet names. External workbook refs are
                // represented separately by the parser/lowerer (`SheetReference::External`), but
                // some grids may still attempt to resolve them via `resolve_sheet_name`. Reject
                // them here to avoid accidentally treating external keys as local sheets.
                if name.starts_with('[') {
                    return None;
                }
                Some(SheetId::Local(grid.resolve_sheet_name(name)?))
            }
            crate::eval::SheetReference::SheetRange(start, end) => {
                if start.starts_with('[') || end.starts_with('[') {
                    return None;
                }
                let start_id = grid.resolve_sheet_name(start)?;
                let end_id = grid.resolve_sheet_name(end)?;
                (start_id == end_id).then_some(SheetId::Local(start_id))
            }
            crate::eval::SheetReference::External(key) => {
                // Match `functions::builtins_reference::INDIRECT`: allow single-sheet external
                // workbook references (e.g. `"[Book.xlsx]Sheet1"`), but reject external 3D spans
                // like `"[Book.xlsx]Sheet1:Sheet3"`.
                crate::eval::is_valid_external_sheet_key(key)
                    .then_some(SheetId::External(Arc::from(key.as_str())))
            }
        }
    }

    fn resolve_coord(n: u32, max: i32) -> Option<i32> {
        if n == crate::eval::CellAddr::SHEET_END {
            Some(max)
        } else {
            i32::try_from(n).ok()
        }
    }

    let make_range_value =
        |sheet: SheetId, start: crate::eval::CellAddr, end: crate::eval::CellAddr| -> Value {
            let (rows, cols) = grid.bounds_on_sheet(&sheet);
            if rows <= 0 || cols <= 0 {
                return Value::Error(ErrorKind::Ref);
            }
            let max_row = rows - 1;
            let max_col = cols - 1;

            let Some(start_row) = resolve_coord(start.row, max_row) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(start_col) = resolve_coord(start.col, max_col) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(end_row) = resolve_coord(end.row, max_row) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(end_col) = resolve_coord(end.col, max_col) else {
                return Value::Error(ErrorKind::Ref);
            };

            if start_row < 0 || start_col < 0 || end_row < 0 || end_col < 0 {
                return Value::Error(ErrorKind::Ref);
            }

            let range = RangeRef::new(
                Ref::new(start_row, start_col, true, true),
                Ref::new(end_row, end_col, true, true),
            );

            match sheet {
                SheetId::Local(sheet_id) if sheet_id == current_sheet => Value::Range(range),
                other_sheet => Value::MultiRange(MultiRangeRef::new(
                    vec![SheetRangeRef::new(other_sheet, range)].into(),
                )),
            }
        };

    match lowered {
        crate::eval::Expr::CellRef(r) => {
            let Some(sheet_id) = resolve_sheet(grid, &r.sheet) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(addr) = r.addr.resolve(origin_eval) else {
                return Value::Error(ErrorKind::Ref);
            };
            make_range_value(sheet_id, addr, addr)
        }
        crate::eval::Expr::RangeRef(r) => {
            let Some(sheet_id) = resolve_sheet(grid, &r.sheet) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(start) = r.start.resolve(origin_eval) else {
                return Value::Error(ErrorKind::Ref);
            };
            let Some(end) = r.end.resolve(origin_eval) else {
                return Value::Error(ErrorKind::Ref);
            };
            make_range_value(sheet_id, start, end)
        }
        crate::eval::Expr::Error(e) => Value::Error(ErrorKind::from(e)),
        crate::eval::Expr::NameRef(_) => Value::Error(ErrorKind::Ref),
        _ => Value::Error(ErrorKind::Ref),
    }
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(c, '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9')
}

fn starts_like_a1_cell_ref(s: &str) -> bool {
    // The lexer tokenizes A1-style cell references even when followed by additional identifier
    // characters (e.g. `A1B`), so treat any sheet name *starting* with a valid A1 ref as requiring
    // quotes. This matches `ast.rs` sheet-name formatting rules.
    let bytes = s.as_bytes();
    let mut i = 0;
    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }

    let start_letters = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start_letters {
        return false;
    }

    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }

    let start_digits = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == start_digits {
        return false;
    }

    crate::eval::parse_a1(&s[..i]).is_ok()
}

fn quote_sheet_name(name: &str) -> String {
    if name.is_empty() {
        return String::new();
    }

    let starts_like_number = matches!(name.chars().next(), Some('0'..='9' | '.'));
    let starts_like_r1c1 = matches!(name.chars().next(), Some('R' | 'r' | 'C' | 'c'))
        && matches!(name.chars().nth(1), Some('0'..='9' | '['));
    let starts_like_a1 = starts_like_a1_cell_ref(name);
    // The formula lexer treats TRUE/FALSE as booleans rather than identifiers; quoting is required
    // to disambiguate sheet names that match those keywords.
    let is_reserved = name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE");
    let needs_quote = starts_like_number
        || is_reserved
        || starts_like_r1c1
        || starts_like_a1
        || name.chars().any(|c| !is_ident_cont_char(c));

    if !needs_quote {
        return name.to_string();
    }

    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn col_to_name(col: u32) -> String {
    let mut n = col;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are always valid UTF-8")
}

fn format_a1_address(row_num: u32, col_num: u32, row_abs: bool, col_abs: bool) -> String {
    let mut out = String::new();
    if col_abs {
        out.push('$');
    }
    out.push_str(&col_to_name(col_num));
    if row_abs {
        out.push('$');
    }
    out.push_str(&row_num.to_string());
    out
}

fn format_r1c1_address(row_num: i64, col_num: i64, row_abs: bool, col_abs: bool) -> String {
    let mut out = String::new();
    if row_abs {
        out.push('R');
        out.push_str(&row_num.to_string());
    } else {
        out.push_str("R[");
        out.push_str(&row_num.to_string());
        out.push(']');
    }
    if col_abs {
        out.push('C');
        out.push_str(&col_num.to_string());
    } else {
        out.push_str("C[");
        out.push_str(&col_num.to_string());
        out.push(']');
    }
    out
}
fn xor_range(grid: &dyn Grid, range: ResolvedRange, acc: &mut bool) -> Option<ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Some(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => *acc ^= n != 0.0,
                    Value::Bool(b) => *acc ^= b,
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Some(err);
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value(CellCoord { row, col }) {
                Value::Error(e) => return Some(e),
                Value::Number(n) => *acc ^= n != 0.0,
                Value::Bool(b) => *acc ^= b,
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Lambda(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_) => {}
            }
        }
    }
    None
}

fn xor_multi_range(
    grid: &dyn Grid,
    range: &MultiRangeRef,
    base: CellCoord,
    acc: &mut bool,
) -> Option<ErrorKind> {
    // XOR is sensitive to overlap (duplicate cells must be visited once), so we use
    // `multirange_unique_areas` to subtract overlaps.
    //
    // However, the resulting disjoint rectangles are not guaranteed to preserve the AST
    // evaluator's row-major cell visitation order when iterated sequentially (e.g. a range with a
    // removed middle column yields left/right strips that must be interleaved row-by-row).
    //
    // XOR is commutative, so we can safely compute the accumulator in any order, but we must
    // preserve AST-like error precedence:
    // - Errors in earlier areas win.
    // - Within an area, the first error is the one with the smallest `(row, col)` coordinate.
    let mut current_area_idx: Option<usize> = None;
    let mut best_error_in_area: Option<(i32, i32, ErrorKind)> = None;

    for area in multirange_unique_areas(range, grid, base) {
        if current_area_idx != Some(area.area_idx) {
            if let Some((_, _, err)) = best_error_in_area {
                return Some(err);
            }
            current_area_idx = Some(area.area_idx);
            best_error_in_area = None;
        }

        if let Some((coord, err)) = xor_range_on_sheet(grid, &area.sheet, area.range, acc) {
            record_error_row_major(&mut best_error_in_area, coord, err);
        }
    }

    best_error_in_area.map(|(_, _, err)| err)
}

fn xor_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
    acc: &mut bool,
) -> Option<(CellCoord, ErrorKind)> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Some((
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            ErrorKind::Ref,
        ));
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Number(n) => *acc ^= n != 0.0,
                    Value::Bool(b) => *acc ^= b,
                    // Text/blanks in references are ignored.
                    Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Lambda(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_) => {}
                }
            }
            if let Some((row, col, err)) = best_error {
                return Some((CellCoord { row, col }, err));
            }
            return None;
        }
    }

    for row in range.row_start..=range.row_end {
        for col in range.col_start..=range.col_end {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Error(e) => return Some((CellCoord { row, col }, e)),
                Value::Number(n) => *acc ^= n != 0.0,
                Value::Bool(b) => *acc ^= b,
                // Text/blanks in references are ignored.
                Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    None
}

fn coerce_to_cow_str(v: &Value) -> Result<Cow<'_, str>, ErrorKind> {
    fn coerce_engine_value_to_string(v: &EngineValue) -> Result<String, ErrorKind> {
        match v {
            EngineValue::Text(s) => Ok(s.clone()),
            EngineValue::Entity(entity) => Ok(entity.display.clone()),
            EngineValue::Record(record) => {
                if let Some(display_field) = record.display_field.as_deref() {
                    if let Some(value) = record.get_field_case_insensitive(display_field) {
                        return coerce_engine_value_to_string(&value);
                    }
                }
                Ok(record.display.clone())
            }
            EngineValue::Number(n) => Ok(format_number_general_with_options(
                *n,
                thread_value_locale().separators,
                thread_date_system(),
            )),
            EngineValue::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
            EngineValue::Blank => Ok(String::new()),
            EngineValue::Error(e) => Err(ErrorKind::from(*e)),
            EngineValue::Reference(_)
            | EngineValue::ReferenceUnion(_)
            | EngineValue::Array(_)
            | EngineValue::Lambda(_)
            | EngineValue::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    match v {
        Value::Text(s) => Ok(Cow::Borrowed(s.as_ref())),
        Value::Entity(v) => Ok(Cow::Borrowed(v.display.as_str())),
        Value::Record(v) => {
            if let Some(display_field) = v.display_field.as_deref() {
                if let Some(value) = v.get_field_case_insensitive(display_field) {
                    return Ok(Cow::Owned(coerce_engine_value_to_string(&value)?));
                }
            }
            Ok(Cow::Borrowed(v.display.as_str()))
        }
        Value::Number(n) => Ok(Cow::Owned(format_number_general_with_options(
            *n,
            thread_value_locale().separators,
            thread_date_system(),
        ))),
        Value::Bool(b) => Ok(Cow::Borrowed(if *b { "TRUE" } else { "FALSE" })),
        Value::Empty | Value::Missing => Ok(Cow::Borrowed("")),
        Value::Error(e) => Err(*e),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => Err(ErrorKind::Value),
    }
}

fn coerce_to_string(v: &Value) -> Result<String, ErrorKind> {
    Ok(coerce_to_cow_str(v)?.into_owned())
}

fn concat_binary(left: &Value, right: &Value) -> Value {
    // Elementwise concatenation: propagate errors per-element.
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let left_str = match coerce_to_cow_str(left) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let right_str = match coerce_to_cow_str(right) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let mut out = String::with_capacity(
        left_str
            .as_ref()
            .len()
            .saturating_add(right_str.as_ref().len()),
    );
    out.push_str(left_str.as_ref());
    out.push_str(right_str.as_ref());
    Value::Text(out.into())
}

fn fn_concat_scalar(args: &[Value]) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out = String::new();
    for arg in args {
        match coerce_to_cow_str(arg) {
            Ok(s) => out.push_str(s.as_ref()),
            Err(e) => return Value::Error(e),
        }
    }
    Value::Text(out.into())
}

fn fn_concat_op(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 2 {
        return Value::Error(ErrorKind::Value);
    }

    // Common fast path: scalar/array concatenation without any reference values. Avoid cloning
    // each argument into a temporary vector only to discover that no dereferencing is needed.
    if !args
        .iter()
        .any(|arg| matches!(arg, Value::Range(_) | Value::MultiRange(_)))
    {
        if !args.iter().any(|arg| matches!(arg, Value::Array(_))) {
            return fn_concat_scalar(args);
        }

        let mut acc = args[0].clone();
        for next in args.iter().skip(1) {
            acc = elementwise_binary(&acc, next, concat_binary);
            if matches!(acc, Value::Error(_)) {
                break;
            }
        }
        return acc;
    }

    // Dereference any range arguments so `A1:A2&"c"` produces a spilled array.
    // Keep single-cell ranges scalar (deref returns the cell value directly).
    let mut deref_args = Vec::with_capacity(args.len());
    let mut saw_array = false;
    for arg in args {
        let v = deref_value_dynamic(arg.clone(), grid, base);
        saw_array |= matches!(v, Value::Array(_));
        deref_args.push(v);
    }

    if !saw_array {
        return fn_concat_scalar(&deref_args);
    }

    let mut acc = deref_args
        .first()
        .cloned()
        .unwrap_or(Value::Error(ErrorKind::Value));
    for next in deref_args.iter().skip(1) {
        acc = elementwise_binary(&acc, next, concat_binary);
        if matches!(acc, Value::Error(_)) {
            break;
        }
    }
    acc
}

fn fn_concat(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let mut out = String::new();

    for arg in args {
        match arg {
            Value::Array(arr) => {
                for v in arr.iter() {
                    let s = match coerce_to_cow_str(v) {
                        Ok(s) => s,
                        Err(e) => return Value::Error(e),
                    };
                    out.push_str(s.as_ref());
                }
            }
            Value::Range(r) => {
                let range = r.resolve(base);
                if !range_in_bounds(grid, range) {
                    return Value::Error(ErrorKind::Ref);
                }
                grid.record_reference(
                    grid.sheet_id(),
                    CellCoord {
                        row: range.row_start,
                        col: range.col_start,
                    },
                    CellCoord {
                        row: range.row_end,
                        col: range.col_end,
                    },
                );
                for row in range.row_start..=range.row_end {
                    for col in range.col_start..=range.col_end {
                        let v = grid.get_value(CellCoord { row, col });
                        let s = match coerce_to_cow_str(&v) {
                            Ok(s) => s,
                            Err(e) => return Value::Error(e),
                        };
                        out.push_str(s.as_ref());
                    }
                }
            }
            Value::MultiRange(r) => {
                if r.is_empty() {
                    return Value::Error(ErrorKind::Ref);
                }

                // Match `Evaluator::eval_arg`: ensure stable ordering for deterministic behavior.
                let mut areas: Vec<(SheetRangeRef, ResolvedRange)> = r
                    .areas
                    .iter()
                    .cloned()
                    .map(|area| {
                        let resolved = area.range.resolve(base);
                        (area, resolved)
                    })
                    .collect();
                areas.sort_by(|(a_area, a_range), (b_area, b_range)| {
                    cmp_sheet_ids_in_tab_order(grid, &a_area.sheet, &b_area.sheet)
                        .then_with(|| a_range.row_start.cmp(&b_range.row_start))
                        .then_with(|| a_range.col_start.cmp(&b_range.col_start))
                        .then_with(|| a_range.row_end.cmp(&b_range.row_end))
                        .then_with(|| a_range.col_end.cmp(&b_range.col_end))
                });

                // Excel reference unions behave like set union: overlapping cells should only be
                // visited once. For multi-area references on the *same* sheet (reference unions),
                // we deduplicate by tracking visited `(sheet, cell)` pairs while scanning each
                // area row-major. (This also preserves AST-like ordering for cases where overlap
                // removal would yield disjoint rectangles that must be interleaved row-by-row.)
                let mut seen: HashSet<(SheetId, CellCoord)> = HashSet::new();

                for (area, range) in areas {
                    if !range_in_bounds_on_sheet(grid, &area.sheet, range) {
                        return Value::Error(ErrorKind::Ref);
                    }
                    grid.record_reference_on_sheet(
                        &area.sheet,
                        CellCoord {
                            row: range.row_start,
                            col: range.col_start,
                        },
                        CellCoord {
                            row: range.row_end,
                            col: range.col_end,
                        },
                    );

                    for row in range.row_start..=range.row_end {
                        for col in range.col_start..=range.col_end {
                            if !seen.insert((area.sheet.clone(), CellCoord { row, col })) {
                                continue;
                            }
                            let v = grid.get_value_on_sheet(&area.sheet, CellCoord { row, col });
                            let s = match coerce_to_cow_str(&v) {
                                Ok(s) => s,
                                Err(e) => return Value::Error(e),
                            };
                            out.push_str(s.as_ref());
                        }
                    }
                }
            }
            other => {
                let s = match coerce_to_cow_str(other) {
                    Ok(s) => s,
                    Err(e) => return Value::Error(e),
                };
                out.push_str(s.as_ref());
            }
        }
    }

    Value::Text(out.into())
}

fn fn_concatenate(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let mut out = String::new();
    for arg in args {
        let v = apply_implicit_intersection(arg.clone(), grid, base);
        let s = match coerce_to_cow_str(&v) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        out.push_str(s.as_ref());
    }
    Value::Text(out.into())
}

fn fn_sum(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    let mut sum = 0.0;
    for arg in args {
        match arg {
            Value::Number(v) => sum += v,
            Value::Bool(v) => sum += if *v { 1.0 } else { 0.0 },
            Value::Array(a) => {
                if a.len() >= SIMD_ARRAY_MIN_LEN {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_sum = 0.0;
                    let mut saw_nan = false;

                    for v in a.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                } else if !saw_nan {
                                    buf[len] = *n;
                                    len += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        local_sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Lambda(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_) => {}
                        }
                    }

                    if saw_nan {
                        sum += f64::NAN;
                    } else {
                        if len > 0 {
                            local_sum += simd::sum_ignore_nan_f64(&buf[..len]);
                        }
                        sum += local_sum;
                    }
                } else {
                    for v in a.iter() {
                        match v {
                            Value::Number(n) => sum += n,
                            Value::Error(e) => return Value::Error(*e),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Lambda(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_) => {}
                        }
                    }
                }
            }
            Value::Range(r) => match sum_range(grid, r.resolve(base)) {
                Ok(v) => sum += v,
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                // Like XOR: preserve AST-like error precedence within each union area, even when
                // overlap subtraction yields multiple disjoint rectangles that must be interleaved
                // row-by-row for true row-major ordering.
                let mut current_area_idx: Option<usize> = None;
                let mut best_error_in_area: Option<(i32, i32, ErrorKind)> = None;

                for area in multirange_unique_areas(r, grid, base) {
                    if current_area_idx != Some(area.area_idx) {
                        if let Some((_, _, err)) = best_error_in_area {
                            return Value::Error(err);
                        }
                        current_area_idx = Some(area.area_idx);
                        best_error_in_area = None;
                    }

                    match sum_range_on_sheet_with_coord(grid, &area.sheet, area.range) {
                        Ok(v) => sum += v,
                        Err((coord, err)) => {
                            record_error_row_major(&mut best_error_in_area, coord, err)
                        }
                    }
                }

                if let Some((_, _, err)) = best_error_in_area {
                    return Value::Error(err);
                }
            }
            Value::Empty | Value::Missing => {}
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Text(s) => match parse_value_from_text(s) {
                Ok(v) => sum += v,
                Err(e) => return Value::Error(e),
            },
            Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
        }
    }
    Value::Number(sum)
}

fn fn_average(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut sum = 0.0;
    let mut count = 0usize;
    let mut saw_nan = false;
    for arg in args {
        match arg {
            Value::Number(v) => {
                if v.is_nan() {
                    saw_nan = true;
                    count += 1;
                } else if !saw_nan {
                    sum += v;
                    count += 1;
                }
            }
            Value::Bool(v) => {
                if !saw_nan {
                    sum += if *v { 1.0 } else { 0.0 };
                    count += 1;
                }
            }
            Value::Array(a) => {
                if a.len() >= SIMD_ARRAY_MIN_LEN {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_sum = 0.0;

                    for v in a.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                    count += 1;
                                } else if !saw_nan {
                                    buf[len] = *n;
                                    len += 1;
                                    count += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        local_sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Lambda(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_) => {}
                        }
                    }

                    if !saw_nan {
                        if len > 0 {
                            local_sum += simd::sum_ignore_nan_f64(&buf[..len]);
                        }
                        sum += local_sum;
                    }
                } else {
                    for v in a.iter() {
                        match v {
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                    count += 1;
                                } else if !saw_nan {
                                    sum += n;
                                    count += 1;
                                }
                            }
                            Value::Error(e) => return Value::Error(*e),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Lambda(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_) => {}
                        }
                    }
                }
            }
            Value::Range(r) => match sum_count_range(grid, r.resolve(base)) {
                Ok((s, c)) => {
                    if !saw_nan {
                        sum += s;
                        count += c;
                    }
                }
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                // Preserve AST-like error precedence within each union area (see fn_sum/xor).
                let mut current_area_idx: Option<usize> = None;
                let mut best_error_in_area: Option<(i32, i32, ErrorKind)> = None;

                for area in multirange_unique_areas(r, grid, base) {
                    if current_area_idx != Some(area.area_idx) {
                        if let Some((_, _, err)) = best_error_in_area {
                            return Value::Error(err);
                        }
                        current_area_idx = Some(area.area_idx);
                        best_error_in_area = None;
                    }

                    match sum_count_range_on_sheet_with_coord(grid, &area.sheet, area.range) {
                        Ok((s, c)) => {
                            if !saw_nan {
                                sum += s;
                                count += c;
                            }
                        }
                        Err((coord, err)) => {
                            record_error_row_major(&mut best_error_in_area, coord, err)
                        }
                    }
                }

                if let Some((_, _, err)) = best_error_in_area {
                    return Value::Error(err);
                }
            }
            Value::Empty | Value::Missing => {}
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Text(s) => match parse_value_from_text(s) {
                Ok(v) => {
                    if v.is_nan() {
                        saw_nan = true;
                        count += 1;
                    } else if !saw_nan {
                        sum += v;
                        count += 1;
                    }
                }
                Err(e) => return Value::Error(e),
            },
            Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
        }
    }
    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    if saw_nan {
        Value::Number(f64::NAN)
    } else {
        Value::Number(sum / count as f64)
    }
}

fn fn_min(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out: Option<f64> = None;
    let mut saw_nan_number = false;
    for arg in args {
        match arg {
            Value::Number(v) => {
                if v.is_nan() {
                    saw_nan_number = true;
                } else {
                    out = Some(out.map_or(*v, |prev| prev.min(*v)))
                }
            }
            Value::Bool(v) => {
                out = Some(out.map_or(if *v { 1.0 } else { 0.0 }, |prev| {
                    prev.min(if *v { 1.0 } else { 0.0 })
                }))
            }
            Value::Array(a) => {
                if a.len() >= SIMD_ARRAY_MIN_LEN {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_best: Option<f64> = None;

                    for v in a.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }
                                buf[len] = *n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::min_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.min(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_)
                            | Value::Lambda(_) => {}
                        }
                    }

                    if len > 0 {
                        if let Some(m) = simd::min_ignore_nan_f64(&buf[..len]) {
                            local_best = Some(local_best.map_or(m, |b| b.min(m)));
                        }
                    }
                    if let Some(m) = local_best {
                        out = Some(out.map_or(m, |prev| prev.min(m)));
                    }
                } else {
                    for v in a.iter() {
                        match v {
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                } else {
                                    out = Some(out.map_or(*n, |prev| prev.min(*n)))
                                }
                            }
                            Value::Error(e) => return Value::Error(*e),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_)
                            | Value::Lambda(_) => {}
                        }
                    }
                }
            }
            Value::Range(r) => match min_range(grid, r.resolve(base)) {
                Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.min(m))),
                Ok(None) => {}
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                // Preserve AST-like error precedence within each union area (see fn_sum/xor).
                let mut current_area_idx: Option<usize> = None;
                let mut best_error_in_area: Option<(i32, i32, ErrorKind)> = None;

                for area in multirange_unique_areas(r, grid, base) {
                    if current_area_idx != Some(area.area_idx) {
                        if let Some((_, _, err)) = best_error_in_area {
                            return Value::Error(err);
                        }
                        current_area_idx = Some(area.area_idx);
                        best_error_in_area = None;
                    }

                    match min_range_on_sheet_with_coord(grid, &area.sheet, area.range) {
                        Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.min(m))),
                        Ok(None) => {}
                        Err((coord, err)) => {
                            record_error_row_major(&mut best_error_in_area, coord, err)
                        }
                    }
                }

                if let Some((_, _, err)) = best_error_in_area {
                    return Value::Error(err);
                }
            }
            Value::Empty | Value::Missing => out = Some(out.map_or(0.0, |prev| prev.min(0.0))),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Text(s) => match parse_value_from_text(s) {
                Ok(v) => {
                    if v.is_nan() {
                        saw_nan_number = true;
                    } else {
                        out = Some(out.map_or(v, |prev| prev.min(v)))
                    }
                }
                Err(e) => return Value::Error(e),
            },
            Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
        }
    }
    Value::Number(match out {
        Some(v) => v,
        None if saw_nan_number => f64::NAN,
        None => 0.0,
    })
}

fn fn_max(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut out: Option<f64> = None;
    let mut saw_nan_number = false;
    for arg in args {
        match arg {
            Value::Number(v) => {
                if v.is_nan() {
                    saw_nan_number = true;
                } else {
                    out = Some(out.map_or(*v, |prev| prev.max(*v)))
                }
            }
            Value::Bool(v) => {
                out = Some(out.map_or(if *v { 1.0 } else { 0.0 }, |prev| {
                    prev.max(if *v { 1.0 } else { 0.0 })
                }))
            }
            Value::Array(a) => {
                if a.len() >= SIMD_ARRAY_MIN_LEN {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_best: Option<f64> = None;

                    for v in a.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }
                                buf[len] = *n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::max_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.max(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_)
                            | Value::Lambda(_) => {}
                        }
                    }

                    if len > 0 {
                        if let Some(m) = simd::max_ignore_nan_f64(&buf[..len]) {
                            local_best = Some(local_best.map_or(m, |b| b.max(m)));
                        }
                    }
                    if let Some(m) = local_best {
                        out = Some(out.map_or(m, |prev| prev.max(m)));
                    }
                } else {
                    for v in a.iter() {
                        match v {
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                } else {
                                    out = Some(out.map_or(*n, |prev| prev.max(*n)))
                                }
                            }
                            Value::Error(e) => return Value::Error(*e),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Empty
                            | Value::Missing
                            | Value::Array(_)
                            | Value::Range(_)
                            | Value::MultiRange(_)
                            | Value::Lambda(_) => {}
                        }
                    }
                }
            }
            Value::Range(r) => match max_range(grid, r.resolve(base)) {
                Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.max(m))),
                Ok(None) => {}
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                // Preserve AST-like error precedence within each union area (see fn_sum/xor).
                let mut current_area_idx: Option<usize> = None;
                let mut best_error_in_area: Option<(i32, i32, ErrorKind)> = None;

                for area in multirange_unique_areas(r, grid, base) {
                    if current_area_idx != Some(area.area_idx) {
                        if let Some((_, _, err)) = best_error_in_area {
                            return Value::Error(err);
                        }
                        current_area_idx = Some(area.area_idx);
                        best_error_in_area = None;
                    }

                    match max_range_on_sheet_with_coord(grid, &area.sheet, area.range) {
                        Ok(Some(m)) => out = Some(out.map_or(m, |prev| prev.max(m))),
                        Ok(None) => {}
                        Err((coord, err)) => {
                            record_error_row_major(&mut best_error_in_area, coord, err)
                        }
                    }
                }

                if let Some((_, _, err)) = best_error_in_area {
                    return Value::Error(err);
                }
            }
            Value::Empty | Value::Missing => out = Some(out.map_or(0.0, |prev| prev.max(0.0))),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Text(s) => match parse_value_from_text(s) {
                Ok(v) => {
                    if v.is_nan() {
                        saw_nan_number = true;
                    } else {
                        out = Some(out.map_or(v, |prev| prev.max(v)))
                    }
                }
                Err(e) => return Value::Error(e),
            },
            Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
        }
    }
    Value::Number(match out {
        Some(v) => v,
        None if saw_nan_number => f64::NAN,
        None => 0.0,
    })
}

fn fn_count(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    let mut count = 0usize;
    for arg in args {
        match arg {
            Value::Number(_) => count += 1,
            Value::Array(a) => {
                if a.len() >= SIMD_ARRAY_MIN_LEN {
                    let mut buf = [f64::NAN; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_count = 0usize;

                    for v in a.iter() {
                        buf[len] = if matches!(v, Value::Number(_)) {
                            0.0
                        } else {
                            f64::NAN
                        };
                        len += 1;
                        if len == SIMD_AGGREGATE_BLOCK {
                            local_count += simd::count_ignore_nan_f64(&buf);
                            len = 0;
                        }
                    }
                    if len > 0 {
                        local_count += simd::count_ignore_nan_f64(&buf[..len]);
                    }
                    count += local_count;
                } else {
                    count += a.iter().filter(|v| matches!(v, Value::Number(_))).count();
                }
            }
            Value::Range(r) => match count_range(grid, r.resolve(base)) {
                Ok(c) => count += c,
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                for area in multirange_unique_areas(r, grid, base) {
                    match count_range_on_sheet(grid, &area.sheet, area.range) {
                        Ok(c) => count += c,
                        Err(e) => return Value::Error(e),
                    }
                }
            }
            Value::Bool(_)
            | Value::Empty
            | Value::Missing
            | Value::Error(_)
            | Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Lambda(_) => {}
        }
    }
    Value::Number(count as f64)
}

fn fn_counta(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    let mut total = 0usize;
    for arg in args {
        match arg {
            // Scalars.
            Value::Empty | Value::Missing => {}
            Value::Number(_)
            | Value::Bool(_)
            | Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Error(_)
            | Value::Lambda(_) => total += 1,
            Value::Array(a) => {
                total += a
                    .iter()
                    .filter(|v| !matches!(v, Value::Empty | Value::Missing))
                    .count();
            }
            // References scan the grid.
            Value::Range(r) => match counta_range(grid, r.resolve(base)) {
                Ok(c) => total += c,
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                for area in multirange_unique_areas(r, grid, base) {
                    match counta_range_on_sheet(grid, &area.sheet, area.range) {
                        Ok(c) => total += c,
                        Err(e) => return Value::Error(e),
                    }
                }
            }
        }
    }
    Value::Number(total as f64)
}

fn fn_countblank(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.is_empty() {
        return Value::Error(ErrorKind::Value);
    }
    let mut total = 0usize;
    for arg in args {
        match arg {
            Value::Empty | Value::Missing => total += 1,
            Value::Text(s) if s.is_empty() => total += 1,
            Value::Entity(v) if v.display.is_empty() => total += 1,
            Value::Record(v) if v.display.is_empty() => total += 1,
            Value::Array(a) => {
                total += a
                    .iter()
                    .filter(|v| {
                        matches!(v, Value::Empty | Value::Missing)
                            || matches!(v, Value::Text(s) if s.is_empty())
                            || matches!(v, Value::Entity(ent) if ent.display.is_empty())
                            || matches!(v, Value::Record(rec) if rec.display.is_empty())
                    })
                    .count();
            }
            Value::Range(r) => match countblank_range(grid, r.resolve(base)) {
                Ok(c) => total += c,
                Err(e) => return Value::Error(e),
            },
            Value::MultiRange(r) => {
                for area in multirange_unique_areas(r, grid, base) {
                    match countblank_range_on_sheet(grid, &area.sheet, area.range) {
                        Ok(c) => total += c,
                        Err(e) => return Value::Error(e),
                    }
                }
            }
            Value::Number(_)
            | Value::Bool(_)
            | Value::Text(_)
            | Value::Error(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Lambda(_) => {}
        }
    }
    Value::Number(total as f64)
}

fn fn_countif(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }
    let range = match &args[0] {
        Value::Range(r) => RangeArg::Range(*r),
        Value::MultiRange(r) => RangeArg::MultiRange(r),
        Value::Array(a) => RangeArg::Array(a),
        _ => return Value::Error(ErrorKind::Value),
    };
    let criteria = match parse_countif_criteria(&args[1], locale) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    // Fast path: criteria that can be represented as a simple numeric comparator.
    if let Some(numeric) = criteria.as_numeric_criteria() {
        let count = match range {
            RangeArg::Range(r) => match count_if_range(grid, r.resolve(base), numeric) {
                Ok(c) => c,
                Err(e) => return Value::Error(e),
            },
            RangeArg::MultiRange(r) => {
                let mut count = 0usize;
                for area in multirange_unique_areas(r, grid, base) {
                    match count_if_range_on_sheet(grid, &area.sheet, area.range, numeric) {
                        Ok(c) => count += c,
                        Err(e) => return Value::Error(e),
                    }
                }
                count
            }
            RangeArg::Array(a) => count_if_array_numeric_criteria(a, numeric),
        };
        return Value::Number(count as f64);
    }

    let count = match range {
        RangeArg::Range(r) => match count_if_range_criteria(grid, r.resolve(base), &criteria) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        },
        RangeArg::MultiRange(r) => {
            let mut count = 0usize;
            for area in multirange_unique_areas(r, grid, base) {
                match count_if_range_criteria_on_sheet(grid, &area.sheet, area.range, &criteria) {
                    Ok(c) => count += c,
                    Err(e) => return Value::Error(e),
                }
            }
            count
        }
        RangeArg::Array(a) => count_if_array_criteria(a, &criteria),
    };
    Value::Number(count as f64)
}

fn fn_sumif(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() != 2 && args.len() != 3 {
        return Value::Error(ErrorKind::Value);
    }

    // Fast path: range-only SUMIF (existing optimized implementation).
    if matches!(args[0], Value::Range(_))
        && matches!(
            args.get(2),
            None | Some(Value::Missing) | Some(Value::Range(_))
        )
    {
        let criteria_range_ref = match &args[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let criteria = match parse_countif_criteria(&args[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        // Excel treats `SUMIF(range, criteria,)` the same as omitting the optional sum_range arg.
        let sum_range_ref = match args.get(2) {
            None => None,
            Some(Value::Missing) => None,
            Some(Value::Range(r)) => Some(*r),
            Some(_) => return Value::Error(ErrorKind::Value),
        }
        .unwrap_or(criteria_range_ref);

        let crit_range = criteria_range_ref.resolve(base);
        let sum_range = sum_range_ref.resolve(base);

        if !range_in_bounds(grid, crit_range) || !range_in_bounds(grid, sum_range) {
            return Value::Error(ErrorKind::Ref);
        }
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: crit_range.row_start,
                col: crit_range.col_start,
            },
            CellCoord {
                row: crit_range.row_end,
                col: crit_range.col_end,
            },
        );
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: sum_range.row_start,
                col: sum_range.col_start,
            },
            CellCoord {
                row: sum_range.row_end,
                col: sum_range.col_end,
            },
        );
        if crit_range.rows() != sum_range.rows() || crit_range.cols() != sum_range.cols() {
            return Value::Error(ErrorKind::Value);
        }

        let rows = crit_range.rows();
        let cols = crit_range.cols();
        if rows <= 0 || cols <= 0 {
            return Value::Number(0.0);
        }

        let cells = i64::from(rows)
            .checked_mul(i64::from(cols))
            .unwrap_or(i64::MAX);
        if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
            if let Some(iter) = grid.iter_cells() {
                let row_delta = crit_range.row_start - sum_range.row_start;
                let col_delta = crit_range.col_start - sum_range.col_start;
                let mut sum = 0.0;
                let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
                for (coord, v) in iter {
                    if !coord_in_range(coord, sum_range) {
                        continue;
                    }
                    let crit_cell = CellCoord {
                        row: coord.row + row_delta,
                        col: coord.col + col_delta,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(crit_cell));
                    if !criteria.matches(&engine_value) {
                        continue;
                    }
                    match v {
                        Value::Number(n) => sum += n,
                        Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Empty
                        | Value::Missing
                        | Value::Array(_)
                        | Value::Range(_)
                        | Value::MultiRange(_)
                        | Value::Lambda(_) => {}
                    }
                }
                if let Some((_, _, e)) = earliest_error {
                    return Value::Error(e);
                }
                return Value::Number(sum);
            }
        }

        if let Some(numeric) = criteria.as_numeric_criteria() {
            // Only use the numeric SIMD fast path when *all* required slices are available across the
            // full range. When any slice is missing (blocked rows, errors, etc.) we fall back to the
            // generic row-major scan so error precedence matches the AST evaluator.
            let mut slices_ok = true;
            for col_off in 0..cols {
                let crit_col = crit_range.col_start + col_off;
                let sum_col = sum_range.col_start + col_off;
                if grid
                    .column_slice_strict_numeric(crit_col, crit_range.row_start, crit_range.row_end)
                    .is_none()
                    || grid
                        .column_slice(sum_col, sum_range.row_start, sum_range.row_end)
                        .is_none()
                {
                    slices_ok = false;
                    break;
                }
            }

            if slices_ok {
                let mut sum = 0.0;
                for col_off in 0..cols {
                    let crit_col = crit_range.col_start + col_off;
                    let sum_col = sum_range.col_start + col_off;
                    let crit_slice = grid
                        .column_slice_strict_numeric(
                            crit_col,
                            crit_range.row_start,
                            crit_range.row_end,
                        )
                        .unwrap();
                    let sum_slice = grid
                        .column_slice(sum_col, sum_range.row_start, sum_range.row_end)
                        .unwrap();
                    sum += simd::sum_if_f64(sum_slice, crit_slice, numeric);
                }
                return Value::Number(sum);
            }
        }

        let mut sum = 0.0;
        for row_off in 0..rows {
            for col_off in 0..cols {
                let crit_cell = CellCoord {
                    row: crit_range.row_start + row_off,
                    col: crit_range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(crit_cell));
                if !criteria.matches(&engine_value) {
                    continue;
                }

                let sum_cell = CellCoord {
                    row: sum_range.row_start + row_off,
                    col: sum_range.col_start + col_off,
                };
                match grid.get_value(sum_cell) {
                    Value::Number(v) => sum += v,
                    Value::Error(e) => return Value::Error(e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
        }
        return Value::Number(sum);
    }

    // Generic path: support array arguments and mixed array/range cases (matching the AST
    // evaluator's Range2D semantics).
    let criteria_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let criteria = match parse_countif_criteria(&args[1], locale) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    let sum_range = match args.get(2) {
        None | Some(Value::Missing) => criteria_range,
        Some(v) => match range2d_from_value(v, grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        },
    };

    let rows = criteria_range.rows();
    let cols = criteria_range.cols();
    if sum_range.rows() != rows || sum_range.cols() != cols {
        return Value::Error(ErrorKind::Value);
    }
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    // SIMD fast path: SUMIF over in-memory arrays with numeric criteria.
    if let Some(numeric) = criteria.as_numeric_criteria() {
        if let (Range2DArg::Array(criteria_arr), Range2DArg::Array(sum_arr)) =
            (criteria_range, sum_range)
        {
            if let Some(sum) = sum_if_array_numeric_criteria(sum_arr, criteria_arr, numeric) {
                return Value::Number(sum);
            }
        }
    }

    let mut sum = 0.0;
    for row_off in 0..rows {
        for col_off in 0..cols {
            let crit_v = criteria_range.get_value_at(grid, row_off, col_off);
            let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
            if !criteria.matches(&engine_value) {
                continue;
            }

            match sum_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => sum += *v,
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    Value::Number(sum)
}

fn fn_sumifs(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    // Generic path: support array arguments and mixed array/range cases (matching the AST
    // evaluator's Range2D semantics). Preserve the existing optimized implementation as a fast
    // path for range-only inputs.
    if !matches!(args[0], Value::Range(_))
        || args[1..]
            .chunks_exact(2)
            .any(|pair| !matches!(pair[0], Value::Range(_)))
    {
        return sumifs_with_array_ranges(args, grid, base, locale);
    }

    let sum_range_ref = match &args[0] {
        Value::Range(r) => *r,
        _ => return Value::Error(ErrorKind::Value),
    };
    let sum_range = sum_range_ref.resolve(base);

    if !range_in_bounds(grid, sum_range) {
        return Value::Error(ErrorKind::Ref);
    }
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: sum_range.row_start,
            col: sum_range.col_start,
        },
        CellCoord {
            row: sum_range.row_end,
            col: sum_range.col_end,
        },
    );

    let rows = sum_range.rows();
    let cols = sum_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<ResolvedRange> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);
    let mut numeric_crits: Vec<NumericCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range_ref = match &pair[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let range = range_ref.resolve(base);
        if !range_in_bounds(grid, range) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        if let Some(nc) = crit.as_numeric_criteria() {
            numeric_crits.push(nc);
        } else {
            numeric_crits.clear();
        }

        crit_ranges.push(range);
        crits.push(crit);
    }
    for range in &crit_ranges {
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let all_numeric = !numeric_crits.is_empty() && numeric_crits.len() == crits.len();

    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);
    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
        if let Some(iter) = grid.iter_cells() {
            let mut sum = 0.0;
            let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
            'cell: for (coord, v) in iter {
                if !coord_in_range(coord, sum_range) {
                    continue;
                }
                let row_off = coord.row - sum_range.row_start;
                let col_off = coord.col - sum_range.col_start;
                for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                    let cell = CellCoord {
                        row: range.row_start + row_off,
                        col: range.col_start + col_off,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                    if !crit.matches(&engine_value) {
                        continue 'cell;
                    }
                }
                match v {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((_, _, e)) = earliest_error {
                return Value::Error(e);
            }
            return Value::Number(sum);
        }
    }

    if all_numeric {
        // Like MINIFS/MAXIFS, only take the numeric slice fast path when all slices are available
        // for the full rectangular region. Otherwise fall back to a row-major scan so error
        // precedence matches the AST evaluator.
        let mut slices_ok = true;
        for col_off in 0..cols {
            let sum_col = sum_range.col_start + col_off;
            if grid
                .column_slice(sum_col, sum_range.row_start, sum_range.row_end)
                .is_none()
            {
                slices_ok = false;
                break;
            }
            for range in &crit_ranges {
                let col = range.col_start + col_off;
                if grid
                    .column_slice_strict_numeric(col, range.row_start, range.row_end)
                    .is_none()
                {
                    slices_ok = false;
                    break;
                }
            }
            if !slices_ok {
                break;
            }
        }

        if slices_ok {
            let mut sum = 0.0;
            for col_off in 0..cols {
                let sum_col = sum_range.col_start + col_off;
                let sum_slice = grid
                    .column_slice(sum_col, sum_range.row_start, sum_range.row_end)
                    .unwrap();
                let mut crit_slices: SmallVec<[&[f64]; 4]> = SmallVec::with_capacity(crits.len());
                for range in &crit_ranges {
                    let col = range.col_start + col_off;
                    let slice = grid
                        .column_slice_strict_numeric(col, range.row_start, range.row_end)
                        .unwrap();
                    crit_slices.push(slice);
                }

                // Tight numeric scan.
                if numeric_crits.len() == 1 {
                    sum += simd::sum_if_f64(sum_slice, crit_slices[0], numeric_crits[0]);
                    continue;
                }

                let len = sum_slice.len();
                let mut i = 0usize;
                while i + 4 <= len {
                    for lane in 0..4 {
                        let idx = i + lane;
                        let mut matches = true;
                        for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                            let mut v = slice[idx];
                            if v.is_nan() {
                                v = 0.0;
                            }
                            if !matches_numeric_criteria(v, *crit) {
                                matches = false;
                                break;
                            }
                        }
                        if !matches {
                            continue;
                        }
                        let v = sum_slice[idx];
                        if !v.is_nan() {
                            sum += v;
                        }
                    }
                    i += 4;
                }
                for idx in i..len {
                    let mut matches = true;
                    for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                        let mut v = slice[idx];
                        if v.is_nan() {
                            v = 0.0;
                        }
                        if !matches_numeric_criteria(v, *crit) {
                            matches = false;
                            break;
                        }
                    }
                    if !matches {
                        continue;
                    }
                    let v = sum_slice[idx];
                    if !v.is_nan() {
                        sum += v;
                    }
                }
            }

            return Value::Number(sum);
        }
    }

    let mut sum = 0.0;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                let cell = CellCoord {
                    row: range.row_start + row_off,
                    col: range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match grid.get_value(CellCoord {
                row: sum_range.row_start + row_off,
                col: sum_range.col_start + col_off,
            }) {
                Value::Number(v) => sum += v,
                Value::Error(e) => return Value::Error(e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    Value::Number(sum)
}

fn sumifs_with_array_ranges(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    let sum_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let rows = sum_range.rows();
    let cols = sum_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<Range2DArg<'_>> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range = match range2d_from_value(&pair[0], grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        crit_ranges.push(range);
        crits.push(crit);
    }

    // SIMD fast path: single-criteria SUMIFS over in-memory arrays with numeric criteria.
    if crit_ranges.len() == 1 {
        if let Some(numeric) = crits[0].as_numeric_criteria() {
            if let (Range2DArg::Array(sum_arr), Range2DArg::Array(criteria_arr)) =
                (sum_range, crit_ranges[0])
            {
                if let Some(sum) = sum_if_array_numeric_criteria(sum_arr, criteria_arr, numeric) {
                    return Value::Number(sum);
                }
            }
        }
    }

    let mut sum = 0.0;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().copied().zip(crits.iter()) {
                let crit_v = range.get_value_at(grid, row_off, col_off);
                let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match sum_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => sum += *v,
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    Value::Number(sum)
}

fn fn_countifs(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() < 2 || args.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    if args
        .chunks_exact(2)
        .any(|pair| matches!(pair[0], Value::Array(_)))
    {
        return countifs_with_array_ranges(args, grid, base, locale);
    }

    let mut ranges: Vec<ResolvedRange> = Vec::with_capacity(args.len() / 2);
    let mut criteria: Vec<EngineCriteria> = Vec::with_capacity(args.len() / 2);
    let mut numeric: Vec<NumericCriteria> = Vec::with_capacity(args.len() / 2);

    for pair in args.chunks_exact(2) {
        let range_ref = match &pair[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let range = range_ref.resolve(base);
        if !range_in_bounds(grid, range) {
            return Value::Error(ErrorKind::Ref);
        }
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        if let Some(nc) = crit.as_numeric_criteria() {
            numeric.push(nc);
        } else {
            numeric.clear();
        }

        ranges.push(range);
        criteria.push(crit);
    }

    let (rows, cols) = (ranges[0].rows(), ranges[0].cols());
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }
    for range in &ranges[1..] {
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }
    }

    let all_numeric = !numeric.is_empty() && numeric.len() == criteria.len();
    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);

    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
        if let Some(iter) = grid.iter_cells() {
            let implicit_matches_all = criteria.iter().all(|c| c.matches(&EngineValue::Blank));
            let total_cells = cells;

            // Track the set of (row_off, col_off) offsets where at least one input range contains a
            // stored (non-implicit-blank) cell. Offsets not present in this set are implicit blanks
            // across *all* criteria ranges and can be accounted for in one shot.
            let mut offsets: HashSet<i64> = HashSet::new();
            for (coord, _) in iter {
                for range in &ranges {
                    if !coord_in_range(coord, *range) {
                        continue;
                    }
                    let row_off = coord.row - range.row_start;
                    let col_off = coord.col - range.col_start;
                    let key = ((row_off as i64) << 32) | (col_off as u32 as i64);
                    offsets.insert(key);
                }
            }

            let mut count: i64 = if implicit_matches_all { total_cells } else { 0 };

            for key in offsets {
                let row_off = (key >> 32) as i32;
                let col_off = (key as u32) as i32;
                let mut matches = true;
                for (range, crit) in ranges.iter().zip(criteria.iter()) {
                    let cell = CellCoord {
                        row: range.row_start + row_off,
                        col: range.col_start + col_off,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                    if !crit.matches(&engine_value) {
                        matches = false;
                        break;
                    }
                }

                if implicit_matches_all {
                    if !matches {
                        count -= 1;
                    }
                } else if matches {
                    count += 1;
                }
            }

            return Value::Number(count as f64);
        }
    }

    let mut count = 0usize;
    for col_off in 0..cols {
        if all_numeric {
            let mut slices: SmallVec<[&[f64]; 4]> = SmallVec::with_capacity(ranges.len());
            for range in &ranges {
                let col = range.col_start + col_off;
                let Some(slice) =
                    grid.column_slice_strict_numeric(col, range.row_start, range.row_end)
                else {
                    slices.clear();
                    break;
                };
                slices.push(slice);
            }

            if slices.len() == ranges.len() {
                if numeric.len() == 1 {
                    count += simd::count_if_blank_as_zero_f64(slices[0], numeric[0]);
                    continue;
                }

                let len = slices[0].len();
                let mut i = 0usize;
                while i + 4 <= len {
                    for lane in 0..4 {
                        let idx = i + lane;
                        let mut matches = true;
                        for (slice, crit) in slices.iter().zip(numeric.iter()) {
                            let mut v = slice[idx];
                            if v.is_nan() {
                                v = 0.0;
                            }
                            if !matches_numeric_criteria(v, *crit) {
                                matches = false;
                                break;
                            }
                        }
                        if matches {
                            count += 1;
                        }
                    }
                    i += 4;
                }
                for idx in i..len {
                    let mut matches = true;
                    for (slice, crit) in slices.iter().zip(numeric.iter()) {
                        let mut v = slice[idx];
                        if v.is_nan() {
                            v = 0.0;
                        }
                        if !matches_numeric_criteria(v, *crit) {
                            matches = false;
                            break;
                        }
                    }
                    if matches {
                        count += 1;
                    }
                }
                continue;
            }
        }

        // Fallback: per-cell scan for this column.
        'row: for row_off in 0..rows {
            for (range, crit) in ranges.iter().zip(criteria.iter()) {
                let cell = CellCoord {
                    row: range.row_start + row_off,
                    col: range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                if !crit.matches(&engine_value) {
                    continue 'row;
                }
            }
            count += 1;
        }
    }

    Value::Number(count as f64)
}

fn countifs_with_array_ranges(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    #[derive(Clone, Copy)]
    enum CriteriaRange<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }

    let mut ranges: Vec<CriteriaRange<'_>> = Vec::with_capacity(args.len() / 2);
    let mut criteria: Vec<EngineCriteria> = Vec::with_capacity(args.len() / 2);
    let mut shape: Option<(usize, usize)> = None;

    for pair in args.chunks_exact(2) {
        let range = match &pair[0] {
            Value::Range(r) => {
                let range = r.resolve(base);
                if !range_in_bounds(grid, range) {
                    return Value::Error(ErrorKind::Ref);
                }
                let rows = match usize::try_from(range.rows()) {
                    Ok(v) => v,
                    Err(_) => return Value::Error(ErrorKind::Num),
                };
                let cols = match usize::try_from(range.cols()) {
                    Ok(v) => v,
                    Err(_) => return Value::Error(ErrorKind::Num),
                };
                match shape {
                    None => shape = Some((rows, cols)),
                    Some((expected_rows, expected_cols)) => {
                        if (rows, cols) != (expected_rows, expected_cols) {
                            return Value::Error(ErrorKind::Value);
                        }
                    }
                }
                grid.record_reference(
                    grid.sheet_id(),
                    CellCoord {
                        row: range.row_start,
                        col: range.col_start,
                    },
                    CellCoord {
                        row: range.row_end,
                        col: range.col_end,
                    },
                );
                CriteriaRange::Range(range)
            }
            Value::Array(a) => {
                let rows = a.rows;
                let cols = a.cols;
                match shape {
                    None => shape = Some((rows, cols)),
                    Some((expected_rows, expected_cols)) => {
                        if (rows, cols) != (expected_rows, expected_cols) {
                            return Value::Error(ErrorKind::Value);
                        }
                    }
                }
                CriteriaRange::Array(a)
            }
            _ => return Value::Error(ErrorKind::Value),
        };

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        ranges.push(range);
        criteria.push(crit);
    }

    let (rows, cols) = shape.unwrap_or((0, 0));
    let len = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    if len == 0 {
        return Value::Number(0.0);
    }

    // Fast path: COUNTIFS with a single criteria pair is equivalent to COUNTIF.
    if ranges.len() == 1 {
        if let Some(numeric) = criteria[0].as_numeric_criteria() {
            if let CriteriaRange::Array(arr) = ranges[0] {
                return Value::Number(count_if_array_numeric_criteria(arr, numeric) as f64);
            }
        }
    }

    let mut count = 0usize;
    for idx in 0..len {
        let row_off = idx / cols;
        let col_off = idx % cols;

        let mut matches = true;
        for (range, crit) in ranges.iter().zip(criteria.iter()) {
            let engine_value = match range {
                CriteriaRange::Range(r) => bytecode_value_to_engine(grid.get_value(CellCoord {
                    row: r.row_start + row_off as i32,
                    col: r.col_start + col_off as i32,
                })),
                CriteriaRange::Array(a) => {
                    let v = a.values.get(idx).unwrap_or(&Value::Empty);
                    bytecode_value_to_engine_ref(v)
                }
            };
            if !crit.matches(&engine_value) {
                matches = false;
                break;
            }
        }

        if matches {
            count += 1;
        }
    }

    Value::Number(count as f64)
}

fn fn_averageif(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() != 2 && args.len() != 3 {
        return Value::Error(ErrorKind::Value);
    }

    // Fast path: range-only AVERAGEIF (existing optimized implementation).
    if matches!(args[0], Value::Range(_))
        && matches!(
            args.get(2),
            None | Some(Value::Missing) | Some(Value::Range(_))
        )
    {
        let criteria_range_ref = match &args[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let criteria = match parse_countif_criteria(&args[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        // Excel treats `AVERAGEIF(range, criteria,)` as omitting the optional average_range.
        let average_range_ref = match args.get(2) {
            None => None,
            Some(Value::Missing) => None,
            Some(Value::Range(r)) => Some(*r),
            Some(_) => return Value::Error(ErrorKind::Value),
        }
        .unwrap_or(criteria_range_ref);

        let crit_range = criteria_range_ref.resolve(base);
        let avg_range = average_range_ref.resolve(base);

        if !range_in_bounds(grid, crit_range) || !range_in_bounds(grid, avg_range) {
            return Value::Error(ErrorKind::Ref);
        }
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: crit_range.row_start,
                col: crit_range.col_start,
            },
            CellCoord {
                row: crit_range.row_end,
                col: crit_range.col_end,
            },
        );
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: avg_range.row_start,
                col: avg_range.col_start,
            },
            CellCoord {
                row: avg_range.row_end,
                col: avg_range.col_end,
            },
        );
        if crit_range.rows() != avg_range.rows() || crit_range.cols() != avg_range.cols() {
            return Value::Error(ErrorKind::Value);
        }

        let rows = crit_range.rows();
        let cols = crit_range.cols();
        if rows <= 0 || cols <= 0 {
            return Value::Error(ErrorKind::Div0);
        }

        let cells = i64::from(rows)
            .checked_mul(i64::from(cols))
            .unwrap_or(i64::MAX);
        if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
            if let Some(iter) = grid.iter_cells() {
                let row_delta = crit_range.row_start - avg_range.row_start;
                let col_delta = crit_range.col_start - avg_range.col_start;
                let mut sum = 0.0;
                let mut count = 0usize;
                let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
                for (coord, v) in iter {
                    if !coord_in_range(coord, avg_range) {
                        continue;
                    }
                    let crit_cell = CellCoord {
                        row: coord.row + row_delta,
                        col: coord.col + col_delta,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(crit_cell));
                    if !criteria.matches(&engine_value) {
                        continue;
                    }
                    match v {
                        Value::Number(n) => {
                            sum += n;
                            count += 1;
                        }
                        Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Empty
                        | Value::Missing
                        | Value::Array(_)
                        | Value::Range(_)
                        | Value::MultiRange(_)
                        | Value::Lambda(_) => {}
                    }
                }

                if let Some((_, _, e)) = earliest_error {
                    return Value::Error(e);
                }
                if count == 0 {
                    return Value::Error(ErrorKind::Div0);
                }
                return Value::Number(sum / count as f64);
            }
        }

        if let Some(numeric) = criteria.as_numeric_criteria() {
            // Only use the numeric SIMD fast path when *all* required slices are available across the
            // full range. When any slice is missing (blocked rows, errors, etc.) we fall back to the
            // generic row-major scan so error precedence matches the AST evaluator.
            let mut slices_ok = true;
            for col_off in 0..cols {
                let crit_col = crit_range.col_start + col_off;
                let avg_col = avg_range.col_start + col_off;
                if grid
                    .column_slice_strict_numeric(crit_col, crit_range.row_start, crit_range.row_end)
                    .is_none()
                    || grid
                        .column_slice(avg_col, avg_range.row_start, avg_range.row_end)
                        .is_none()
                {
                    slices_ok = false;
                    break;
                }
            }

            if slices_ok {
                let mut sum = 0.0;
                let mut count = 0usize;
                for col_off in 0..cols {
                    let crit_col = crit_range.col_start + col_off;
                    let avg_col = avg_range.col_start + col_off;
                    let crit_slice = grid
                        .column_slice_strict_numeric(
                            crit_col,
                            crit_range.row_start,
                            crit_range.row_end,
                        )
                        .unwrap();
                    let avg_slice = grid
                        .column_slice(avg_col, avg_range.row_start, avg_range.row_end)
                        .unwrap();
                    let (s, c) = simd::sum_count_if_f64(avg_slice, crit_slice, numeric);
                    sum += s;
                    count += c;
                }

                if count == 0 {
                    return Value::Error(ErrorKind::Div0);
                }
                return Value::Number(sum / count as f64);
            }
        }

        let mut sum = 0.0;
        let mut count = 0usize;
        for row_off in 0..rows {
            for col_off in 0..cols {
                let crit_cell = CellCoord {
                    row: crit_range.row_start + row_off,
                    col: crit_range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(crit_cell));
                if !criteria.matches(&engine_value) {
                    continue;
                }

                let avg_cell = CellCoord {
                    row: avg_range.row_start + row_off,
                    col: avg_range.col_start + col_off,
                };
                match grid.get_value(avg_cell) {
                    Value::Number(v) => {
                        sum += v;
                        count += 1;
                    }
                    Value::Error(e) => return Value::Error(e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
        }

        if count == 0 {
            return Value::Error(ErrorKind::Div0);
        }
        return Value::Number(sum / count as f64);
    }

    // Generic path: support array arguments and mixed array/range cases.
    let criteria_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let criteria = match parse_countif_criteria(&args[1], locale) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    let avg_range = match args.get(2) {
        None | Some(Value::Missing) => criteria_range,
        Some(v) => match range2d_from_value(v, grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        },
    };

    let rows = criteria_range.rows();
    let cols = criteria_range.cols();
    if avg_range.rows() != rows || avg_range.cols() != cols {
        return Value::Error(ErrorKind::Value);
    }
    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Div0);
    }

    // SIMD fast path: AVERAGEIF over in-memory arrays with numeric criteria.
    if let Some(numeric) = criteria.as_numeric_criteria() {
        if let (Range2DArg::Array(criteria_arr), Range2DArg::Array(avg_arr)) =
            (criteria_range, avg_range)
        {
            if let Some((sum, count)) =
                sum_count_if_array_numeric_criteria(avg_arr, criteria_arr, numeric)
            {
                if count == 0 {
                    return Value::Error(ErrorKind::Div0);
                }
                return Value::Number(sum / count as f64);
            }
        }
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    for row_off in 0..rows {
        for col_off in 0..cols {
            let crit_v = criteria_range.get_value_at(grid, row_off, col_off);
            let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
            if !criteria.matches(&engine_value) {
                continue;
            }

            match avg_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => {
                    sum += *v;
                    count += 1;
                }
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

fn fn_averageifs(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    // Generic path: support array arguments and mixed array/range cases (matching the AST
    // evaluator's Range2D semantics). Preserve the existing optimized implementation as a fast
    // path for range-only inputs.
    if !matches!(args[0], Value::Range(_))
        || args[1..]
            .chunks_exact(2)
            .any(|pair| !matches!(pair[0], Value::Range(_)))
    {
        return averageifs_with_array_ranges(args, grid, base, locale);
    }

    let avg_range_ref = match &args[0] {
        Value::Range(r) => *r,
        _ => return Value::Error(ErrorKind::Value),
    };
    let avg_range = avg_range_ref.resolve(base);
    if !range_in_bounds(grid, avg_range) {
        return Value::Error(ErrorKind::Ref);
    }
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: avg_range.row_start,
            col: avg_range.col_start,
        },
        CellCoord {
            row: avg_range.row_end,
            col: avg_range.col_end,
        },
    );

    let rows = avg_range.rows();
    let cols = avg_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Div0);
    }

    let mut crit_ranges: Vec<ResolvedRange> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);
    let mut numeric_crits: Vec<NumericCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range_ref = match &pair[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let range = range_ref.resolve(base);
        if !range_in_bounds(grid, range) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        if let Some(nc) = crit.as_numeric_criteria() {
            numeric_crits.push(nc);
        } else {
            numeric_crits.clear();
        }

        crit_ranges.push(range);
        crits.push(crit);
    }
    for range in &crit_ranges {
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let all_numeric = !numeric_crits.is_empty() && numeric_crits.len() == crits.len();
    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);

    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
        if let Some(iter) = grid.iter_cells() {
            let mut sum = 0.0;
            let mut count = 0usize;
            let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
            'cell: for (coord, v) in iter {
                if !coord_in_range(coord, avg_range) {
                    continue;
                }
                let row_off = coord.row - avg_range.row_start;
                let col_off = coord.col - avg_range.col_start;
                for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                    let cell = CellCoord {
                        row: range.row_start + row_off,
                        col: range.col_start + col_off,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                    if !crit.matches(&engine_value) {
                        continue 'cell;
                    }
                }

                match v {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }

            if let Some((_, _, e)) = earliest_error {
                return Value::Error(e);
            }
            if count == 0 {
                return Value::Error(ErrorKind::Div0);
            }
            return Value::Number(sum / count as f64);
        }
    }

    if all_numeric {
        // Like MINIFS/MAXIFS, only take the numeric slice fast path when all slices are available
        // for the full rectangular region. Otherwise fall back to a row-major scan so error
        // precedence matches the AST evaluator.
        let mut slices_ok = true;
        for col_off in 0..cols {
            let avg_col = avg_range.col_start + col_off;
            if grid
                .column_slice(avg_col, avg_range.row_start, avg_range.row_end)
                .is_none()
            {
                slices_ok = false;
                break;
            }
            for range in &crit_ranges {
                let col = range.col_start + col_off;
                if grid
                    .column_slice_strict_numeric(col, range.row_start, range.row_end)
                    .is_none()
                {
                    slices_ok = false;
                    break;
                }
            }
            if !slices_ok {
                break;
            }
        }

        if slices_ok {
            let mut sum = 0.0;
            let mut count = 0usize;
            for col_off in 0..cols {
                let avg_col = avg_range.col_start + col_off;
                let avg_slice = grid
                    .column_slice(avg_col, avg_range.row_start, avg_range.row_end)
                    .unwrap();
                let mut crit_slices: SmallVec<[&[f64]; 4]> = SmallVec::with_capacity(crits.len());
                for range in &crit_ranges {
                    let col = range.col_start + col_off;
                    let slice = grid
                        .column_slice_strict_numeric(col, range.row_start, range.row_end)
                        .unwrap();
                    crit_slices.push(slice);
                }

                if numeric_crits.len() == 1 {
                    let (s, c) =
                        simd::sum_count_if_f64(avg_slice, crit_slices[0], numeric_crits[0]);
                    sum += s;
                    count += c;
                    continue;
                }

                let len = avg_slice.len();
                let mut i = 0usize;
                while i + 4 <= len {
                    for lane in 0..4 {
                        let idx = i + lane;
                        let mut matches = true;
                        for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                            let mut v = slice[idx];
                            if v.is_nan() {
                                v = 0.0;
                            }
                            if !matches_numeric_criteria(v, *crit) {
                                matches = false;
                                break;
                            }
                        }
                        if !matches {
                            continue;
                        }
                        let v = avg_slice[idx];
                        if !v.is_nan() {
                            sum += v;
                            count += 1;
                        }
                    }
                    i += 4;
                }
                for idx in i..len {
                    let mut matches = true;
                    for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                        let mut v = slice[idx];
                        if v.is_nan() {
                            v = 0.0;
                        }
                        if !matches_numeric_criteria(v, *crit) {
                            matches = false;
                            break;
                        }
                    }
                    if !matches {
                        continue;
                    }
                    let v = avg_slice[idx];
                    if !v.is_nan() {
                        sum += v;
                        count += 1;
                    }
                }
            }

            if count == 0 {
                return Value::Error(ErrorKind::Div0);
            }
            return Value::Number(sum / count as f64);
        }
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                let cell = CellCoord {
                    row: range.row_start + row_off,
                    col: range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match grid.get_value(CellCoord {
                row: avg_range.row_start + row_off,
                col: avg_range.col_start + col_off,
            }) {
                Value::Number(v) => {
                    sum += v;
                    count += 1;
                }
                Value::Error(e) => return Value::Error(e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

fn averageifs_with_array_ranges(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    let avg_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let rows = avg_range.rows();
    let cols = avg_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Div0);
    }

    let mut crit_ranges: Vec<Range2DArg<'_>> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range = match range2d_from_value(&pair[0], grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        crit_ranges.push(range);
        crits.push(crit);
    }

    // SIMD fast path: single-criteria AVERAGEIFS over in-memory arrays with numeric criteria.
    if crit_ranges.len() == 1 {
        if let Some(numeric) = crits[0].as_numeric_criteria() {
            if let (Range2DArg::Array(avg_arr), Range2DArg::Array(criteria_arr)) =
                (avg_range, crit_ranges[0])
            {
                if let Some((sum, count)) =
                    sum_count_if_array_numeric_criteria(avg_arr, criteria_arr, numeric)
                {
                    if count == 0 {
                        return Value::Error(ErrorKind::Div0);
                    }
                    return Value::Number(sum / count as f64);
                }
            }
        }
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().copied().zip(crits.iter()) {
                let crit_v = range.get_value_at(grid, row_off, col_off);
                let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match avg_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => {
                    sum += *v;
                    count += 1;
                }
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

fn fn_minifs(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    // Generic path: support array arguments and mixed array/range cases (matching the AST
    // evaluator's Range2D semantics). Preserve the existing optimized implementation as a fast
    // path for range-only inputs.
    if !matches!(args[0], Value::Range(_))
        || args[1..]
            .chunks_exact(2)
            .any(|pair| !matches!(pair[0], Value::Range(_)))
    {
        return minifs_with_array_ranges(args, grid, base, locale);
    }

    let min_range_ref = match &args[0] {
        Value::Range(r) => *r,
        _ => return Value::Error(ErrorKind::Value),
    };
    let min_range = min_range_ref.resolve(base);
    if !range_in_bounds(grid, min_range) {
        return Value::Error(ErrorKind::Ref);
    }
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: min_range.row_start,
            col: min_range.col_start,
        },
        CellCoord {
            row: min_range.row_end,
            col: min_range.col_end,
        },
    );

    let rows = min_range.rows();
    let cols = min_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<ResolvedRange> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);
    let mut numeric_crits: Vec<NumericCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range_ref = match &pair[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let range = range_ref.resolve(base);
        if !range_in_bounds(grid, range) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        if let Some(nc) = crit.as_numeric_criteria() {
            numeric_crits.push(nc);
        } else {
            numeric_crits.clear();
        }

        crit_ranges.push(range);
        crits.push(crit);
    }
    for range in &crit_ranges {
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let all_numeric = !numeric_crits.is_empty() && numeric_crits.len() == crits.len();
    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);

    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
        if let Some(iter) = grid.iter_cells() {
            let mut best: Option<f64> = None;
            let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
            'cell: for (coord, v) in iter {
                if !coord_in_range(coord, min_range) {
                    continue;
                }
                let row_off = coord.row - min_range.row_start;
                let col_off = coord.col - min_range.col_start;
                for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                    let cell = CellCoord {
                        row: range.row_start + row_off,
                        col: range.col_start + col_off,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                    if !crit.matches(&engine_value) {
                        continue 'cell;
                    }
                }

                match v {
                    Value::Number(n) => best = Some(best.map_or(n, |b| b.min(n))),
                    Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }

            if let Some((_, _, e)) = earliest_error {
                return Value::Error(e);
            }
            return Value::Number(best.unwrap_or(0.0));
        }
    }

    if all_numeric {
        // Only use the numeric fast path when all required slices are available (no blocked rows).
        let mut slices_ok = true;
        let mut best: Option<f64> = None;
        for col_off in 0..cols {
            let min_col = min_range.col_start + col_off;
            let Some(min_slice) =
                grid.column_slice(min_col, min_range.row_start, min_range.row_end)
            else {
                slices_ok = false;
                break;
            };

            let mut crit_slices: SmallVec<[&[f64]; 4]> = SmallVec::with_capacity(crits.len());
            for range in &crit_ranges {
                let col = range.col_start + col_off;
                let Some(slice) =
                    grid.column_slice_strict_numeric(col, range.row_start, range.row_end)
                else {
                    crit_slices.clear();
                    break;
                };
                crit_slices.push(slice);
            }
            if crit_slices.len() != crits.len() {
                slices_ok = false;
                break;
            }

            if numeric_crits.len() == 1 {
                if let Some(col_best) =
                    simd::min_if_f64(min_slice, crit_slices[0], numeric_crits[0])
                {
                    best = Some(best.map_or(col_best, |b| b.min(col_best)));
                }
                continue;
            }

            for idx in 0..min_slice.len() {
                let mut matches = true;
                for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                    let mut v = slice[idx];
                    if v.is_nan() {
                        v = 0.0;
                    }
                    if !matches_numeric_criteria(v, *crit) {
                        matches = false;
                        break;
                    }
                }
                if !matches {
                    continue;
                }
                let v = min_slice[idx];
                if v.is_nan() {
                    continue;
                }
                best = Some(best.map_or(v, |b| b.min(v)));
            }
        }

        if slices_ok {
            return Value::Number(best.unwrap_or(0.0));
        }
    }

    // Fallback: row-major scan with stable error propagation.
    let mut best: Option<f64> = None;
    let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
    for row_off in 0..rows {
        'col: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                let cell = CellCoord {
                    row: range.row_start + row_off,
                    col: range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                if !crit.matches(&engine_value) {
                    continue 'col;
                }
            }

            let value_cell = CellCoord {
                row: min_range.row_start + row_off,
                col: min_range.col_start + col_off,
            };
            match grid.get_value(value_cell) {
                Value::Number(v) => best = Some(best.map_or(v, |b| b.min(v))),
                Value::Error(e) => match earliest_error {
                    None => earliest_error = Some((row_off, col_off, e)),
                    Some((best_row, best_col, _)) => {
                        if (row_off, col_off) < (best_row, best_col) {
                            earliest_error = Some((row_off, col_off, e));
                        }
                    }
                },
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    if let Some((_, _, e)) = earliest_error {
        return Value::Error(e);
    }
    Value::Number(best.unwrap_or(0.0))
}

fn minifs_with_array_ranges(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    let min_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let rows = min_range.rows();
    let cols = min_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<Range2DArg<'_>> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range = match range2d_from_value(&pair[0], grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        crit_ranges.push(range);
        crits.push(crit);
    }

    let mut best: Option<f64> = None;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().copied().zip(crits.iter()) {
                let crit_v = range.get_value_at(grid, row_off, col_off);
                let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match min_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => best = Some(best.map_or(*v, |b| b.min(*v))),
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    Value::Number(best.unwrap_or(0.0))
}

fn fn_maxifs(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    // Generic path: support array arguments and mixed array/range cases (matching the AST
    // evaluator's Range2D semantics). Preserve the existing optimized implementation as a fast
    // path for range-only inputs.
    if !matches!(args[0], Value::Range(_))
        || args[1..]
            .chunks_exact(2)
            .any(|pair| !matches!(pair[0], Value::Range(_)))
    {
        return maxifs_with_array_ranges(args, grid, base, locale);
    }

    let max_range_ref = match &args[0] {
        Value::Range(r) => *r,
        _ => return Value::Error(ErrorKind::Value),
    };
    let max_range = max_range_ref.resolve(base);
    if !range_in_bounds(grid, max_range) {
        return Value::Error(ErrorKind::Ref);
    }
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: max_range.row_start,
            col: max_range.col_start,
        },
        CellCoord {
            row: max_range.row_end,
            col: max_range.col_end,
        },
    );

    let rows = max_range.rows();
    let cols = max_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<ResolvedRange> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);
    let mut numeric_crits: Vec<NumericCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range_ref = match &pair[0] {
            Value::Range(r) => *r,
            _ => return Value::Error(ErrorKind::Value),
        };
        let range = range_ref.resolve(base);
        if !range_in_bounds(grid, range) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        if let Some(nc) = crit.as_numeric_criteria() {
            numeric_crits.push(nc);
        } else {
            numeric_crits.clear();
        }

        crit_ranges.push(range);
        crits.push(crit);
    }
    for range in &crit_ranges {
        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let all_numeric = !numeric_crits.is_empty() && numeric_crits.len() == crits.len();
    let cells = i64::from(rows)
        .checked_mul(i64::from(cols))
        .unwrap_or(i64::MAX);

    if rows > BYTECODE_SPARSE_RANGE_ROW_THRESHOLD || cells > BYTECODE_MAX_RANGE_CELLS as i64 {
        if let Some(iter) = grid.iter_cells() {
            let mut best: Option<f64> = None;
            let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
            'cell: for (coord, v) in iter {
                if !coord_in_range(coord, max_range) {
                    continue;
                }
                let row_off = coord.row - max_range.row_start;
                let col_off = coord.col - max_range.col_start;
                for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                    let cell = CellCoord {
                        row: range.row_start + row_off,
                        col: range.col_start + col_off,
                    };
                    let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                    if !crit.matches(&engine_value) {
                        continue 'cell;
                    }
                }

                match v {
                    Value::Number(n) => best = Some(best.map_or(n, |b| b.max(n))),
                    Value::Error(e) => record_error_row_major(&mut earliest_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }

            if let Some((_, _, e)) = earliest_error {
                return Value::Error(e);
            }
            return Value::Number(best.unwrap_or(0.0));
        }
    }

    if all_numeric {
        // Only use the numeric fast path when all required slices are available (no blocked rows).
        let mut slices_ok = true;
        let mut best: Option<f64> = None;
        for col_off in 0..cols {
            let max_col = max_range.col_start + col_off;
            let Some(max_slice) =
                grid.column_slice(max_col, max_range.row_start, max_range.row_end)
            else {
                slices_ok = false;
                break;
            };

            let mut crit_slices: SmallVec<[&[f64]; 4]> = SmallVec::with_capacity(crits.len());
            for range in &crit_ranges {
                let col = range.col_start + col_off;
                let Some(slice) =
                    grid.column_slice_strict_numeric(col, range.row_start, range.row_end)
                else {
                    crit_slices.clear();
                    break;
                };
                crit_slices.push(slice);
            }
            if crit_slices.len() != crits.len() {
                slices_ok = false;
                break;
            }

            if numeric_crits.len() == 1 {
                if let Some(col_best) =
                    simd::max_if_f64(max_slice, crit_slices[0], numeric_crits[0])
                {
                    best = Some(best.map_or(col_best, |b| b.max(col_best)));
                }
                continue;
            }

            for idx in 0..max_slice.len() {
                let mut matches = true;
                for (slice, crit) in crit_slices.iter().zip(numeric_crits.iter()) {
                    let mut v = slice[idx];
                    if v.is_nan() {
                        v = 0.0;
                    }
                    if !matches_numeric_criteria(v, *crit) {
                        matches = false;
                        break;
                    }
                }
                if !matches {
                    continue;
                }
                let v = max_slice[idx];
                if v.is_nan() {
                    continue;
                }
                best = Some(best.map_or(v, |b| b.max(v)));
            }
        }

        if slices_ok {
            return Value::Number(best.unwrap_or(0.0));
        }
    }

    // Fallback: row-major scan with stable error propagation.
    let mut best: Option<f64> = None;
    let mut earliest_error: Option<(i32, i32, ErrorKind)> = None;
    for row_off in 0..rows {
        'col: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().zip(crits.iter()) {
                let cell = CellCoord {
                    row: range.row_start + row_off,
                    col: range.col_start + col_off,
                };
                let engine_value = bytecode_value_to_engine(grid.get_value(cell));
                if !crit.matches(&engine_value) {
                    continue 'col;
                }
            }

            let value_cell = CellCoord {
                row: max_range.row_start + row_off,
                col: max_range.col_start + col_off,
            };
            match grid.get_value(value_cell) {
                Value::Number(v) => best = Some(best.map_or(v, |b| b.max(v))),
                Value::Error(e) => match earliest_error {
                    None => earliest_error = Some((row_off, col_off, e)),
                    Some((best_row, best_col, _)) => {
                        if (row_off, col_off) < (best_row, best_col) {
                            earliest_error = Some((row_off, col_off, e));
                        }
                    }
                },
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    if let Some((_, _, e)) = earliest_error {
        return Value::Error(e);
    }
    Value::Number(best.unwrap_or(0.0))
}

fn maxifs_with_array_ranges(
    args: &[Value],
    grid: &dyn Grid,
    base: CellCoord,
    locale: &crate::LocaleConfig,
) -> Value {
    let max_range = match range2d_from_value(&args[0], grid, base) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    let rows = max_range.rows();
    let cols = max_range.cols();
    if rows <= 0 || cols <= 0 {
        return Value::Number(0.0);
    }

    let mut crit_ranges: Vec<Range2DArg<'_>> = Vec::with_capacity((args.len() - 1) / 2);
    let mut crits: Vec<EngineCriteria> = Vec::with_capacity((args.len() - 1) / 2);

    for pair in args[1..].chunks_exact(2) {
        let range = match range2d_from_value(&pair[0], grid, base) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if range.rows() != rows || range.cols() != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit = match parse_countif_criteria(&pair[1], locale) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        crit_ranges.push(range);
        crits.push(crit);
    }

    let mut best: Option<f64> = None;
    for row_off in 0..rows {
        'cell: for col_off in 0..cols {
            for (range, crit) in crit_ranges.iter().copied().zip(crits.iter()) {
                let crit_v = range.get_value_at(grid, row_off, col_off);
                let engine_value = bytecode_value_to_engine_ref(crit_v.as_ref());
                if !crit.matches(&engine_value) {
                    continue 'cell;
                }
            }

            match max_range.get_value_at(grid, row_off, col_off).as_ref() {
                Value::Number(v) => best = Some(best.map_or(*v, |b| b.max(*v))),
                Value::Error(e) => return Value::Error(*e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }

    Value::Number(best.unwrap_or(0.0))
}

fn fn_sumproduct(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }

    // Preserve Excel-like argument error precedence.
    if let Value::Error(e) = &args[0] {
        return Value::Error(*e);
    }
    if let Value::Error(e) = &args[1] {
        return Value::Error(*e);
    }

    struct RangeOperand<'a> {
        grid: &'a dyn Grid,
        range: ResolvedRange,
        rows: usize,
        cols: usize,
        // Per-column strict-numeric slices (numbers/blanks only) for faster reads when available.
        col_slices: Vec<Option<&'a [f64]>>,
    }

    impl<'a> RangeOperand<'a> {
        fn new(grid: &'a dyn Grid, range: ResolvedRange) -> Result<Self, ErrorKind> {
            if !range_in_bounds(grid, range) {
                return Err(ErrorKind::Ref);
            }
            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: range.row_start,
                    col: range.col_start,
                },
                CellCoord {
                    row: range.row_end,
                    col: range.col_end,
                },
            );

            let rows_i32 = range.rows();
            let cols_i32 = range.cols();
            if rows_i32 <= 0 || cols_i32 <= 0 {
                return Err(ErrorKind::Value);
            }
            let rows = rows_i32 as usize;
            let cols = cols_i32 as usize;
            let mut col_slices = Vec::with_capacity(cols);
            for col in range.col_start..=range.col_end {
                col_slices.push(grid.column_slice_strict_numeric(
                    col,
                    range.row_start,
                    range.row_end,
                ));
            }
            Ok(Self {
                grid,
                range,
                rows,
                cols,
                col_slices,
            })
        }

        fn len(&self) -> usize {
            self.rows.saturating_mul(self.cols)
        }

        fn coerce_number_at(&self, idx: usize) -> Result<f64, ErrorKind> {
            if self.cols == 0 {
                return Err(ErrorKind::Value);
            }
            let row_offset = idx / self.cols;
            let col_offset = idx % self.cols;
            if row_offset >= self.rows {
                return Err(ErrorKind::Value);
            }

            if let Some(slice) = self.col_slices[col_offset] {
                let n = slice[row_offset];
                // Column slices represent blanks as NaN; SUMPRODUCT treats blanks as 0.
                return Ok(if n.is_nan() { 0.0 } else { n });
            }

            let row = self.range.row_start + row_offset as i32;
            let col = self.range.col_start + col_offset as i32;
            coerce_sumproduct_number(&self.grid.get_value(CellCoord { row, col }))
        }
    }

    enum Operand<'a> {
        Scalar(&'a Value),
        Array(&'a ArrayValue),
        Range(RangeOperand<'a>),
        MultiRange,
    }

    impl Operand<'_> {
        fn len(&self) -> usize {
            match self {
                Operand::Scalar(_) => 1,
                Operand::Array(a) => a.len(),
                Operand::Range(r) => r.len(),
                Operand::MultiRange => 0,
            }
        }

        fn coerce_number_at(&self, idx: usize) -> Result<f64, ErrorKind> {
            match self {
                Operand::Scalar(v) => coerce_sumproduct_number(v),
                Operand::Array(arr) => {
                    let v = arr.values.get(idx).ok_or(ErrorKind::Value)?;
                    coerce_sumproduct_number(v)
                }
                Operand::Range(r) => r.coerce_number_at(idx),
                Operand::MultiRange => Err(ErrorKind::Value),
            }
        }
    }

    let a = match &args[0] {
        Value::Range(r) => match RangeOperand::new(grid, r.resolve(base)) {
            Ok(v) => Operand::Range(v),
            Err(e) => return Value::Error(e),
        },
        Value::Array(arr) => Operand::Array(arr),
        Value::MultiRange(_) => Operand::MultiRange,
        other => Operand::Scalar(other),
    };
    let b = match &args[1] {
        Value::Range(r) => match RangeOperand::new(grid, r.resolve(base)) {
            Ok(v) => Operand::Range(v),
            Err(e) => return Value::Error(e),
        },
        Value::Array(arr) => Operand::Array(arr),
        Value::MultiRange(_) => Operand::MultiRange,
        other => Operand::Scalar(other),
    };

    let len_a = a.len();
    let len_b = b.len();
    let len = len_a.max(len_b);
    if len == 0 {
        return Value::Error(ErrorKind::Value);
    }
    if (len_a != len && len_a != 1) || (len_b != len && len_b != 1) {
        return Value::Error(ErrorKind::Value);
    }

    // Range fast paths:
    // - When every column has strict-numeric slices we can use the SIMD-optimized `sumproduct_range`.
    // - For huge ranges (e.g. `A:A`), `sumproduct_range` also has a sparse-iteration mode that avoids
    //   scanning implicit blanks when the grid supports `iter_cells()`.
    if let (Operand::Range(ra), Operand::Range(rb)) = (&a, &b) {
        if len_a == len_b && len_a == len && ra.rows == rb.rows && ra.cols == rb.cols {
            // Prefer the sparse-aware `sumproduct_range` implementation for huge ranges even when
            // column slices are unavailable (e.g. because the column cache builder skipped `A:A`).
            if ra.rows > (BYTECODE_SPARSE_RANGE_ROW_THRESHOLD as usize) {
                return match sumproduct_range(grid, ra.range, rb.range) {
                    Ok(v) => Value::Number(v),
                    Err(e) => Value::Error(e),
                };
            }

            let all_slices = ra
                .col_slices
                .iter()
                .zip(rb.col_slices.iter())
                .all(|(sa, sb)| sa.is_some() && sb.is_some());
            if all_slices {
                return match sumproduct_range(grid, ra.range, rb.range) {
                    Ok(v) => Value::Number(v),
                    Err(e) => Value::Error(e),
                };
            }
        }
    }

    let result = (|| -> Result<f64, ErrorKind> {
        if len_a == 1 && len_b == 1 {
            let x = a.coerce_number_at(0)?;
            let y = b.coerce_number_at(0)?;
            return Ok(x * y);
        }

        if len_a == 1 {
            let x = a.coerce_number_at(0)?;
            let mut sum = 0.0;
            for idx in 0..len {
                let y = b.coerce_number_at(if len_b == 1 { 0 } else { idx })?;
                sum += x * y;
            }
            return Ok(sum);
        }

        if len_b == 1 {
            // Preserve error precedence: for idx=0 we must coerce `a[0]` before `b[0]`.
            let x0 = a.coerce_number_at(0)?;
            let y = b.coerce_number_at(0)?;
            let mut sum = x0 * y;
            for idx in 1..len {
                let x = a.coerce_number_at(idx)?;
                sum += x * y;
            }
            return Ok(sum);
        }

        let mut sum = 0.0;
        for idx in 0..len {
            let x = a.coerce_number_at(idx)?;
            let y = b.coerce_number_at(idx)?;
            sum += x * y;
        }
        Ok(sum)
    })();

    match result {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

fn fn_vlookup(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 3 || args.len() > 4 {
        return Value::Error(ErrorKind::Value);
    }

    let lookup_value = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(apply_implicit_intersection(args[0].clone(), grid, base))
        }
        _ => None,
    };
    let lookup_value = lookup_value.as_ref().unwrap_or(&args[0]);

    if let Value::Error(e) = lookup_value {
        return Value::Error(*e);
    }
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(
        lookup_value,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_)
    ) {
        return Value::Error(ErrorKind::Spill);
    }

    enum LookupTable<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }

    let table = match &args[1] {
        Value::Range(r) => LookupTable::Range(r.resolve(base)),
        Value::Array(a) => LookupTable::Array(a),
        Value::Error(e) => return Value::Error(*e),
        Value::MultiRange(_) => return Value::Error(ErrorKind::Value),
        _ => return Value::Error(ErrorKind::Value),
    };
    if let LookupTable::Range(table) = &table {
        if !range_in_bounds(grid, *table) {
            return Value::Error(ErrorKind::Ref);
        }

        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: table.row_start,
                col: table.col_start,
            },
            CellCoord {
                row: table.row_end,
                col: table.col_end,
            },
        );
    }

    let col_index = match coerce_to_i64(&args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if col_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let cols = match &table {
        LookupTable::Range(r) => r.cols() as i64,
        LookupTable::Array(a) => a.cols as i64,
    };
    if col_index > cols {
        return Value::Error(ErrorKind::Ref);
    }

    let approx = if args.len() == 4 {
        match coerce_to_bool(&args[3]) {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    match table {
        LookupTable::Range(table) => {
            let row_offset = if approx {
                match approximate_match_in_first_col(grid, lookup_value, table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_col(grid, lookup_value, table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let row = table.row_start + row_offset;
            let col = table.col_start + (col_index as i32) - 1;
            grid.get_value(CellCoord { row, col })
        }
        LookupTable::Array(table) => {
            let row_offset = if approx {
                match approximate_match_in_first_col_array(lookup_value, table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_col_array(lookup_value, table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let row = match usize::try_from(row_offset) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::NA),
            };
            let col = match usize::try_from(col_index - 1) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Ref),
            };
            table.get(row, col).cloned().unwrap_or(Value::Empty)
        }
    }
}

fn fn_hlookup(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 3 || args.len() > 4 {
        return Value::Error(ErrorKind::Value);
    }

    let lookup_value = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(apply_implicit_intersection(args[0].clone(), grid, base))
        }
        _ => None,
    };
    let lookup_value = lookup_value.as_ref().unwrap_or(&args[0]);

    if let Value::Error(e) = lookup_value {
        return Value::Error(*e);
    }
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(
        lookup_value,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_)
    ) {
        return Value::Error(ErrorKind::Spill);
    }

    enum LookupTable<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }

    let table = match &args[1] {
        Value::Range(r) => LookupTable::Range(r.resolve(base)),
        Value::Array(a) => LookupTable::Array(a),
        Value::Error(e) => return Value::Error(*e),
        Value::MultiRange(_) => return Value::Error(ErrorKind::Value),
        _ => return Value::Error(ErrorKind::Value),
    };
    if let LookupTable::Range(table) = &table {
        if !range_in_bounds(grid, *table) {
            return Value::Error(ErrorKind::Ref);
        }

        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: table.row_start,
                col: table.col_start,
            },
            CellCoord {
                row: table.row_end,
                col: table.col_end,
            },
        );
    }

    let row_index = match coerce_to_i64(&args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if row_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let rows = match &table {
        LookupTable::Range(r) => r.rows() as i64,
        LookupTable::Array(a) => a.rows as i64,
    };
    if row_index > rows {
        return Value::Error(ErrorKind::Ref);
    }

    let approx = if args.len() == 4 {
        match coerce_to_bool(&args[3]) {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    match table {
        LookupTable::Range(table) => {
            let col_offset = if approx {
                match approximate_match_in_first_row(grid, lookup_value, table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_row(grid, lookup_value, table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let row = table.row_start + (row_index as i32) - 1;
            let col = table.col_start + col_offset;
            grid.get_value(CellCoord { row, col })
        }
        LookupTable::Array(table) => {
            let col_offset = if approx {
                match approximate_match_in_first_row_array(lookup_value, table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_row_array(lookup_value, table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let row = match usize::try_from(row_index - 1) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Ref),
            };
            let col = match usize::try_from(col_offset) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::NA),
            };
            table.get(row, col).cloned().unwrap_or(Value::Empty)
        }
    }
}

fn fn_match(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 2 || args.len() > 3 {
        return Value::Error(ErrorKind::Value);
    }

    let lookup_value = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(apply_implicit_intersection(args[0].clone(), grid, base))
        }
        _ => None,
    };
    let lookup_value = lookup_value.as_ref().unwrap_or(&args[0]);

    if let Value::Error(e) = lookup_value {
        return Value::Error(*e);
    }
    // MATCH treats LAMBDA values as invalid lookup values (Excel returns #VALUE! when a lambda is
    // passed where a scalar is expected). Keep behavior aligned with the AST evaluator's MATCH
    // implementation in `functions/builtins_lookup.rs`.
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(
        lookup_value,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_)
    ) {
        return Value::Error(ErrorKind::Spill);
    }
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let match_type = if args.len() == 3 {
        match coerce_to_i64(&args[2]) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    enum LookupArray<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }

    let lookup_array = match &args[1] {
        Value::Range(r) => LookupArray::Range(r.resolve(base)),
        Value::Array(a) => LookupArray::Array(a),
        Value::Error(e) => return Value::Error(*e),
        Value::MultiRange(_) => return Value::Error(ErrorKind::Value),
        _ => return Value::Error(ErrorKind::Value),
    };

    if let LookupArray::Range(range) = &lookup_array {
        if !range_in_bounds(grid, *range) {
            return Value::Error(ErrorKind::Ref);
        }

        grid.record_reference(
            grid.sheet_id(),
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let pos = match lookup_array {
        LookupArray::Range(range) => {
            if range.row_start == range.row_end {
                let len = range.cols() as usize;
                let row = range.row_start;
                let value_at = |idx: usize| {
                    grid.get_value(CellCoord {
                        row,
                        col: range.col_start + idx as i32,
                    })
                };
                match match_type {
                    0 => exact_match_1d(lookup_value, len, &value_at),
                    1 => approximate_match_1d(lookup_value, len, &value_at, true),
                    -1 => approximate_match_1d(lookup_value, len, &value_at, false),
                    _ => return Value::Error(ErrorKind::NA),
                }
            } else if range.col_start == range.col_end {
                let len = range.rows() as usize;
                let col = range.col_start;
                let value_at = |idx: usize| {
                    grid.get_value(CellCoord {
                        row: range.row_start + idx as i32,
                        col,
                    })
                };
                match match_type {
                    0 => exact_match_1d(lookup_value, len, &value_at),
                    1 => approximate_match_1d(lookup_value, len, &value_at, true),
                    -1 => approximate_match_1d(lookup_value, len, &value_at, false),
                    _ => return Value::Error(ErrorKind::NA),
                }
            } else {
                // MATCH requires a 1D array/range.
                return Value::Error(ErrorKind::NA);
            }
        }
        LookupArray::Array(arr) => {
            if arr.rows != 1 && arr.cols != 1 {
                return Value::Error(ErrorKind::NA);
            }
            let len = arr.len();
            if len == 0 {
                return Value::Error(ErrorKind::NA);
            }
            match match_type {
                0 => exact_match_1d_slice(lookup_value, arr.values.as_slice()),
                1 => approximate_match_1d_slice(lookup_value, arr.values.as_slice(), true),
                -1 => approximate_match_1d_slice(lookup_value, arr.values.as_slice(), false),
                _ => return Value::Error(ErrorKind::NA),
            }
        }
    };

    match pos {
        Some(idx) => Value::Number((idx + 1) as f64),
        None => Value::Error(ErrorKind::NA),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum XlookupVectorShape {
    Horizontal,
    Vertical,
}

fn resolved_range_1d_shape_len(range: ResolvedRange) -> Option<(XlookupVectorShape, usize)> {
    let rows = range.rows();
    let cols = range.cols();
    if rows <= 0 || cols <= 0 {
        return None;
    }
    if rows == 1 && cols == 1 {
        // Match `builtins_lookup`: treat a single-cell lookup vector as vertical.
        return Some((XlookupVectorShape::Vertical, 1));
    }
    if rows == 1 {
        return Some((XlookupVectorShape::Horizontal, usize::try_from(cols).ok()?));
    }
    if cols == 1 {
        return Some((XlookupVectorShape::Vertical, usize::try_from(rows).ok()?));
    }
    None
}

fn array_1d_shape_len(array: &ArrayValue) -> Option<(XlookupVectorShape, usize)> {
    let rows = array.rows;
    let cols = array.cols;
    if rows == 0 || cols == 0 {
        return None;
    }
    if rows == 1 && cols == 1 {
        // Match `builtins_lookup`: treat a single-cell lookup vector as vertical.
        return Some((XlookupVectorShape::Vertical, 1));
    }
    if rows == 1 {
        return Some((XlookupVectorShape::Horizontal, cols));
    }
    if cols == 1 {
        return Some((XlookupVectorShape::Vertical, rows));
    }
    None
}

fn parse_xmatch_match_mode(
    arg: Option<&Value>,
    grid: &dyn Grid,
    base: CellCoord,
) -> Result<lookup::MatchMode, ErrorKind> {
    match arg {
        None | Some(Value::Missing) => Ok(lookup::MatchMode::Exact),
        Some(v) => {
            // match_mode is a scalar argument; when provided as a range reference Excel applies
            // implicit intersection based on the formula cell (matching `eval_scalar_arg` in the
            // AST evaluator).
            let v = match v {
                Value::Range(_) | Value::MultiRange(_) => {
                    apply_implicit_intersection(v.clone(), grid, base)
                }
                _ => v.clone(),
            };
            if let Value::Error(e) = v {
                return Err(e);
            }
            // Arrays are not valid mode arguments (they cannot be implicitly intersected).
            if matches!(v, Value::Array(_) | Value::Range(_) | Value::MultiRange(_)) {
                return Err(ErrorKind::Value);
            }
            let n = coerce_to_i64(&v)?;
            lookup::MatchMode::try_from(n).map_err(ErrorKind::from)
        }
    }
}

fn parse_xmatch_search_mode(
    arg: Option<&Value>,
    grid: &dyn Grid,
    base: CellCoord,
) -> Result<lookup::SearchMode, ErrorKind> {
    match arg {
        None | Some(Value::Missing) => Ok(lookup::SearchMode::FirstToLast),
        Some(v) => {
            // search_mode is a scalar argument; when provided as a range reference Excel applies
            // implicit intersection based on the formula cell (matching `eval_scalar_arg` in the
            // AST evaluator).
            let v = match v {
                Value::Range(_) | Value::MultiRange(_) => {
                    apply_implicit_intersection(v.clone(), grid, base)
                }
                _ => v.clone(),
            };
            if let Value::Error(e) = v {
                return Err(e);
            }
            // Arrays are not valid mode arguments (they cannot be implicitly intersected).
            if matches!(v, Value::Array(_) | Value::Range(_) | Value::MultiRange(_)) {
                return Err(ErrorKind::Value);
            }
            let n = coerce_to_i64(&v)?;
            lookup::SearchMode::try_from(n).map_err(ErrorKind::from)
        }
    }
}

fn fn_xmatch(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 2 || args.len() > 4 {
        return Value::Error(ErrorKind::Value);
    }
    // `lookup_value` is a scalar argument; when provided as a range reference Excel applies implicit
    // intersection based on the formula cell. Match the AST evaluator's `eval_scalar_arg` behavior.
    let lookup_value = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            apply_implicit_intersection(args[0].clone(), grid, base)
        }
        // Arrays cannot be used as scalar lookup values.
        Value::Array(_) => return Value::Error(ErrorKind::Spill),
        _ => args[0].clone(),
    };
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }
    // XMATCH treats LAMBDA values as invalid lookup values (Excel returns #VALUE! when a lambda is
    // passed where a scalar is expected). Keep behavior aligned with the AST evaluator which
    // rejects `Value::Lambda` in `lookup::xmatch_with_modes_impl`.
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(
        lookup_value,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_)
    ) {
        return Value::Error(ErrorKind::Spill);
    }
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let match_mode = match parse_xmatch_match_mode(args.get(2), grid, base) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };
    let search_mode = match parse_xmatch_search_mode(args.get(3), grid, base) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };

    let pos = match &args[1] {
        Value::Range(r) => {
            let lookup_value = bytecode_value_to_engine_ref(&lookup_value);
            let lookup_range = r.resolve(base);
            if !range_in_bounds(grid, lookup_range) {
                return Value::Error(ErrorKind::Ref);
            }

            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: lookup_range.row_start,
                    col: lookup_range.col_start,
                },
                CellCoord {
                    row: lookup_range.row_end,
                    col: lookup_range.col_end,
                },
            );

            let Some((shape, len)) = resolved_range_1d_shape_len(lookup_range) else {
                return Value::Error(ErrorKind::Value);
            };

            lookup::xmatch_with_modes_accessor_with_locale(
                &lookup_value,
                len,
                |idx| {
                    let coord = match shape {
                        XlookupVectorShape::Vertical => CellCoord {
                            row: lookup_range.row_start + idx as i32,
                            col: lookup_range.col_start,
                        },
                        XlookupVectorShape::Horizontal => CellCoord {
                            row: lookup_range.row_start,
                            col: lookup_range.col_start + idx as i32,
                        },
                    };
                    bytecode_value_to_engine(grid.get_value(coord))
                },
                match_mode,
                search_mode,
                thread_value_locale(),
                thread_date_system(),
                thread_now_utc(),
            )
        }
        Value::Array(arr) => {
            let lookup_value = bytecode_value_to_engine_ref(&lookup_value);
            let Some((shape, len)) = array_1d_shape_len(arr) else {
                return Value::Error(ErrorKind::Value);
            };

            lookup::xmatch_with_modes_accessor_with_locale(
                &lookup_value,
                len,
                |idx| {
                    let raw_idx = match shape {
                        XlookupVectorShape::Vertical => idx.saturating_mul(arr.cols),
                        XlookupVectorShape::Horizontal => idx,
                    };
                    let v = arr.values.get(raw_idx).unwrap_or(&Value::Empty);
                    bytecode_value_to_engine_ref(v)
                },
                match_mode,
                search_mode,
                thread_value_locale(),
                thread_date_system(),
                thread_now_utc(),
            )
        }
        Value::Error(e) => return Value::Error(*e),
        _ => return Value::Error(ErrorKind::Value),
    };

    match pos {
        Ok(p) => Value::Number(p as f64),
        Err(e) => Value::Error(ErrorKind::from(e)),
    }
}

fn fn_xlookup(args: &[Value], grid: &dyn Grid, base: CellCoord) -> Value {
    if args.len() < 3 || args.len() > 6 {
        return Value::Error(ErrorKind::Value);
    }
    // `lookup_value` is a scalar argument; when provided as a range reference Excel applies implicit
    // intersection based on the formula cell. Match the AST evaluator's `eval_scalar_arg` behavior.
    let lookup_value = match &args[0] {
        Value::Range(_) | Value::MultiRange(_) => {
            apply_implicit_intersection(args[0].clone(), grid, base)
        }
        // Arrays cannot be used as scalar lookup values.
        Value::Array(_) => return Value::Error(ErrorKind::Spill),
        _ => args[0].clone(),
    };
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }
    // XLOOKUP treats LAMBDA values as invalid lookup values (Excel returns #VALUE! when a lambda is
    // passed where a scalar is expected). Keep behavior aligned with the AST evaluator.
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    if matches!(
        lookup_value,
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_)
    ) {
        return Value::Error(ErrorKind::Spill);
    }

    let if_not_found_arg = match args.get(3) {
        None | Some(Value::Missing) => None,
        Some(v) => Some(v),
    };
    // `if_not_found` is scalar when provided as a reference/range (Excel applies implicit
    // intersection), but may also be an array literal/expression (which can spill). Preserve array
    // values, but eagerly intersect references to match the AST evaluator's `eval_scalar_arg`
    // behavior.
    let if_not_found_intersected = if_not_found_arg.and_then(|v| match v {
        Value::Range(_) | Value::MultiRange(_) => {
            Some(apply_implicit_intersection(v.clone(), grid, base))
        }
        _ => None,
    });

    let match_mode = match parse_xmatch_match_mode(args.get(4), grid, base) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };
    let search_mode = match parse_xmatch_search_mode(args.get(5), grid, base) {
        Ok(m) => m,
        Err(e) => return Value::Error(e),
    };

    #[derive(Clone, Copy)]
    enum LookupVector<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }
    #[derive(Clone, Copy)]
    enum ReturnArray<'a> {
        Range(ResolvedRange),
        Array(&'a ArrayValue),
    }

    let lookup_vector = match &args[1] {
        Value::Range(r) => {
            let lookup_range = r.resolve(base);
            if !range_in_bounds(grid, lookup_range) {
                return Value::Error(ErrorKind::Ref);
            }
            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: lookup_range.row_start,
                    col: lookup_range.col_start,
                },
                CellCoord {
                    row: lookup_range.row_end,
                    col: lookup_range.col_end,
                },
            );
            LookupVector::Range(lookup_range)
        }
        Value::Array(arr) => LookupVector::Array(arr),
        Value::Error(e) => return Value::Error(*e),
        _ => return Value::Error(ErrorKind::Value),
    };
    let return_array = match &args[2] {
        Value::Range(r) => {
            let return_range = r.resolve(base);
            if !range_in_bounds(grid, return_range) {
                return Value::Error(ErrorKind::Ref);
            }
            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: return_range.row_start,
                    col: return_range.col_start,
                },
                CellCoord {
                    row: return_range.row_end,
                    col: return_range.col_end,
                },
            );
            ReturnArray::Range(return_range)
        }
        Value::Array(arr) => ReturnArray::Array(arr),
        Value::Error(e) => return Value::Error(*e),
        _ => return Value::Error(ErrorKind::Value),
    };

    let (lookup_shape, lookup_len) = match lookup_vector {
        LookupVector::Range(range) => match resolved_range_1d_shape_len(range) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Value),
        },
        LookupVector::Array(arr) => match array_1d_shape_len(arr) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Value),
        },
    };

    // Validate return_array shape:
    // - vertical lookup_array (Nx1) requires return_array.rows == N; result spills horizontally.
    // - horizontal lookup_array (1xN) requires return_array.cols == N; result spills vertically.
    match return_array {
        ReturnArray::Range(r) => {
            let lookup_len_i32 = match i32::try_from(lookup_len) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Value),
            };
            match lookup_shape {
                XlookupVectorShape::Vertical => {
                    if r.rows() != lookup_len_i32 {
                        return Value::Error(ErrorKind::Value);
                    }
                }
                XlookupVectorShape::Horizontal => {
                    if r.cols() != lookup_len_i32 {
                        return Value::Error(ErrorKind::Value);
                    }
                }
            }
        }
        ReturnArray::Array(arr) => match lookup_shape {
            XlookupVectorShape::Vertical => {
                if arr.rows != lookup_len {
                    return Value::Error(ErrorKind::Value);
                }
            }
            XlookupVectorShape::Horizontal => {
                if arr.cols != lookup_len {
                    return Value::Error(ErrorKind::Value);
                }
            }
        },
    };

    let lookup_value_engine = bytecode_value_to_engine(lookup_value);
    let match_pos = lookup::xmatch_with_modes_accessor_with_locale(
        &lookup_value_engine,
        lookup_len,
        |idx| match lookup_vector {
            LookupVector::Range(range) => {
                let coord = match lookup_shape {
                    XlookupVectorShape::Vertical => CellCoord {
                        row: range.row_start + idx as i32,
                        col: range.col_start,
                    },
                    XlookupVectorShape::Horizontal => CellCoord {
                        row: range.row_start,
                        col: range.col_start + idx as i32,
                    },
                };
                bytecode_value_to_engine(grid.get_value(coord))
            }
            LookupVector::Array(arr) => {
                let raw_idx = match lookup_shape {
                    XlookupVectorShape::Vertical => idx.saturating_mul(arr.cols),
                    XlookupVectorShape::Horizontal => idx,
                };
                let v = arr.values.get(raw_idx).unwrap_or(&Value::Empty);
                bytecode_value_to_engine_ref(v)
            }
        },
        match_mode,
        search_mode,
        thread_value_locale(),
        thread_date_system(),
        thread_now_utc(),
    );

    let match_pos = match match_pos {
        Ok(p) => p,
        Err(EngineErrorKind::NA) => {
            return match (if_not_found_arg, if_not_found_intersected) {
                (None, _) => Value::Error(ErrorKind::NA),
                (Some(_), Some(v)) => v,
                (Some(v), None) => v.clone(),
            };
        }
        Err(e) => return Value::Error(ErrorKind::from(e)),
    };

    let idx = match match_pos
        .checked_sub(1)
        .and_then(|v| usize::try_from(v).ok())
    {
        Some(v) if v < lookup_len => v,
        _ => return Value::Error(ErrorKind::Value),
    };

    match lookup_shape {
        XlookupVectorShape::Vertical => {
            // Return the matched row from return_array.
            match return_array {
                ReturnArray::Range(return_range) => {
                    let row = return_range.row_start + idx as i32;
                    let cols = match usize::try_from(return_range.cols()) {
                        Ok(v) => v,
                        Err(_) => return Value::Error(ErrorKind::Value),
                    };
                    if cols == 1 {
                        return grid.get_value(CellCoord {
                            row,
                            col: return_range.col_start,
                        });
                    }
                    let mut values = Vec::with_capacity(cols);
                    for col_offset in 0..cols {
                        values.push(grid.get_value(CellCoord {
                            row,
                            col: return_range.col_start + col_offset as i32,
                        }));
                    }
                    Value::Array(ArrayValue::new(1, cols, values))
                }
                ReturnArray::Array(arr) => {
                    let cols = arr.cols;
                    if cols == 1 {
                        return arr
                            .values
                            .get(idx.saturating_mul(cols))
                            .cloned()
                            .unwrap_or(Value::Empty);
                    }
                    let mut values = Vec::with_capacity(cols);
                    let row_start = idx.saturating_mul(cols);
                    for col_offset in 0..cols {
                        values.push(
                            arr.values
                                .get(row_start + col_offset)
                                .cloned()
                                .unwrap_or(Value::Empty),
                        );
                    }
                    Value::Array(ArrayValue::new(1, cols, values))
                }
            }
        }
        XlookupVectorShape::Horizontal => {
            // Return the matched column from return_array.
            match return_array {
                ReturnArray::Range(return_range) => {
                    let col = return_range.col_start + idx as i32;
                    let rows = match usize::try_from(return_range.rows()) {
                        Ok(v) => v,
                        Err(_) => return Value::Error(ErrorKind::Value),
                    };
                    if rows == 1 {
                        return grid.get_value(CellCoord {
                            row: return_range.row_start,
                            col,
                        });
                    }
                    let mut values = Vec::with_capacity(rows);
                    for row_offset in 0..rows {
                        values.push(grid.get_value(CellCoord {
                            row: return_range.row_start + row_offset as i32,
                            col,
                        }));
                    }
                    Value::Array(ArrayValue::new(rows, 1, values))
                }
                ReturnArray::Array(arr) => {
                    let rows = arr.rows;
                    if rows == 1 {
                        return arr.values.get(idx).cloned().unwrap_or(Value::Empty);
                    }
                    let mut values = Vec::with_capacity(rows);
                    for row_offset in 0..rows {
                        let raw_idx = row_offset.saturating_mul(arr.cols).saturating_add(idx);
                        values.push(arr.values.get(raw_idx).cloned().unwrap_or(Value::Empty));
                    }
                    Value::Array(ArrayValue::new(rows, 1, values))
                }
            }
        }
    }
}

fn wildcard_pattern_for_lookup(lookup: &Value) -> Option<WildcardPattern> {
    let pattern = match lookup {
        Value::Text(pattern) => Cow::Borrowed(pattern.as_ref()),
        Value::Entity(v) => Cow::Borrowed(v.display.as_str()),
        Value::Record(v) => match coerce_to_cow_str(lookup) {
            Ok(pattern) => pattern,
            Err(_) => Cow::Borrowed(v.display.as_str()),
        },
        _ => return None,
    };
    if !pattern.contains('*') && !pattern.contains('?') && !pattern.contains('~') {
        return None;
    }
    Some(WildcardPattern::new(pattern.as_ref()))
}

fn values_equal_for_lookup(lookup_value: &Value, candidate: &Value) -> bool {
    fn text_like_str(v: &Value) -> Option<&str> {
        match v {
            Value::Text(s) => Some(s.as_ref()),
            Value::Entity(v) => Some(v.display.as_str()),
            Value::Record(v) => Some(v.display.as_str()),
            _ => None,
        }
    }

    match (lookup_value, candidate) {
        (Value::Number(a), Value::Number(b)) => a == b,
        (a, b) if text_like_str(a).is_some() && text_like_str(b).is_some() => {
            cmp_case_insensitive(text_like_str(a).unwrap(), text_like_str(b).unwrap())
                == Ordering::Equal
        }
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::Error(a), Value::Error(b)) => a == b,
        (Value::Empty | Value::Missing, Value::Empty | Value::Missing) => true,
        (Value::Number(a), b) if text_like_str(b).is_some() => {
            let trimmed = text_like_str(b).unwrap().trim();
            if trimmed.is_empty() {
                false
            } else {
                crate::coercion::datetime::parse_value_text(
                    trimmed,
                    thread_value_locale(),
                    thread_now_utc(),
                    thread_date_system(),
                )
                .is_ok_and(|parsed| parsed == *a)
            }
        }
        (a, Value::Number(b)) if text_like_str(a).is_some() => {
            let trimmed = text_like_str(a).unwrap().trim();
            if trimmed.is_empty() {
                false
            } else {
                crate::coercion::datetime::parse_value_text(
                    trimmed,
                    thread_value_locale(),
                    thread_now_utc(),
                    thread_date_system(),
                )
                .is_ok_and(|parsed| parsed == *b)
            }
        }
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => {
            (*b == 0.0 && !*a) || (*b == 1.0 && *a)
        }
        _ => false,
    }
}

fn error_code(e: ErrorKind) -> u8 {
    EngineErrorKind::from(e).code()
}

fn excel_le(a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(a, b).map(|o| o <= 0)
}

fn excel_ge(a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(a, b).map(|o| o >= 0)
}

fn excel_cmp(a: &Value, b: &Value) -> Option<i32> {
    fn ordering_to_i32(ord: std::cmp::Ordering) -> i32 {
        match ord {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }
    }

    fn text_like_str(v: &Value) -> Option<&str> {
        match v {
            Value::Text(s) => Some(s.as_ref()),
            Value::Entity(v) => Some(v.display.as_str()),
            Value::Record(v) => Some(v.display.as_str()),
            _ => None,
        }
    }

    fn type_rank(v: &Value) -> Option<u8> {
        match v {
            Value::Number(_) => Some(0),
            Value::Text(_) | Value::Entity(_) | Value::Record(_) => Some(1),
            Value::Bool(_) => Some(2),
            Value::Empty | Value::Missing => Some(3),
            Value::Error(_) => Some(4),
            Value::Array(_) | Value::Range(_) | Value::MultiRange(_) | Value::Lambda(_) => None,
        }
    }

    match (a, b) {
        // Blank coerces to the other type (Excel semantics).
        (Value::Empty | Value::Missing, Value::Number(y)) => {
            Some(ordering_to_i32(0.0_f64.partial_cmp(y)?))
        }
        (Value::Number(x), Value::Empty | Value::Missing) => {
            Some(ordering_to_i32(x.partial_cmp(&0.0_f64)?))
        }
        (Value::Empty | Value::Missing, other) if text_like_str(other).is_some() => Some(
            ordering_to_i32(cmp_case_insensitive("", text_like_str(other).unwrap())),
        ),
        (other, Value::Empty | Value::Missing) if text_like_str(other).is_some() => Some(
            ordering_to_i32(cmp_case_insensitive(text_like_str(other).unwrap(), "")),
        ),
        (Value::Empty | Value::Missing, Value::Bool(y)) => Some(ordering_to_i32(false.cmp(y))),
        (Value::Bool(x), Value::Empty | Value::Missing) => Some(ordering_to_i32(x.cmp(&false))),
        _ => {
            let ra = type_rank(a)?;
            let rb = type_rank(b)?;
            if ra != rb {
                return Some(ordering_to_i32(ra.cmp(&rb)));
            }

            match (a, b) {
                (Value::Number(x), Value::Number(y)) => Some(ordering_to_i32(x.partial_cmp(y)?)),
                (a, b) if type_rank(a) == Some(1) => Some(ordering_to_i32(cmp_case_insensitive(
                    text_like_str(a)?,
                    text_like_str(b)?,
                ))),
                (Value::Bool(x), Value::Bool(y)) => Some(ordering_to_i32(x.cmp(y))),
                (Value::Empty | Value::Missing, Value::Empty | Value::Missing) => Some(0),
                (Value::Error(x), Value::Error(y)) => {
                    Some(ordering_to_i32(error_code(*x).cmp(&error_code(*y))))
                }
                _ => None,
            }
        }
    }
}

fn coerce_to_string_for_lookup(v: &Value) -> Result<Cow<'_, str>, ErrorKind> {
    // Use the bytecode runtime's locale-aware coercion (shared with CONCAT/& fixes) so wildcard
    // matching behaves consistently across backends.
    coerce_to_cow_str(v)
}

fn exact_match_in_first_col(grid: &dyn Grid, lookup: &Value, table: ResolvedRange) -> Option<i32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    let rows = table.row_start..=table.row_end;
    for (idx, row) in rows.enumerate() {
        let key = grid.get_value(CellCoord {
            row,
            col: table.col_start,
        });
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(&key) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as i32);
            }
        } else if values_equal_for_lookup(lookup, &key) {
            return Some(idx as i32);
        }
    }
    None
}

fn exact_match_in_first_col_array(lookup: &Value, table: &ArrayValue) -> Option<i32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for row in 0..table.rows {
        let key = table.get(row, 0).unwrap_or(&Value::Empty);
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(key) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(row as i32);
            }
        } else if values_equal_for_lookup(lookup, key) {
            return Some(row as i32);
        }
    }
    None
}

fn exact_match_in_first_row(grid: &dyn Grid, lookup: &Value, table: ResolvedRange) -> Option<i32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    let cols = table.col_start..=table.col_end;
    for (idx, col) in cols.enumerate() {
        let key = grid.get_value(CellCoord {
            row: table.row_start,
            col,
        });
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(&key) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as i32);
            }
        } else if values_equal_for_lookup(lookup, &key) {
            return Some(idx as i32);
        }
    }
    None
}

fn exact_match_in_first_row_array(lookup: &Value, table: &ArrayValue) -> Option<i32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for col in 0..table.cols {
        let key = table.get(0, col).unwrap_or(&Value::Empty);
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(key) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(col as i32);
            }
        } else if values_equal_for_lookup(lookup, key) {
            return Some(col as i32);
        }
    }
    None
}

fn approximate_match_in_first_col(
    grid: &dyn Grid,
    lookup: &Value,
    table: ResolvedRange,
) -> Option<i32> {
    let len = (table.row_end - table.row_start + 1) as usize;
    if len == 0 {
        return None;
    }

    // Fast path: numeric-only contiguous column slice (blanks are NaN).
    let lookup_num = match lookup {
        Value::Number(n) if !n.is_nan() => Some(*n),
        Value::Empty | Value::Missing => Some(0.0),
        _ => None,
    };
    if let Some(lookup_num) = lookup_num {
        if let Some(slice) =
            grid.column_slice_strict_numeric(table.col_start, table.row_start, table.row_end)
        {
            let mut lo = 0usize;
            let mut hi = slice.len();
            while lo < hi {
                let mid = lo + (hi - lo) / 2;
                let key = slice[mid];
                let key = if key.is_nan() { 0.0 } else { key };
                if key <= lookup_num {
                    lo = mid + 1;
                } else {
                    hi = mid;
                }
            }
            return lo.checked_sub(1).map(|idx| idx as i32);
        }
    }

    // General path: Excel-style compare semantics over cell values.
    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = grid.get_value(CellCoord {
            row: table.row_start + mid as i32,
            col: table.col_start,
        });
        if excel_le(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }
    lo.checked_sub(1).map(|idx| idx as i32)
}

fn approximate_match_in_first_col_array(lookup: &Value, table: &ArrayValue) -> Option<i32> {
    let len = table.rows;
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(mid, 0).unwrap_or(&Value::Empty);
        if excel_le(key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as i32)
}

fn approximate_match_in_first_row(
    grid: &dyn Grid,
    lookup: &Value,
    table: ResolvedRange,
) -> Option<i32> {
    let len = (table.col_end - table.col_start + 1) as usize;
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = grid.get_value(CellCoord {
            row: table.row_start,
            col: table.col_start + mid as i32,
        });
        if excel_le(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }
    lo.checked_sub(1).map(|idx| idx as i32)
}

fn approximate_match_in_first_row_array(lookup: &Value, table: &ArrayValue) -> Option<i32> {
    let len = table.cols;
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(0, mid).unwrap_or(&Value::Empty);
        if excel_le(key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as i32)
}

fn exact_match_1d(lookup: &Value, len: usize, value_at: &dyn Fn(usize) -> Value) -> Option<usize> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for idx in 0..len {
        let candidate = value_at(idx);
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(&candidate) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx);
            }
        } else if values_equal_for_lookup(lookup, &candidate) {
            return Some(idx);
        }
    }
    None
}

fn exact_match_1d_slice(lookup: &Value, values: &[Value]) -> Option<usize> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for (idx, candidate) in values.iter().enumerate() {
        if let Some(pattern) = &wildcard_pattern {
            let text = match coerce_to_string_for_lookup(candidate) {
                Ok(s) => s,
                Err(_) => continue,
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx);
            }
        } else if values_equal_for_lookup(lookup, candidate) {
            return Some(idx);
        }
    }
    None
}

fn approximate_match_1d(
    lookup: &Value,
    len: usize,
    value_at: &dyn Fn(usize) -> Value,
    ascending: bool,
) -> Option<usize> {
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let v = value_at(mid);
        let ok = if ascending {
            excel_le(&v, lookup)?
        } else {
            excel_ge(&v, lookup)?
        };
        if ok {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1)
}

fn approximate_match_1d_slice(lookup: &Value, values: &[Value], ascending: bool) -> Option<usize> {
    let len = values.len();
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let v = &values[mid];
        let ok = if ascending {
            excel_le(v, lookup)?
        } else {
            excel_ge(v, lookup)?
        };
        if ok {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1)
}

enum RangeArg<'a> {
    Range(RangeRef),
    MultiRange(&'a super::value::MultiRangeRef),
    Array(&'a ArrayValue),
}

#[derive(Clone, Copy, Debug)]
enum Range2DArg<'a> {
    Range(ResolvedRange),
    Array(&'a ArrayValue),
}

impl<'a> Range2DArg<'a> {
    fn rows(self) -> i32 {
        match self {
            Range2DArg::Range(r) => r.rows(),
            Range2DArg::Array(a) => i32::try_from(a.rows).unwrap_or(i32::MAX),
        }
    }

    fn cols(self) -> i32 {
        match self {
            Range2DArg::Range(r) => r.cols(),
            Range2DArg::Array(a) => i32::try_from(a.cols).unwrap_or(i32::MAX),
        }
    }

    fn get_value_at(self, grid: &dyn Grid, row_off: i32, col_off: i32) -> Cow<'a, Value> {
        match self {
            Range2DArg::Range(r) => {
                let row = r.row_start + row_off;
                let col = r.col_start + col_off;
                Cow::Owned(grid.get_value(CellCoord { row, col }))
            }
            Range2DArg::Array(a) => {
                let row = row_off as usize;
                let col = col_off as usize;
                let idx = row * a.cols + col;
                match a.values.get(idx) {
                    Some(v) => Cow::Borrowed(v),
                    None => Cow::Owned(Value::Empty),
                }
            }
        }
    }
}

fn range2d_from_value<'a>(
    v: &'a Value,
    grid: &dyn Grid,
    base: CellCoord,
) -> Result<Range2DArg<'a>, ErrorKind> {
    match v {
        Value::Range(r) => {
            let range = r.resolve(base);
            if !range_in_bounds(grid, range) {
                return Err(ErrorKind::Ref);
            }
            grid.record_reference(
                grid.sheet_id(),
                CellCoord {
                    row: range.row_start,
                    col: range.col_start,
                },
                CellCoord {
                    row: range.row_end,
                    col: range.col_end,
                },
            );
            Ok(Range2DArg::Range(range))
        }
        Value::Array(a) => Ok(Range2DArg::Array(a)),
        Value::Error(e) => Err(*e),
        Value::MultiRange(_) => Err(ErrorKind::Value),
        _ => Err(ErrorKind::Value),
    }
}
fn bytecode_value_to_engine(value: Value) -> EngineValue {
    match value {
        Value::Number(n) => EngineValue::Number(n),
        Value::Bool(b) => EngineValue::Bool(b),
        Value::Text(s) => EngineValue::Text(s.to_string()),
        Value::Entity(v) => match Arc::try_unwrap(v) {
            Ok(entity) => EngineValue::Entity(entity),
            Err(shared) => EngineValue::Entity(shared.as_ref().clone()),
        },
        Value::Record(v) => match Arc::try_unwrap(v) {
            Ok(record) => EngineValue::Record(record),
            Err(shared) => EngineValue::Record(shared.as_ref().clone()),
        },
        Value::Empty | Value::Missing => EngineValue::Blank,
        Value::Error(e) => EngineValue::Error(e.into()),
        // Lambdas are not valid scalar values in criteria matching contexts.
        Value::Lambda(_) => EngineValue::Error(EngineErrorKind::Value),
        // Array/range values are not valid scalar values, but the bytecode runtime uses `#SPILL!`
        // for "range-as-scalar" cases elsewhere.
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {
            EngineValue::Error(EngineErrorKind::Spill)
        }
    }
}

fn bytecode_value_to_engine_ref(value: &Value) -> EngineValue {
    match value {
        Value::Number(n) => EngineValue::Number(*n),
        Value::Bool(b) => EngineValue::Bool(*b),
        Value::Text(s) => EngineValue::Text(s.to_string()),
        Value::Entity(v) => EngineValue::Entity(v.as_ref().clone()),
        Value::Record(v) => EngineValue::Record(v.as_ref().clone()),
        Value::Empty | Value::Missing => EngineValue::Blank,
        Value::Error(e) => EngineValue::Error((*e).into()),
        Value::Lambda(_) => EngineValue::Error(EngineErrorKind::Value),
        Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {
            EngineValue::Error(EngineErrorKind::Spill)
        }
    }
}

fn parse_countif_criteria(
    criteria: &Value,
    locale: &crate::LocaleConfig,
) -> Result<EngineCriteria, ErrorKind> {
    // Errors in the criteria argument always propagate (they don't act as "match error" criteria).
    if let Value::Error(e) = criteria {
        return Err(*e);
    }

    let criteria_value = match criteria {
        Value::Number(_)
        | Value::Bool(_)
        | Value::Text(_)
        | Value::Entity(_)
        | Value::Record(_)
        | Value::Empty
        | Value::Missing => bytecode_value_to_engine_ref(criteria),
        Value::Error(_) => unreachable!("handled above"),
        Value::Lambda(_) | Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {
            return Err(ErrorKind::Value);
        }
    };

    EngineCriteria::parse_with_date_system_and_locales(
        &criteria_value,
        thread_date_system(),
        thread_value_locale(),
        thread_now_utc(),
        locale.clone(),
    )
    .map_err(ErrorKind::from)
}

fn count_if_range_criteria(
    grid: &dyn Grid,
    range: ResolvedRange,
    criteria: &EngineCriteria,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut count = 0usize;
            let mut seen = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                seen += 1;
                let engine_value = bytecode_value_to_engine(v);
                if criteria.matches(&engine_value) {
                    count += 1;
                }
            }

            if criteria.matches(&EngineValue::Blank) {
                let total_cells = (range.rows() as i64) * (range.cols() as i64);
                let implicit_blanks = total_cells.saturating_sub(seen as i64);
                count = count.saturating_add(implicit_blanks as usize);
            }

            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        for row in range.row_start..=range.row_end {
            let engine_value = bytecode_value_to_engine(grid.get_value(CellCoord { row, col }));
            if criteria.matches(&engine_value) {
                count += 1;
            }
        }
    }
    Ok(count)
}

fn count_if_range_criteria_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
    criteria: &EngineCriteria,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut count = 0usize;
            let mut seen = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                seen += 1;
                let engine_value = bytecode_value_to_engine(v);
                if criteria.matches(&engine_value) {
                    count += 1;
                }
            }

            if criteria.matches(&EngineValue::Blank) {
                let total_cells = (range.rows() as i64) * (range.cols() as i64);
                let implicit_blanks = total_cells.saturating_sub(seen as i64);
                count = count.saturating_add(implicit_blanks as usize);
            }

            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        for row in range.row_start..=range.row_end {
            let engine_value =
                bytecode_value_to_engine(grid.get_value_on_sheet(sheet, CellCoord { row, col }));
            if criteria.matches(&engine_value) {
                count += 1;
            }
        }
    }
    Ok(count)
}

fn count_if_array_criteria(arr: &ArrayValue, criteria: &EngineCriteria) -> usize {
    // Avoid cloning the bytecode value just to convert it into an `EngineValue` for criteria
    // matching. In particular, cloning `Value::Array` would deep-clone the entire nested array
    // even though it is coerced to `#SPILL!` for COUNTIF matching.
    arr.iter()
        .filter(|&v| criteria.matches(&bytecode_value_to_engine_ref(v)))
        .count()
}

fn count_if_array_numeric_criteria(arr: &ArrayValue, criteria: NumericCriteria) -> usize {
    if arr.len() < SIMD_ARRAY_MIN_LEN {
        return arr
            .iter()
            .filter(|v| {
                let Some(n) = coerce_countif_value_to_number(*v) else {
                    return false;
                };
                matches_numeric_criteria(n, criteria)
            })
            .count();
    }

    let mut count = 0usize;
    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut len = 0usize;

    for v in arr.iter() {
        let Some(n) = coerce_countif_value_to_number(v) else {
            continue;
        };
        if n.is_nan() {
            if matches_numeric_criteria(n, criteria) {
                count += 1;
            }
            continue;
        }

        buf[len] = n;
        len += 1;
        if len == SIMD_AGGREGATE_BLOCK {
            count += simd::count_if_f64(&buf, criteria);
            len = 0;
        }
    }

    if len > 0 {
        count += simd::count_if_f64(&buf[..len], criteria);
    }

    count
}

fn sum_if_array_numeric_criteria(
    values: &ArrayValue,
    criteria_values: &ArrayValue,
    criteria: NumericCriteria,
) -> Option<f64> {
    if values.len() != criteria_values.len() || values.len() < SIMD_ARRAY_MIN_LEN {
        return None;
    }

    let mut sum = 0.0;
    let mut values_buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut criteria_buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut len = 0usize;

    for (crit_v, sum_v) in criteria_values.iter().zip(values.iter()) {
        let Some(n) = coerce_countif_value_to_number(crit_v) else {
            return None;
        };
        // Preserve scalar semantics for NaN numeric values by falling back. The SIMD criteria kernels
        // treat NaNs as blanks.
        if n.is_nan() {
            return None;
        }
        criteria_buf[len] = n;

        match sum_v {
            Value::Number(x) => {
                if x.is_nan() {
                    return None;
                }
                values_buf[len] = *x;
            }
            // Errors/lambdas in the value range must be able to short-circuit when criteria matches.
            Value::Error(_) | Value::Lambda(_) => return None,
            _ => values_buf[len] = f64::NAN,
        }

        len += 1;
        if len == SIMD_AGGREGATE_BLOCK {
            sum += simd::sum_if_f64(&values_buf, &criteria_buf, criteria);
            len = 0;
        }
    }

    if len > 0 {
        sum += simd::sum_if_f64(&values_buf[..len], &criteria_buf[..len], criteria);
    }
    Some(sum)
}

fn sum_count_if_array_numeric_criteria(
    values: &ArrayValue,
    criteria_values: &ArrayValue,
    criteria: NumericCriteria,
) -> Option<(f64, usize)> {
    if values.len() != criteria_values.len() || values.len() < SIMD_ARRAY_MIN_LEN {
        return None;
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    let mut values_buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut criteria_buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut len = 0usize;

    for (crit_v, avg_v) in criteria_values.iter().zip(values.iter()) {
        let Some(n) = coerce_countif_value_to_number(crit_v) else {
            return None;
        };
        if n.is_nan() {
            return None;
        }
        criteria_buf[len] = n;

        match avg_v {
            Value::Number(x) => {
                if x.is_nan() {
                    return None;
                }
                values_buf[len] = *x;
            }
            Value::Error(_) | Value::Lambda(_) => return None,
            _ => values_buf[len] = f64::NAN,
        }

        len += 1;
        if len == SIMD_AGGREGATE_BLOCK {
            let (s, c) = simd::sum_count_if_f64(&values_buf, &criteria_buf, criteria);
            sum += s;
            count += c;
            len = 0;
        }
    }

    if len > 0 {
        let (s, c) = simd::sum_count_if_f64(&values_buf[..len], &criteria_buf[..len], criteria);
        sum += s;
        count += c;
    }
    Some((sum, count))
}

#[inline]
fn range_in_bounds(grid: &dyn Grid, range: ResolvedRange) -> bool {
    grid.in_bounds(CellCoord {
        row: range.row_start,
        col: range.col_start,
    }) && grid.in_bounds(CellCoord {
        row: range.row_end,
        col: range.col_end,
    })
}

#[inline]
fn coord_in_range(coord: CellCoord, range: ResolvedRange) -> bool {
    coord.row >= range.row_start
        && coord.row <= range.row_end
        && coord.col >= range.col_start
        && coord.col <= range.col_end
}

#[inline]
fn record_error_row_major(
    best: &mut Option<(i32, i32, ErrorKind)>,
    coord: CellCoord,
    err: ErrorKind,
) {
    // Preserve row-major error precedence for range scans: rows outermost, columns innermost.
    match best {
        None => *best = Some((coord.row, coord.col, err)),
        Some((best_row, best_col, _)) => {
            if (coord.row, coord.col) < (*best_row, *best_col) {
                *best = Some((coord.row, coord.col, err));
            }
        }
    }
}

#[inline]
fn record_error_sumproduct_offset(
    best: &mut Option<(i32, i32, u8, ErrorKind)>,
    row_off: i32,
    col_off: i32,
    source_priority: u8,
    err: ErrorKind,
) {
    // Preserve Excel/AST-style row-major error precedence for SUMPRODUCT: offsets are ordered by
    // (row, col), and within each element we evaluate argument A before argument B.
    match best {
        None => *best = Some((row_off, col_off, source_priority, err)),
        Some((best_row, best_col, best_src, _)) => {
            if (row_off, col_off, source_priority) < (*best_row, *best_col, *best_src) {
                *best = Some((row_off, col_off, source_priority, err));
            }
        }
    }
}

#[inline]
fn range_in_bounds_on_sheet(grid: &dyn Grid, sheet: &SheetId, range: ResolvedRange) -> bool {
    grid.in_bounds_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
    ) && grid.in_bounds_on_sheet(
        sheet,
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    )
}

fn sum_range(grid: &dyn Grid, range: ResolvedRange) -> Result<f64, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut sum = 0.0;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    // SUM ignores text/logicals/blanks in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Err(err);
            }
            return Ok(sum);
        }
    }

    let mut sum = 0.0;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            sum += simd::sum_ignore_nan_f64(slice);
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value(CellCoord { row, col }) {
                Value::Number(v) => sum += v,
                Value::Error(e) => return Err(e),
                // SUM ignores text/logicals/blanks in references.
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(sum)
}
fn sum_range_on_sheet_with_coord(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<f64, (CellCoord, ErrorKind)> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err((
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            ErrorKind::Ref,
        ));
    }

    grid.record_reference_on_sheet(
        sheet,
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut sum = 0.0;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    // SUM ignores text/logicals/blanks in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((row, col, err)) = best_error {
                return Err((CellCoord { row, col }, err));
            }
            return Ok(sum);
        }
    }

    let mut sum = 0.0;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_on_sheet(sheet, col, range.row_start, range.row_end)
        {
            sum += simd::sum_ignore_nan_f64(slice);
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Number(v) => sum += v,
                Value::Error(e) => return Err((CellCoord { row, col }, e)),
                // SUM ignores text/logicals/blanks in references.
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(sum)
}

fn sum_count_range(grid: &dyn Grid, range: ResolvedRange) -> Result<(f64, usize), ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut sum = 0.0;
            let mut count = 0usize;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    // Ignore non-numeric values in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Err(err);
            }
            return Ok((sum, count));
        }
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            let (s, c) = simd::sum_count_ignore_nan_f64(slice);
            sum += s;
            count += c;
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value(CellCoord { row, col }) {
                Value::Number(v) => {
                    sum += v;
                    count += 1;
                }
                Value::Error(e) => return Err(e),
                // Ignore non-numeric values in references.
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok((sum, count))
}

fn sum_count_range_on_sheet_with_coord(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<(f64, usize), (CellCoord, ErrorKind)> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err((
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            ErrorKind::Ref,
        ));
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut sum = 0.0;
            let mut count = 0usize;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    // Ignore non-numeric values in references.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((row, col, err)) = best_error {
                return Err((CellCoord { row, col }, err));
            }
            return Ok((sum, count));
        }
    }

    let mut sum = 0.0;
    let mut count = 0usize;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_on_sheet(sheet, col, range.row_start, range.row_end)
        {
            let (s, c) = simd::sum_count_ignore_nan_f64(slice);
            sum += s;
            count += c;
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Number(v) => {
                    sum += v;
                    count += 1;
                }
                Value::Error(e) => return Err((CellCoord { row, col }, e)),
                // Ignore non-numeric values in references.
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok((sum, count))
}

fn count_range(grid: &dyn Grid, range: ResolvedRange) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut count = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if matches!(v, Value::Number(_)) {
                    count += 1;
                }
            }
            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            count += simd::count_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                if matches!(grid.get_value(CellCoord { row, col }), Value::Number(_)) {
                    count += 1
                }
            }
        }
    }
    Ok(count)
}

fn count_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err(ErrorKind::Ref);
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut count = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if matches!(v, Value::Number(_)) {
                    count += 1;
                }
            }
            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_on_sheet(sheet, col, range.row_start, range.row_end)
        {
            count += simd::count_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                if matches!(
                    grid.get_value_on_sheet(sheet, CellCoord { row, col }),
                    Value::Number(_)
                ) {
                    count += 1
                }
            }
        }
    }
    Ok(count)
}

fn counta_range(grid: &dyn Grid, range: ResolvedRange) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut count = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if !matches!(v, Value::Empty | Value::Missing) {
                    count += 1;
                }
            }
            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        // Fast path: numeric-only columns (NaN = blank/non-numeric). This is only correct when
        // we validate that the slice contains no non-numeric values. (Non-numeric cells that
        // would be indistinguishable from blanks in the slice force a fallback scan.)
        if let Some(slice) = grid.column_slice_strict_numeric(col, range.row_start, range.row_end) {
            count += simd::count_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                if !matches!(
                    grid.get_value(CellCoord { row, col }),
                    Value::Empty | Value::Missing
                ) {
                    count += 1;
                }
            }
        }
    }
    Ok(count)
}

fn counta_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err(ErrorKind::Ref);
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut count = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if !matches!(v, Value::Empty | Value::Missing) {
                    count += 1;
                }
            }
            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        // Same strict-numeric slice requirement as `counta_range`.
        if let Some(slice) =
            grid.column_slice_on_sheet_strict_numeric(sheet, col, range.row_start, range.row_end)
        {
            count += simd::count_ignore_nan_f64(slice);
        } else {
            for row in range.row_start..=range.row_end {
                if !matches!(
                    grid.get_value_on_sheet(sheet, CellCoord { row, col }),
                    Value::Empty | Value::Missing
                ) {
                    count += 1;
                }
            }
        }
    }
    Ok(count)
}

fn countblank_range(grid: &dyn Grid, range: ResolvedRange) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    let size = (range.rows() as u64).saturating_mul(range.cols() as u64);

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut non_blank = 0u64;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if !matches!(v, Value::Empty | Value::Missing)
                    && !matches!(v, Value::Text(ref s) if s.is_empty())
                    && !matches!(v, Value::Entity(ref ent) if ent.display.is_empty())
                    && !matches!(v, Value::Record(ref rec) if rec.display.is_empty())
                {
                    non_blank += 1;
                }
            }
            return Ok(size.saturating_sub(non_blank) as usize);
        }
    }

    let mut non_blank = 0u64;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_strict_numeric(col, range.row_start, range.row_end) {
            non_blank += simd::count_ignore_nan_f64(slice) as u64;
        } else {
            for row in range.row_start..=range.row_end {
                let v = grid.get_value(CellCoord { row, col });
                if !matches!(v, Value::Empty | Value::Missing)
                    && !matches!(v, Value::Text(ref s) if s.is_empty())
                    && !matches!(v, Value::Entity(ref ent) if ent.display.is_empty())
                    && !matches!(v, Value::Record(ref rec) if rec.display.is_empty())
                {
                    non_blank += 1;
                }
            }
        }
    }

    Ok(size.saturating_sub(non_blank) as usize)
}

fn countblank_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err(ErrorKind::Ref);
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    let size = (range.rows() as u64).saturating_mul(range.cols() as u64);

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut non_blank = 0u64;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                if !matches!(v, Value::Empty | Value::Missing)
                    && !matches!(v, Value::Text(ref s) if s.is_empty())
                    && !matches!(v, Value::Entity(ref ent) if ent.display.is_empty())
                    && !matches!(v, Value::Record(ref rec) if rec.display.is_empty())
                {
                    non_blank += 1;
                }
            }
            return Ok(size.saturating_sub(non_blank) as usize);
        }
    }

    let mut non_blank = 0u64;
    for col in range.col_start..=range.col_end {
        if let Some(slice) =
            grid.column_slice_on_sheet_strict_numeric(sheet, col, range.row_start, range.row_end)
        {
            non_blank += simd::count_ignore_nan_f64(slice) as u64;
        } else {
            for row in range.row_start..=range.row_end {
                let v = grid.get_value_on_sheet(sheet, CellCoord { row, col });
                if !matches!(v, Value::Empty | Value::Missing)
                    && !matches!(v, Value::Text(ref s) if s.is_empty())
                    && !matches!(v, Value::Entity(ref ent) if ent.display.is_empty())
                    && !matches!(v, Value::Record(ref rec) if rec.display.is_empty())
                {
                    non_blank += 1;
                }
            }
        }
    }

    Ok(size.saturating_sub(non_blank) as usize)
}

fn min_range(grid: &dyn Grid, range: ResolvedRange) -> Result<Option<f64>, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut out: Option<f64> = None;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => out = Some(out.map_or(n, |prev| prev.min(n))),
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Err(err);
            }
            return Ok(out);
        }
    }

    let mut out: Option<f64> = None;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            if let Some(m) = simd::min_ignore_nan_f64(slice) {
                out = Some(out.map_or(m, |prev| prev.min(m)));
            }
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value(CellCoord { row, col }) {
                Value::Number(v) => out = Some(out.map_or(v, |prev| prev.min(v))),
                Value::Error(e) => return Err(e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(out)
}

fn min_range_on_sheet_with_coord(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<Option<f64>, (CellCoord, ErrorKind)> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err((
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            ErrorKind::Ref,
        ));
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut out: Option<f64> = None;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => out = Some(out.map_or(n, |prev| prev.min(n))),
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((row, col, err)) = best_error {
                return Err((CellCoord { row, col }, err));
            }
            return Ok(out);
        }
    }

    let mut out: Option<f64> = None;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_on_sheet(sheet, col, range.row_start, range.row_end)
        {
            if let Some(m) = simd::min_ignore_nan_f64(slice) {
                out = Some(out.map_or(m, |prev| prev.min(m)));
            }
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Number(v) => out = Some(out.map_or(v, |prev| prev.min(v))),
                Value::Error(e) => return Err((CellCoord { row, col }, e)),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(out)
}

fn max_range(grid: &dyn Grid, range: ResolvedRange) -> Result<Option<f64>, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut out: Option<f64> = None;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => out = Some(out.map_or(n, |prev| prev.max(n))),
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((_, _, err)) = best_error {
                return Err(err);
            }
            return Ok(out);
        }
    }

    let mut out: Option<f64> = None;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice(col, range.row_start, range.row_end) {
            if let Some(m) = simd::max_ignore_nan_f64(slice) {
                out = Some(out.map_or(m, |prev| prev.max(m)));
            }
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value(CellCoord { row, col }) {
                Value::Number(v) => out = Some(out.map_or(v, |prev| prev.max(v))),
                Value::Error(e) => return Err(e),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(out)
}

fn max_range_on_sheet_with_coord(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
) -> Result<Option<f64>, (CellCoord, ErrorKind)> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err((
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            ErrorKind::Ref,
        ));
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut out: Option<f64> = None;
            let mut best_error: Option<(i32, i32, ErrorKind)> = None;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                match v {
                    Value::Number(n) => out = Some(out.map_or(n, |prev| prev.max(n))),
                    Value::Error(e) => record_error_row_major(&mut best_error, coord, e),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Empty
                    | Value::Missing
                    | Value::Array(_)
                    | Value::Range(_)
                    | Value::MultiRange(_)
                    | Value::Lambda(_) => {}
                }
            }
            if let Some((row, col, err)) = best_error {
                return Err((CellCoord { row, col }, err));
            }
            return Ok(out);
        }
    }

    let mut out: Option<f64> = None;
    let mut scan_cols: Vec<i32> = Vec::new();
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_on_sheet(sheet, col, range.row_start, range.row_end)
        {
            if let Some(m) = simd::max_ignore_nan_f64(slice) {
                out = Some(out.map_or(m, |prev| prev.max(m)));
            }
        } else {
            scan_cols.push(col);
        }
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    for row in range.row_start..=range.row_end {
        for &col in &scan_cols {
            match grid.get_value_on_sheet(sheet, CellCoord { row, col }) {
                Value::Number(v) => out = Some(out.map_or(v, |prev| prev.max(v))),
                Value::Error(e) => return Err((CellCoord { row, col }, e)),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Empty
                | Value::Missing
                | Value::Array(_)
                | Value::Range(_)
                | Value::MultiRange(_)
                | Value::Lambda(_) => {}
            }
        }
    }
    Ok(out)
}

fn count_if_range(
    grid: &dyn Grid,
    range: ResolvedRange,
    criteria: NumericCriteria,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds(grid, range) {
        return Err(ErrorKind::Ref);
    }

    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: range.row_start,
            col: range.col_start,
        },
        CellCoord {
            row: range.row_end,
            col: range.col_end,
        },
    );

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells() {
            let mut count = 0usize;
            let mut seen = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                seen += 1;
                if let Some(n) = coerce_countif_value_to_number(&v) {
                    if matches_numeric_criteria(n, criteria) {
                        count += 1;
                    }
                }
            }

            // COUNTIF treats implicit blanks as zero for numeric criteria.
            if matches_numeric_criteria(0.0, criteria) {
                let total_cells = (range.rows() as i64) * (range.cols() as i64);
                let implicit_blanks = total_cells.saturating_sub(seen as i64);
                count = count.saturating_add(implicit_blanks as usize);
            }

            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) = grid.column_slice_strict_numeric(col, range.row_start, range.row_end) {
            count += simd::count_if_blank_as_zero_f64(slice, criteria);
        } else {
            for row in range.row_start..=range.row_end {
                if let Some(v) =
                    coerce_countif_value_to_number(&grid.get_value(CellCoord { row, col }))
                {
                    if matches_numeric_criteria(v, criteria) {
                        count += 1;
                    }
                }
            }
        }
    }
    Ok(count)
}

fn count_if_range_on_sheet(
    grid: &dyn Grid,
    sheet: &SheetId,
    range: ResolvedRange,
    criteria: NumericCriteria,
) -> Result<usize, ErrorKind> {
    if !range_in_bounds_on_sheet(grid, sheet, range) {
        return Err(ErrorKind::Ref);
    }

    if let SheetId::Local(sheet_id) = sheet {
        grid.record_reference(
            *sheet_id,
            CellCoord {
                row: range.row_start,
                col: range.col_start,
            },
            CellCoord {
                row: range.row_end,
                col: range.col_end,
            },
        );
    }

    if range_should_iterate_sparse(range) {
        if let Some(iter) = grid.iter_cells_on_sheet(sheet) {
            let mut count = 0usize;
            let mut seen = 0usize;
            for (coord, v) in iter {
                if !coord_in_range(coord, range) {
                    continue;
                }
                seen += 1;
                if let Some(n) = coerce_countif_value_to_number(&v) {
                    if matches_numeric_criteria(n, criteria) {
                        count += 1;
                    }
                }
            }

            // COUNTIF treats implicit blanks as zero for numeric criteria.
            if matches_numeric_criteria(0.0, criteria) {
                let total_cells = (range.rows() as i64) * (range.cols() as i64);
                let implicit_blanks = total_cells.saturating_sub(seen as i64);
                count = count.saturating_add(implicit_blanks as usize);
            }

            return Ok(count);
        }
    }

    let mut count = 0usize;
    for col in range.col_start..=range.col_end {
        if let Some(slice) =
            grid.column_slice_on_sheet_strict_numeric(sheet, col, range.row_start, range.row_end)
        {
            count += simd::count_if_blank_as_zero_f64(slice, criteria);
        } else {
            for row in range.row_start..=range.row_end {
                if let Some(v) = coerce_countif_value_to_number(
                    &grid.get_value_on_sheet(sheet, CellCoord { row, col }),
                ) {
                    if matches_numeric_criteria(v, criteria) {
                        count += 1;
                    }
                }
            }
        }
    }
    Ok(count)
}

fn coerce_sumproduct_number(v: &Value) -> Result<f64, ErrorKind> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Text(s) => match parse_number(s, thread_number_locale()) {
            Ok(n) => Ok(n),
            Err(ExcelError::Value) => Ok(0.0),
            Err(ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(ExcelError::Num) => Err(ErrorKind::Num),
        },
        // Rich values behave like text for numeric coercions.
        Value::Entity(entity) => match parse_number(&entity.display, thread_number_locale()) {
            Ok(n) => Ok(n),
            Err(ExcelError::Value) => Ok(0.0),
            Err(ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(ExcelError::Num) => Err(ErrorKind::Num),
        },
        Value::Record(record) => match parse_number(&record.display, thread_number_locale()) {
            Ok(n) => Ok(n),
            Err(ExcelError::Value) => Ok(0.0),
            Err(ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(ExcelError::Num) => Err(ErrorKind::Num),
        },
        Value::Empty | Value::Missing => Ok(0.0),
        Value::Error(e) => Err(*e),
        Value::Lambda(_) | Value::Array(_) | Value::Range(_) | Value::MultiRange(_) => {
            Err(ErrorKind::Value)
        }
    }
}

fn sumproduct_range(grid: &dyn Grid, a: ResolvedRange, b: ResolvedRange) -> Result<f64, ErrorKind> {
    if !range_in_bounds(grid, a) || !range_in_bounds(grid, b) {
        return Err(ErrorKind::Ref);
    }
    if a.rows() != b.rows() || a.cols() != b.cols() {
        return Err(ErrorKind::Value);
    }

    // SUMPRODUCT reads values from *both* ranges; record both rectangles once so dynamic dependency
    // tracing stays compact (no per-cell events).
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: a.row_start,
            col: a.col_start,
        },
        CellCoord {
            row: a.row_end,
            col: a.col_end,
        },
    );
    grid.record_reference(
        grid.sheet_id(),
        CellCoord {
            row: b.row_start,
            col: b.col_start,
        },
        CellCoord {
            row: b.row_end,
            col: b.col_end,
        },
    );

    let rows = a.rows();
    let cols = a.cols();

    if range_should_iterate_sparse(a) {
        if let Some(iter) = grid.iter_cells() {
            let mut sum = 0.0;
            let mut best_error: Option<(i32, i32, u8, ErrorKind)> = None;
            let mut seen_offsets: HashSet<i64> = HashSet::new();
            let cols_i64 = cols as i64;

            for (coord, v) in iter {
                let in_a = coord_in_range(coord, a);
                let in_b = coord_in_range(coord, b);

                match (in_a, in_b) {
                    (false, false) => {}
                    (true, false) => {
                        let row_off = coord.row - a.row_start;
                        let col_off = coord.col - a.col_start;
                        let key = (row_off as i64) * cols_i64 + (col_off as i64);
                        if !seen_offsets.insert(key) {
                            continue;
                        }

                        let x = match coerce_sumproduct_number(&v) {
                            Ok(n) => n,
                            Err(e) => {
                                record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off,
                                    col_off,
                                    0,
                                    e,
                                );
                                continue;
                            }
                        };
                        let rb = CellCoord {
                            row: b.row_start + row_off,
                            col: b.col_start + col_off,
                        };
                        let rb_value = grid.get_value(rb);
                        let y = match coerce_sumproduct_number(&rb_value) {
                            Ok(n) => n,
                            Err(e) => {
                                record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off,
                                    col_off,
                                    1,
                                    e,
                                );
                                continue;
                            }
                        };
                        sum += x * y;
                    }
                    (false, true) => {
                        let row_off = coord.row - b.row_start;
                        let col_off = coord.col - b.col_start;
                        let key = (row_off as i64) * cols_i64 + (col_off as i64);
                        if !seen_offsets.insert(key) {
                            continue;
                        }

                        let ra = CellCoord {
                            row: a.row_start + row_off,
                            col: a.col_start + col_off,
                        };
                        let ra_value = grid.get_value(ra);
                        let x = match coerce_sumproduct_number(&ra_value) {
                            Ok(n) => n,
                            Err(e) => {
                                record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off,
                                    col_off,
                                    0,
                                    e,
                                );
                                continue;
                            }
                        };
                        let y = match coerce_sumproduct_number(&v) {
                            Ok(n) => n,
                            Err(e) => {
                                record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off,
                                    col_off,
                                    1,
                                    e,
                                );
                                continue;
                            }
                        };
                        sum += x * y;
                    }
                    (true, true) => {
                        // Overlap: the same stored cell may correspond to *two* different SUMPRODUCT
                        // element offsets (once as A, once as B).
                        let v_num = coerce_sumproduct_number(&v);

                        // As A
                        let row_off_a = coord.row - a.row_start;
                        let col_off_a = coord.col - a.col_start;
                        let key_a = (row_off_a as i64) * cols_i64 + (col_off_a as i64);
                        if seen_offsets.insert(key_a) {
                            match &v_num {
                                Ok(x) => {
                                    let x = *x;
                                    let rb = CellCoord {
                                        row: b.row_start + row_off_a,
                                        col: b.col_start + col_off_a,
                                    };
                                    let rb_value = grid.get_value(rb);
                                    match coerce_sumproduct_number(&rb_value) {
                                        Ok(y) => sum += x * y,
                                        Err(e) => record_error_sumproduct_offset(
                                            &mut best_error,
                                            row_off_a,
                                            col_off_a,
                                            1,
                                            e,
                                        ),
                                    }
                                }
                                Err(e) => record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off_a,
                                    col_off_a,
                                    0,
                                    *e,
                                ),
                            }
                        }

                        // As B
                        let row_off_b = coord.row - b.row_start;
                        let col_off_b = coord.col - b.col_start;
                        let key_b = (row_off_b as i64) * cols_i64 + (col_off_b as i64);
                        if seen_offsets.insert(key_b) {
                            let ra = CellCoord {
                                row: a.row_start + row_off_b,
                                col: a.col_start + col_off_b,
                            };
                            let ra_value = grid.get_value(ra);
                            match coerce_sumproduct_number(&ra_value) {
                                Ok(x) => match &v_num {
                                    Ok(y) => sum += x * (*y),
                                    Err(e) => record_error_sumproduct_offset(
                                        &mut best_error,
                                        row_off_b,
                                        col_off_b,
                                        1,
                                        *e,
                                    ),
                                },
                                Err(e) => record_error_sumproduct_offset(
                                    &mut best_error,
                                    row_off_b,
                                    col_off_b,
                                    0,
                                    e,
                                ),
                            }
                        }
                    }
                }
            }

            if let Some((_, _, _, err)) = best_error {
                return Err(err);
            }
            return Ok(sum);
        }
    }

    // For strict numeric slices, SUMPRODUCT can run in SIMD without per-cell reads. Errors and
    // non-numeric values disqualify strict numeric slices, so error precedence only matters on
    // the fallback path below.
    let cols_usize = cols as usize;
    let mut slice_a: Vec<Option<&[f64]>> = Vec::with_capacity(cols_usize);
    let mut slice_b: Vec<Option<&[f64]>> = Vec::with_capacity(cols_usize);
    for col_offset in 0..cols {
        let col_a = a.col_start + col_offset;
        let col_b = b.col_start + col_offset;
        slice_a.push(grid.column_slice_strict_numeric(col_a, a.row_start, a.row_end));
        slice_b.push(grid.column_slice_strict_numeric(col_b, b.row_start, b.row_end));
    }
    let all_slices = slice_a
        .iter()
        .zip(slice_b.iter())
        .all(|(sa, sb)| sa.is_some() && sb.is_some());
    if all_slices {
        let mut sum = 0.0;
        for col_offset in 0..cols_usize {
            let sa = slice_a[col_offset].expect("validated all_slices");
            let sb = slice_b[col_offset].expect("validated all_slices");
            sum += simd::sumproduct_ignore_nan_f64(sa, sb);
        }
        return Ok(sum);
    }

    // Dense fallback: scan in row-major order so error precedence matches the AST evaluator.
    let mut sum = 0.0;
    for row_offset in 0..rows {
        let row_idx = row_offset as usize;
        for col_offset in 0..cols {
            let col_idx = col_offset as usize;
            if let (Some(sa), Some(sb)) = (slice_a[col_idx], slice_b[col_idx]) {
                let mut x = sa[row_idx];
                let mut y = sb[row_idx];
                // Strict-numeric slices represent blanks as NaN; SUMPRODUCT treats blanks as 0.
                if x.is_nan() {
                    x = 0.0;
                }
                if y.is_nan() {
                    y = 0.0;
                }
                sum += x * y;
                continue;
            }

            let col_a = a.col_start + col_offset;
            let col_b = b.col_start + col_offset;
            let ra = CellCoord {
                row: a.row_start + row_offset,
                col: col_a,
            };
            let rb = CellCoord {
                row: b.row_start + row_offset,
                col: col_b,
            };
            let x = coerce_sumproduct_number(&grid.get_value(ra))?;
            let y = coerce_sumproduct_number(&grid.get_value(rb))?;
            sum += x * y;
        }
    }

    Ok(sum)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bytecode::ColumnarGrid;
    use std::collections::HashMap;
    use std::sync::{Arc, Mutex};

    #[derive(Default)]
    struct TracingGrid {
        values: HashMap<(i32, i32), Value>,
        trace: Mutex<Vec<(usize, CellCoord, CellCoord)>>,
    }

    impl Grid for TracingGrid {
        fn get_value(&self, coord: CellCoord) -> Value {
            self.values
                .get(&(coord.row, coord.col))
                .cloned()
                .unwrap_or(Value::Empty)
        }

        fn sheet_id(&self) -> usize {
            0
        }

        fn record_reference(&self, sheet: usize, start: CellCoord, end: CellCoord) {
            self.trace.lock().unwrap().push((sheet, start, end));
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn bounds(&self) -> (i32, i32) {
            (10, 10)
        }
    }

    #[test]
    fn bytecode_dependency_trace_deref_value_dynamic_records_reference() {
        let grid = TracingGrid::default();

        // Program: push a range reference, then return it (Vm::eval will dynamically dereference).
        let mut program = crate::bytecode::Program::new(Arc::from("trace_test"));
        program.range_refs.push(RangeRef::new(
            Ref::new(0, 0, true, true), // A1
            Ref::new(1, 1, true, true), // B2
        ));
        program.instrs.push(crate::bytecode::Instruction::new(
            crate::bytecode::OpCode::LoadRange,
            0,
            0,
        ));

        let mut vm = crate::bytecode::Vm::new();
        let _ = vm.eval(
            &program,
            &grid,
            0,
            CellCoord { row: 0, col: 0 },
            &crate::LocaleConfig::en_us(),
        );

        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![(
                0,
                CellCoord { row: 0, col: 0 },
                CellCoord { row: 1, col: 1 }
            )]
        );
    }

    #[test]
    fn bytecode_dependency_trace_concat_records_reference() {
        let mut grid = TracingGrid::default();
        grid.values.insert((0, 0), Value::Text(Arc::from("a")));
        grid.values.insert((1, 0), Value::Text(Arc::from("b")));

        let range = RangeRef::new(
            Ref::new(0, 0, true, true), // A1
            Ref::new(1, 0, true, true), // A2
        );
        let base = CellCoord::new(0, 0);
        let out = fn_concat(&[Value::Range(range)], &grid, base);

        assert_eq!(out, Value::Text(Arc::from("ab")));
        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![(
                0,
                CellCoord { row: 0, col: 0 },
                CellCoord { row: 1, col: 0 }
            )]
        );
    }

    #[test]
    fn bytecode_dependency_trace_sumifs_records_reference() {
        let mut grid = TracingGrid::default();
        // Criteria range A1:A2
        grid.values.insert((0, 0), Value::Number(1.0));
        grid.values.insert((1, 0), Value::Number(0.0));
        // Sum range B1:B2
        grid.values.insert((0, 1), Value::Number(10.0));
        grid.values.insert((1, 1), Value::Number(20.0));

        let crit_range = RangeRef::new(Ref::new(0, 0, true, true), Ref::new(1, 0, true, true));
        let sum_range = RangeRef::new(Ref::new(0, 1, true, true), Ref::new(1, 1, true, true));

        let base = CellCoord::new(0, 0);
        let locale = crate::LocaleConfig::en_us();
        let out = fn_sumifs(
            &[
                Value::Range(sum_range),
                Value::Range(crit_range),
                Value::Number(1.0),
            ],
            &grid,
            base,
            &locale,
        );

        assert_eq!(out, Value::Number(10.0));
        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![
                (
                    0,
                    CellCoord { row: 0, col: 1 },
                    CellCoord { row: 1, col: 1 }
                ),
                (
                    0,
                    CellCoord { row: 0, col: 0 },
                    CellCoord { row: 1, col: 0 }
                )
            ]
        );
    }

    #[test]
    fn bytecode_dependency_trace_isblank_records_reference() {
        let grid = TracingGrid::default();

        let range = RangeRef::new(
            Ref::new(0, 0, true, true), // A1
            Ref::new(1, 0, true, true), // A2
        );
        let base = CellCoord::new(0, 0);
        let _ = fn_isblank(&[Value::Range(range)], &grid, base);

        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![(
                0,
                CellCoord { row: 0, col: 0 },
                CellCoord { row: 1, col: 0 }
            )]
        );
    }

    #[test]
    fn bytecode_dependency_trace_isblank_multirange_records_reference() {
        let grid = TracingGrid::default();

        let range = RangeRef::new(
            Ref::new(0, 0, true, true), // A1
            Ref::new(1, 0, true, true), // A2
        );
        let area = SheetRangeRef::new(SheetId::Local(3), range);
        let mr = MultiRangeRef::new(Arc::from([area]));
        let base = CellCoord::new(0, 0);
        let _ = fn_isblank(&[Value::MultiRange(mr)], &grid, base);

        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![(
                3,
                CellCoord { row: 0, col: 0 },
                CellCoord { row: 1, col: 0 }
            )]
        );
    }

    #[test]
    fn bytecode_dependency_trace_type_single_cell_records_reference() {
        let mut grid = TracingGrid::default();
        grid.values.insert((0, 0), Value::Number(3.0));

        let range = RangeRef::new(
            Ref::new(0, 0, true, true), // A1
            Ref::new(0, 0, true, true), // A1
        );
        let base = CellCoord::new(0, 0);
        let out = fn_type(&[Value::Range(range)], &grid, base);
        assert_eq!(out, Value::Number(1.0));

        let trace = grid.trace.lock().unwrap().clone();
        assert_eq!(
            trace,
            vec![(
                0,
                CellCoord { row: 0, col: 0 },
                CellCoord { row: 0, col: 0 }
            )]
        );
    }

    struct SparsePanicGrid {
        bounds: (i32, i32),
        cells: Vec<(CellCoord, Value)>,
        cells_by_sheet: std::collections::HashMap<usize, Vec<(CellCoord, Value)>>,
    }

    impl Grid for SparsePanicGrid {
        fn get_value(&self, _coord: CellCoord) -> Value {
            panic!("unexpected get_value call (expected sparse iteration)");
        }

        fn get_value_on_sheet(&self, _sheet: &SheetId, _coord: CellCoord) -> Value {
            panic!("unexpected get_value_on_sheet call (expected sparse iteration)");
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn column_slice_on_sheet(
            &self,
            _sheet: &SheetId,
            _col: i32,
            _row_start: i32,
            _row_end: i32,
        ) -> Option<&[f64]> {
            None
        }

        fn iter_cells(&self) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
            Some(Box::new(self.cells.iter().cloned()))
        }

        fn iter_cells_on_sheet(
            &self,
            sheet: &SheetId,
        ) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
            match sheet {
                SheetId::Local(sheet) => {
                    Some(Box::new(self.cells_by_sheet.get(sheet)?.iter().cloned()))
                }
                SheetId::External(_) => None,
            }
        }

        fn bounds(&self) -> (i32, i32) {
            self.bounds
        }

        fn bounds_on_sheet(&self, _sheet: &SheetId) -> (i32, i32) {
            self.bounds
        }
    }

    struct PanicGetGrid {
        bounds: (i32, i32),
    }

    impl Grid for PanicGetGrid {
        fn get_value(&self, _coord: CellCoord) -> Value {
            panic!("unexpected get_value call (materialization should have been guarded)");
        }

        fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
            None
        }

        fn bounds(&self) -> (i32, i32) {
            self.bounds
        }
    }

    #[test]
    fn bytecode_materialization_limit() {
        // Construct a range whose resolved cell count is just over the engine's materialization
        // limit and ensure the bytecode runtime returns #SPILL! without allocating or visiting
        // individual cells.
        let end_row = i32::try_from(MAX_MATERIALIZED_ARRAY_CELLS).expect("limit fits in i32");
        let range = RangeRef::new(Ref::new(0, 0, true, true), Ref::new(end_row, 0, true, true));

        let grid = PanicGetGrid {
            bounds: (end_row + 1, 1),
        };
        let origin = CellCoord::new(0, 0);
        let value = deref_value_dynamic(Value::Range(range), &grid, origin);
        assert_eq!(value, Value::Error(ErrorKind::Spill));
    }

    #[test]
    fn eval_ast_choose_is_lazy() {
        #[derive(Clone, Copy)]
        struct PanicGrid {
            panic_coord: CellCoord,
        }

        impl Grid for PanicGrid {
            fn get_value(&self, coord: CellCoord) -> Value {
                if coord == self.panic_coord {
                    panic!("unexpected evaluation of cell {coord:?}");
                }
                Value::Empty
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn bounds(&self) -> (i32, i32) {
                (10, 10)
            }
        }

        let origin = CellCoord::new(0, 0);
        let expr = crate::bytecode::parse_formula("=CHOOSE(2, A2, 7)", origin).expect("parse");
        // A2 relative to origin (A1) => (row=1, col=0)
        let grid = PanicGrid {
            panic_coord: CellCoord::new(1, 0),
        };

        let value = eval_ast(&expr, &grid, 0, origin, &crate::LocaleConfig::en_us());
        assert_eq!(value, Value::Number(7.0));
    }

    #[test]
    fn range_aggregates_return_ref_for_out_of_bounds_ranges() {
        let grid = ColumnarGrid::new(10, 10);

        let range = ResolvedRange {
            row_start: 0,
            row_end: 20,
            col_start: 0,
            col_end: 0,
        };

        assert_eq!(sum_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(sum_count_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(count_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(counta_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(countblank_range(&grid, range), Err(ErrorKind::Ref));

        let criteria = NumericCriteria::new(CmpOp::Gt, 0.0);
        assert_eq!(count_if_range(&grid, range, criteria), Err(ErrorKind::Ref));
        assert_eq!(min_range(&grid, range), Err(ErrorKind::Ref));
        assert_eq!(max_range(&grid, range), Err(ErrorKind::Ref));

        assert_eq!(sumproduct_range(&grid, range, range), Err(ErrorKind::Ref));
    }

    #[test]
    fn sumproduct_uses_sparse_iteration_for_huge_ranges_without_column_slices() {
        use std::collections::HashMap;
        use std::sync::atomic::{AtomicUsize, Ordering};

        struct LimitedGetGrid {
            bounds: (i32, i32),
            max_get_calls: usize,
            get_calls: AtomicUsize,
            cells: HashMap<(i32, i32), Value>,
            stored: Vec<(CellCoord, Value)>,
        }

        impl Grid for LimitedGetGrid {
            fn get_value(&self, coord: CellCoord) -> Value {
                let seen = self.get_calls.fetch_add(1, Ordering::Relaxed);
                if seen >= self.max_get_calls {
                    panic!(
                        "too many get_value calls (saw {seen}, max {})",
                        self.max_get_calls
                    );
                }
                self.cells
                    .get(&(coord.row, coord.col))
                    .cloned()
                    .unwrap_or(Value::Empty)
            }

            fn get_value_on_sheet(&self, sheet: &SheetId, coord: CellCoord) -> Value {
                match sheet {
                    SheetId::Local(_) => self.get_value(coord),
                    SheetId::External(_) => Value::Error(ErrorKind::Ref),
                }
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn iter_cells(&self) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                Some(Box::new(self.stored.iter().cloned()))
            }

            fn bounds(&self) -> (i32, i32) {
                self.bounds
            }
        }

        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range_a = RangeRef::new(Ref::new(0, 0, true, true), Ref::new(row_end, 0, true, true));
        let range_b = RangeRef::new(Ref::new(0, 1, true, true), Ref::new(row_end, 1, true, true));

        let mut cells: HashMap<(i32, i32), Value> = HashMap::new();
        let mut stored = Vec::new();
        for (row, col, n) in [(0, 0, 2.0), (0, 1, 3.0), (1, 0, 4.0), (1, 1, 5.0)] {
            let coord = CellCoord { row, col };
            let value = Value::Number(n);
            cells.insert((row, col), value.clone());
            stored.push((coord, value));
        }

        let grid = LimitedGetGrid {
            bounds: (row_end + 1, 2),
            max_get_calls: 1024,
            get_calls: AtomicUsize::new(0),
            cells,
            stored,
        };

        let args = [Value::Range(range_a), Value::Range(range_b)];
        let out = fn_sumproduct(&args, &grid, CellCoord { row: 0, col: 0 });
        assert_eq!(out, Value::Number(26.0));
    }

    #[test]
    fn sumproduct_sparse_error_precedence_is_row_major_for_huge_ranges() {
        use std::collections::HashMap;

        struct SparseErrorGrid {
            bounds: (i32, i32),
            cells: HashMap<(i32, i32), Value>,
            stored: Vec<(CellCoord, Value)>,
        }

        impl Grid for SparseErrorGrid {
            fn get_value(&self, coord: CellCoord) -> Value {
                self.cells
                    .get(&(coord.row, coord.col))
                    .cloned()
                    .unwrap_or(Value::Empty)
            }

            fn get_value_on_sheet(&self, sheet: &SheetId, coord: CellCoord) -> Value {
                match sheet {
                    SheetId::Local(_) => self.get_value(coord),
                    SheetId::External(_) => Value::Error(ErrorKind::Ref),
                }
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn iter_cells(&self) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                Some(Box::new(self.stored.iter().cloned()))
            }

            fn bounds(&self) -> (i32, i32) {
                self.bounds
            }
        }

        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range_a = RangeRef::new(Ref::new(0, 0, true, true), Ref::new(row_end, 1, true, true));
        let range_b = RangeRef::new(Ref::new(0, 2, true, true), Ref::new(row_end, 3, true, true));

        // Two different errors in the first (A) range:
        // - B1 (row 0, col 1) is earlier in row-major order than A2 (row 1, col 0).
        let mut cells: HashMap<(i32, i32), Value> = HashMap::new();
        let mut stored = Vec::new();

        // Insert out of order to ensure we rely on offset-based precedence, not iterator order.
        let a2 = (CellCoord { row: 1, col: 0 }, Value::Error(ErrorKind::Ref));
        let b1 = (CellCoord { row: 0, col: 1 }, Value::Error(ErrorKind::Div0));

        for (coord, value) in [a2, b1] {
            cells.insert((coord.row, coord.col), value.clone());
            stored.push((coord, value));
        }

        let grid = SparseErrorGrid {
            bounds: (row_end + 1, 4),
            cells,
            stored,
        };

        // SUMPRODUCT should return the earliest error in row-major order: B1 => #DIV/0!
        let args = [Value::Range(range_a), Value::Range(range_b)];
        let out = fn_sumproduct(&args, &grid, CellCoord { row: 0, col: 0 });
        assert_eq!(out, Value::Error(ErrorKind::Div0));
    }

    #[test]
    fn and_range_ignores_missing_in_sparse_refs() {
        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let grid = SparsePanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![(CellCoord { row: 0, col: 0 }, Value::Missing)],
            cells_by_sheet: std::collections::HashMap::new(),
        };

        let mut all_true = true;
        let mut any = false;
        assert_eq!(and_range(&grid, range, &mut all_true, &mut any), None);
        assert_eq!(all_true, true);
        assert_eq!(any, false);
    }

    #[test]
    fn or_range_ignores_missing_in_sparse_refs() {
        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let grid = SparsePanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![(CellCoord { row: 0, col: 0 }, Value::Missing)],
            cells_by_sheet: std::collections::HashMap::new(),
        };

        let mut any_true = false;
        let mut any = false;
        assert_eq!(or_range(&grid, range, &mut any_true, &mut any), None);
        assert_eq!(any_true, false);
        assert_eq!(any, false);
    }

    #[test]
    fn and_range_on_sheet_ignores_missing_in_sparse_refs() {
        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let grid = SparsePanicGrid {
            bounds: (row_end + 1, 1),
            cells: Vec::new(),
            cells_by_sheet: std::collections::HashMap::from([(
                0,
                vec![(CellCoord { row: 0, col: 0 }, Value::Missing)],
            )]),
        };

        let mut all_true = true;
        let mut any = false;
        assert_eq!(
            and_range_on_sheet(&grid, &SheetId::Local(0), range, &mut all_true, &mut any),
            None
        );
        assert_eq!(all_true, true);
        assert_eq!(any, false);
    }

    #[test]
    fn or_range_on_sheet_ignores_missing_in_sparse_refs() {
        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let grid = SparsePanicGrid {
            bounds: (row_end + 1, 1),
            cells: Vec::new(),
            cells_by_sheet: std::collections::HashMap::from([(
                0,
                vec![(CellCoord { row: 0, col: 0 }, Value::Missing)],
            )]),
        };

        let mut any_true = false;
        let mut any = false;
        assert_eq!(
            or_range_on_sheet(&grid, &SheetId::Local(0), range, &mut any_true, &mut any),
            None
        );
        assert_eq!(any_true, false);
        assert_eq!(any, false);
    }

    #[test]
    fn counta_and_countblank_use_sparse_iteration_for_sheet_ranges() {
        use std::collections::HashMap;

        struct PanicGrid {
            bounds: (i32, i32),
            cells_by_sheet: HashMap<usize, Vec<(CellCoord, Value)>>,
        }

        impl Grid for PanicGrid {
            fn get_value(&self, _coord: CellCoord) -> Value {
                panic!("unexpected get_value call (expected sparse iteration)");
            }

            fn get_value_on_sheet(&self, _sheet: &SheetId, _coord: CellCoord) -> Value {
                panic!("unexpected get_value_on_sheet call (expected sparse iteration)");
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn column_slice_on_sheet(
                &self,
                _sheet: &SheetId,
                _col: i32,
                _row_start: i32,
                _row_end: i32,
            ) -> Option<&[f64]> {
                None
            }

            fn iter_cells_on_sheet(
                &self,
                sheet: &SheetId,
            ) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                match sheet {
                    SheetId::Local(sheet) => {
                        Some(Box::new(self.cells_by_sheet.get(sheet)?.iter().cloned()))
                    }
                    SheetId::External(_) => None,
                }
            }

            fn bounds(&self) -> (i32, i32) {
                self.bounds
            }

            fn bounds_on_sheet(&self, _sheet: &SheetId) -> (i32, i32) {
                self.bounds
            }
        }

        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let grid = PanicGrid {
            bounds: (row_end + 1, 2),
            cells_by_sheet: HashMap::from([(
                0,
                vec![
                    (CellCoord { row: 0, col: 0 }, Value::Empty),
                    (CellCoord { row: 1, col: 0 }, Value::Number(1.0)),
                    (CellCoord { row: 2, col: 0 }, Value::Text(Arc::from("x"))),
                    // Empty string is non-empty for COUNTA, blank for COUNTBLANK.
                    (CellCoord { row: 3, col: 0 }, Value::Text(Arc::from(""))),
                    // Outside the range (different col).
                    (CellCoord { row: 1, col: 1 }, Value::Number(2.0)),
                ],
            )]),
        };

        assert_eq!(
            counta_range_on_sheet(&grid, &SheetId::Local(0), range),
            Ok(3)
        );
        assert_eq!(
            countblank_range_on_sheet(&grid, &SheetId::Local(0), range),
            Ok((range.rows() as usize).saturating_sub(2))
        );
    }

    #[test]
    fn countif_uses_sparse_iteration_for_sheet_ranges() {
        use std::collections::HashMap;

        struct PanicGrid {
            bounds: (i32, i32),
            cells_by_sheet: HashMap<usize, Vec<(CellCoord, Value)>>,
        }

        impl Grid for PanicGrid {
            fn get_value(&self, _coord: CellCoord) -> Value {
                panic!("unexpected get_value call (expected sparse iteration)");
            }

            fn get_value_on_sheet(&self, _sheet: &SheetId, _coord: CellCoord) -> Value {
                panic!("unexpected get_value_on_sheet call (expected sparse iteration)");
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn column_slice_on_sheet(
                &self,
                _sheet: &SheetId,
                _col: i32,
                _row_start: i32,
                _row_end: i32,
            ) -> Option<&[f64]> {
                None
            }

            fn iter_cells_on_sheet(
                &self,
                sheet: &SheetId,
            ) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                match sheet {
                    SheetId::Local(sheet) => {
                        Some(Box::new(self.cells_by_sheet.get(sheet)?.iter().cloned()))
                    }
                    SheetId::External(_) => None,
                }
            }

            fn bounds(&self) -> (i32, i32) {
                self.bounds
            }

            fn bounds_on_sheet(&self, _sheet: &SheetId) -> (i32, i32) {
                self.bounds
            }
        }

        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        let seen_in_range = 7usize;
        let explicit_zero_matches = 3usize;
        let total_cells = (range.rows() as usize) * (range.cols() as usize);

        let grid = PanicGrid {
            bounds: (row_end + 1, 2),
            cells_by_sheet: HashMap::from([(
                0,
                vec![
                    (CellCoord { row: 0, col: 0 }, Value::Empty),
                    (CellCoord { row: 1, col: 0 }, Value::Number(0.0)),
                    (CellCoord { row: 2, col: 0 }, Value::Number(2.0)),
                    (CellCoord { row: 3, col: 0 }, Value::Bool(true)),
                    (CellCoord { row: 4, col: 0 }, Value::Text(Arc::from("0"))),
                    (CellCoord { row: 5, col: 0 }, Value::Text(Arc::from("x"))),
                    (CellCoord { row: 6, col: 0 }, Value::Error(ErrorKind::Div0)),
                    // Outside the range (different col).
                    (CellCoord { row: 1, col: 1 }, Value::Number(0.0)),
                ],
            )]),
        };

        let criteria_zero = NumericCriteria::new(CmpOp::Eq, 0.0);
        let expected_zero = explicit_zero_matches + total_cells.saturating_sub(seen_in_range);
        assert_eq!(
            count_if_range_on_sheet(&grid, &SheetId::Local(0), range, criteria_zero),
            Ok(expected_zero)
        );

        let criteria_gt = NumericCriteria::new(CmpOp::Gt, 0.0);
        assert_eq!(
            count_if_range_on_sheet(&grid, &SheetId::Local(0), range, criteria_gt),
            Ok(2)
        );
    }

    #[test]
    fn array_aggregates_preserve_error_precedence_with_nan() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };

        let mut values = Vec::new();
        // Ensure we exceed the SIMD threshold.
        for i in 0..(SIMD_ARRAY_MIN_LEN + 8) {
            values.push(Value::Number(i as f64));
        }
        // Inject a NaN value that should not short-circuit error handling.
        values[3] = Value::Number(f64::NAN);
        // Error after NaN should still propagate.
        values.push(Value::Error(ErrorKind::Div0));

        let arr = ArrayValue::new(1, values.len(), values);

        assert_eq!(
            fn_sum(&[Value::Array(arr.clone())], &grid, base),
            Value::Error(ErrorKind::Div0)
        );
        assert_eq!(
            fn_average(&[Value::Array(arr)], &grid, base),
            Value::Error(ErrorKind::Div0)
        );
    }

    #[test]
    fn count_counts_nan_numbers_in_arrays() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };

        let mut values = Vec::new();
        for i in 0..SIMD_ARRAY_MIN_LEN {
            values.push(Value::Number(i as f64));
        }
        values.push(Value::Number(f64::NAN));
        values.push(Value::Bool(true));
        values.push(Value::Text(Arc::from("x")));
        values.push(Value::Error(ErrorKind::Div0));

        let expected_numbers = SIMD_ARRAY_MIN_LEN + 1;
        let arr = ArrayValue::new(1, values.len(), values);

        assert_eq!(
            fn_count(&[Value::Array(arr)], &grid, base),
            Value::Number(expected_numbers as f64)
        );
    }

    #[test]
    fn minmax_return_nan_when_only_nan_numbers_present() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };

        let mut values = vec![Value::Number(f64::NAN); SIMD_ARRAY_MIN_LEN + 4];
        // Mix in some ignored values to ensure they don't affect the result.
        values[0] = Value::Text(Arc::from("x"));
        values[1] = Value::Bool(true);

        let arr = ArrayValue::new(1, values.len(), values);

        let min = fn_min(&[Value::Array(arr.clone())], &grid, base);
        let max = fn_max(&[Value::Array(arr)], &grid, base);

        assert!(matches!(min, Value::Number(n) if n.is_nan()));
        assert!(matches!(max, Value::Number(n) if n.is_nan()));
    }

    #[test]
    fn countif_array_numeric_criteria_counts_nan_for_not_equal() {
        let mut values = vec![Value::Text(Arc::from("x")); SIMD_ARRAY_MIN_LEN + 5];
        values[0] = Value::Number(f64::NAN);
        values[1] = Value::Number(1.0);
        values[2] = Value::Empty; // coerces to 0.0
        values[3] = Value::Text(Arc::from("2")); // numeric text coerces to 2.0
        values[4] = Value::Number(0.0);

        let arr = ArrayValue::new(1, values.len(), values);

        let ne_zero = NumericCriteria::new(CmpOp::Ne, 0.0);
        // NaN != 0 => true, 1 != 0 => true, empty(0) != 0 => false, 2 != 0 => true.
        assert_eq!(count_if_array_numeric_criteria(&arr, ne_zero), 3);

        let eq_zero = NumericCriteria::new(CmpOp::Eq, 0.0);
        // NaN == 0 => false, explicit 0 == 0 => true, empty(0) == 0 => true.
        assert_eq!(count_if_array_numeric_criteria(&arr, eq_zero), 2);
    }

    #[test]
    fn sumif_averageif_array_numeric_criteria_match_scalar_semantics() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };
        let locale = crate::LocaleConfig::en_us();
        let len = SIMD_ARRAY_MIN_LEN + 17;

        let mut crit_values = Vec::with_capacity(len);
        let mut sum_values = Vec::with_capacity(len);
        let mut avg_values = Vec::with_capacity(len);
        for i in 0..len {
            crit_values.push(Value::Number(i as f64));
            // Mix in some ignored values in sum/avg ranges.
            if i % 7 == 0 {
                sum_values.push(Value::Text(Arc::from("x")));
                avg_values.push(Value::Bool(true));
            } else {
                sum_values.push(Value::Number(1.0));
                avg_values.push(Value::Number(i as f64));
            }
        }

        let criteria_arr = ArrayValue::new(1, len, crit_values);
        let sum_arr = ArrayValue::new(1, len, sum_values);
        let avg_arr = ArrayValue::new(1, len, avg_values);

        let criteria = Value::Text(Arc::from(">10"));

        let sum_out = fn_sumif(
            &[
                Value::Array(criteria_arr.clone()),
                criteria.clone(),
                Value::Array(sum_arr),
            ],
            &grid,
            base,
            &locale,
        );
        let mut expected_sum = 0.0;
        for i in 0..len {
            if (i as f64) > 10.0 && i % 7 != 0 {
                expected_sum += 1.0;
            }
        }
        assert_eq!(sum_out, Value::Number(expected_sum));

        let avg_out = fn_averageif(
            &[Value::Array(criteria_arr), criteria, Value::Array(avg_arr)],
            &grid,
            base,
            &locale,
        );
        let mut expected_sum = 0.0;
        let mut expected_count = 0usize;
        for i in 0..len {
            if (i as f64) > 10.0 && i % 7 != 0 {
                expected_sum += i as f64;
                expected_count += 1;
            }
        }
        assert_eq!(avg_out, Value::Number(expected_sum / expected_count as f64));
    }

    #[test]
    fn sumif_averageif_array_numeric_criteria_propagate_errors() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };
        let locale = crate::LocaleConfig::en_us();
        let len = SIMD_ARRAY_MIN_LEN + 8;

        let mut crit_values = Vec::with_capacity(len);
        let mut sum_values = Vec::with_capacity(len);
        let mut avg_values = Vec::with_capacity(len);
        for i in 0..len {
            crit_values.push(Value::Number(i as f64));
            sum_values.push(Value::Number(1.0));
            avg_values.push(Value::Number(1.0));
        }
        // Error at a matching index (15 > 10).
        sum_values[15] = Value::Error(ErrorKind::Div0);
        avg_values[15] = Value::Error(ErrorKind::Div0);

        let criteria_arr = ArrayValue::new(1, len, crit_values);
        let sum_arr = ArrayValue::new(1, len, sum_values);
        let avg_arr = ArrayValue::new(1, len, avg_values);

        let criteria = Value::Text(Arc::from(">10"));

        assert_eq!(
            fn_sumif(
                &[
                    Value::Array(criteria_arr.clone()),
                    criteria.clone(),
                    Value::Array(sum_arr)
                ],
                &grid,
                base,
                &locale,
            ),
            Value::Error(ErrorKind::Div0)
        );

        assert_eq!(
            fn_averageif(
                &[Value::Array(criteria_arr), criteria, Value::Array(avg_arr)],
                &grid,
                base,
                &locale,
            ),
            Value::Error(ErrorKind::Div0)
        );
    }

    #[test]
    fn sumifs_averageifs_countifs_single_array_criteria_match_countif_semantics() {
        let grid = ColumnarGrid::new(1, 1);
        let base = CellCoord { row: 0, col: 0 };
        let locale = crate::LocaleConfig::en_us();
        let len = SIMD_ARRAY_MIN_LEN + 9;

        let mut crit_values = Vec::with_capacity(len);
        let mut sum_values = Vec::with_capacity(len);
        for i in 0..len {
            crit_values.push(Value::Number(i as f64));
            sum_values.push(Value::Number(1.0));
        }
        let criteria_arr = ArrayValue::new(1, len, crit_values);
        let sum_arr = ArrayValue::new(1, len, sum_values);

        let criteria = Value::Text(Arc::from(">10"));

        // SUMIFS(sum_range, criteria_range, criteria)
        assert_eq!(
            fn_sumifs(
                &[
                    Value::Array(sum_arr.clone()),
                    Value::Array(criteria_arr.clone()),
                    criteria.clone()
                ],
                &grid,
                base,
                &locale,
            ),
            Value::Number((len - 11) as f64)
        );

        // AVERAGEIFS(avg_range, criteria_range, criteria)
        assert_eq!(
            fn_averageifs(
                &[
                    Value::Array(sum_arr),
                    Value::Array(criteria_arr.clone()),
                    criteria.clone()
                ],
                &grid,
                base,
                &locale,
            ),
            Value::Number(1.0)
        );

        // COUNTIFS(range, criteria)
        assert_eq!(
            fn_countifs(
                &[Value::Array(criteria_arr), criteria],
                &grid,
                base,
                &locale
            ),
            Value::Number((len - 11) as f64)
        );
    }

    #[test]
    fn and_or_xor_use_sparse_iteration_for_large_sheet_ranges() {
        struct PanicGrid {
            bounds: (i32, i32),
            cells: Vec<(CellCoord, Value)>,
        }

        impl Grid for PanicGrid {
            fn get_value(&self, _coord: CellCoord) -> Value {
                panic!("unexpected get_value call (expected sparse iteration)");
            }

            fn get_value_on_sheet(&self, _sheet: &SheetId, _coord: CellCoord) -> Value {
                panic!("unexpected get_value_on_sheet call (expected sparse iteration)");
            }

            fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
                None
            }

            fn iter_cells(&self) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                Some(Box::new(self.cells.iter().cloned()))
            }

            fn iter_cells_on_sheet(
                &self,
                sheet: &SheetId,
            ) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
                match sheet {
                    SheetId::Local(0) => Some(Box::new(self.cells.iter().cloned())),
                    _ => None,
                }
            }

            fn bounds(&self) -> (i32, i32) {
                self.bounds
            }

            fn bounds_on_sheet(&self, _sheet: &SheetId) -> (i32, i32) {
                self.bounds
            }
        }

        let row_end = BYTECODE_SPARSE_RANGE_ROW_THRESHOLD; // rows() == threshold + 1
        let range = ResolvedRange {
            row_start: 0,
            row_end,
            col_start: 0,
            col_end: 0,
        };

        // AND: ignore text, observe `0` as false.
        let grid = PanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![
                (CellCoord { row: 2, col: 0 }, Value::Text(Arc::from("x"))),
                (CellCoord { row: 1, col: 0 }, Value::Bool(true)),
                (CellCoord { row: 0, col: 0 }, Value::Number(0.0)),
            ],
        };
        let mut all_true = true;
        let mut any = false;
        assert_eq!(and_range(&grid, range, &mut all_true, &mut any), None);
        assert!(!all_true);
        assert!(any);

        // OR: ignore text, observe `TRUE` as true.
        let grid = PanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![
                (CellCoord { row: 2, col: 0 }, Value::Text(Arc::from("y"))),
                (CellCoord { row: 1, col: 0 }, Value::Bool(true)),
                (CellCoord { row: 0, col: 0 }, Value::Number(0.0)),
            ],
        };
        let mut any_true = false;
        let mut any = false;
        assert_eq!(
            or_range_on_sheet(&grid, &SheetId::Local(0), range, &mut any_true, &mut any),
            None
        );
        assert!(any_true);
        assert!(any);

        // XOR: parity across non-zero/TRUE values.
        let grid = PanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![
                (CellCoord { row: 2, col: 0 }, Value::Number(1.0)),
                (CellCoord { row: 1, col: 0 }, Value::Bool(true)),
            ],
        };
        let mut acc = false;
        assert_eq!(
            xor_range_on_sheet(&grid, &SheetId::Local(0), range, &mut acc),
            None
        );
        assert!(!acc, "TRUE XOR 1 should yield FALSE");

        // Error precedence: row-major (smaller row wins) regardless of iteration order.
        let grid = PanicGrid {
            bounds: (row_end + 1, 1),
            cells: vec![
                (CellCoord { row: 10, col: 0 }, Value::Error(ErrorKind::Div0)),
                (CellCoord { row: 5, col: 0 }, Value::Error(ErrorKind::Num)),
            ],
        };
        let mut all_true = true;
        let mut any = false;
        assert_eq!(
            and_range_on_sheet(&grid, &SheetId::Local(0), range, &mut all_true, &mut any),
            Some(ErrorKind::Num)
        );
    }
}
