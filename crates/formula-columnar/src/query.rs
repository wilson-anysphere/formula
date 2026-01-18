#![forbid(unsafe_code)]

use crate::bitmap::BitVec;
use crate::encoding::{EncodedChunk, U32SequenceEncoding, U64SequenceEncoding};
use crate::table::{ColumnSchema, ColumnarTable, ColumnarTableBuilder, TableOptions};
use crate::types::{ColumnType, Value};
use std::collections::HashMap;
use std::hash::{BuildHasherDefault, Hasher};
use std::sync::Arc;

/// Aggregation operator supported by the columnar query engine.
///
/// ## Null semantics
///
/// Unless noted otherwise, aggregations ignore nulls (blanks).
///
/// - [`AggOp::Count`], [`AggOp::CountNumbers`], and [`AggOp::DistinctCount`] always return a
///   numeric value (never null). When the input has no qualifying values for a group they return
///   `0.0`.
/// - [`AggOp::SumF64`] and [`AggOp::AvgF64`] return `Value::Null` when a group has no numeric
///   values.
/// - [`AggOp::Var`], [`AggOp::VarP`], [`AggOp::StdDev`], and [`AggOp::StdDevP`] return
///   `Value::Null` when a group has no numeric values; the sample variants (`Var`, `StdDev`)
///   additionally return `Value::Null` when the group has fewer than 2 numeric values.
/// - [`AggOp::Min`] / [`AggOp::Max`] return `Value::Null` when a group has no non-null values.
///
/// ## DistinctCount details
///
/// [`AggOp::DistinctCount`] counts distinct **non-null** values. For `ColumnType::Number` it
/// canonicalizes `-0.0` to `0.0` and all NaN bit patterns to a single canonical NaN (mirroring
/// the internal `canonical_f64_bits` normalization) before deduplication.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum AggOp {
    Count,
    /// Average of numeric values (ignoring nulls).
    AvgF64,
    SumF64,
    /// Count distinct non-null values.
    DistinctCount,
    /// Count non-null numeric values (Excel/Pivot `COUNT` semantics).
    CountNumbers,
    /// Sample variance (n-1 denominator), ignoring nulls; null when the group has <2 values.
    Var,
    /// Population variance (n denominator), ignoring nulls; null when the group has 0 values.
    VarP,
    /// Sample standard deviation (sqrt(var)), ignoring nulls; null when the group has <2 values.
    StdDev,
    /// Population standard deviation (sqrt(varp)), ignoring nulls; null when the group has 0 values.
    StdDevP,
    Min,
    Max,
}

/// Aggregation specification for `GROUP BY`.
///
/// Notes:
/// - `AggOp::Count` with `column: None` counts rows in the group.
/// - `AggOp::Count` with `column: Some(i)` counts non-null values of column `i`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AggSpec {
    pub op: AggOp,
    pub column: Option<usize>,
    pub name: Option<String>,
}

impl AggSpec {
    pub fn count_rows() -> Self {
        Self {
            op: AggOp::Count,
            column: None,
            name: None,
        }
    }

    pub fn count_non_null(column: usize) -> Self {
        Self {
            op: AggOp::Count,
            column: Some(column),
            name: None,
        }
    }

    pub fn sum_f64(column: usize) -> Self {
        Self {
            op: AggOp::SumF64,
            column: Some(column),
            name: None,
        }
    }

    pub fn avg_f64(column: usize) -> Self {
        Self {
            op: AggOp::AvgF64,
            column: Some(column),
            name: None,
        }
    }

    pub fn distinct_count(column: usize) -> Self {
        Self {
            op: AggOp::DistinctCount,
            column: Some(column),
            name: None,
        }
    }

    pub fn count_numbers(column: usize) -> Self {
        Self {
            op: AggOp::CountNumbers,
            column: Some(column),
            name: None,
        }
    }

    pub fn var(column: usize) -> Self {
        Self {
            op: AggOp::Var,
            column: Some(column),
            name: None,
        }
    }

    pub fn var_p(column: usize) -> Self {
        Self {
            op: AggOp::VarP,
            column: Some(column),
            name: None,
        }
    }

    pub fn std_dev(column: usize) -> Self {
        Self {
            op: AggOp::StdDev,
            column: Some(column),
            name: None,
        }
    }

    pub fn std_dev_p(column: usize) -> Self {
        Self {
            op: AggOp::StdDevP,
            column: Some(column),
            name: None,
        }
    }

    pub fn min(column: usize) -> Self {
        Self {
            op: AggOp::Min,
            column: Some(column),
            name: None,
        }
    }

    pub fn max(column: usize) -> Self {
        Self {
            op: AggOp::Max,
            column: Some(column),
            name: None,
        }
    }

    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum QueryError {
    EmptyKeys,
    ColumnOutOfBounds { col: usize, column_count: usize },
    RowOutOfBounds { row: usize, row_count: usize },
    UnsupportedColumnType {
        col: usize,
        column_type: ColumnType,
        operation: &'static str,
    },
    MismatchedJoinKeyCount { left: usize, right: usize },
    MismatchedJoinKeyTypes {
        left_type: ColumnType,
        right_type: ColumnType,
    },
    MissingDictionary { col: usize },
    InternalInvariant(&'static str),
}

impl std::fmt::Display for QueryError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::EmptyKeys => write!(f, "at least one key column is required"),
            Self::ColumnOutOfBounds { col, column_count } => write!(
                f,
                "column index {} out of bounds (table has {} columns)",
                col, column_count
            ),
            Self::RowOutOfBounds { row, row_count } => write!(
                f,
                "row index {} out of bounds (table has {} rows)",
                row, row_count
            ),
            Self::UnsupportedColumnType {
                col,
                column_type,
                operation,
            } => write!(
                f,
                "unsupported column type {:?} for column {} in {}",
                column_type, col, operation
            ),
            Self::MismatchedJoinKeyCount { left, right } => write!(
                f,
                "join requires the same number of key columns on both sides (left_keys={}, right_keys={})",
                left, right
            ),
            Self::MismatchedJoinKeyTypes {
                left_type,
                right_type,
            } => write!(
                f,
                "join key column types do not match: left={:?}, right={:?}",
                left_type, right_type
            ),
            Self::MissingDictionary { col } => write!(f, "missing dictionary for string column {}", col),
            Self::InternalInvariant(msg) => write!(f, "internal invariant violated: {}", msg),
        }
    }
}

impl std::error::Error for QueryError {}

/// Comparison operators supported by [`FilterExpr`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum CmpOp {
    Eq,
    Ne,
    Lt,
    Lte,
    Gt,
    Gte,
}

/// A scalar literal used in a [`FilterExpr::Cmp`].
#[derive(Clone, Debug, PartialEq)]
pub enum FilterValue {
    Number(f64),
    Boolean(bool),
    String(Arc<str>),
}

impl FilterValue {
    pub fn string(value: impl Into<Arc<str>>) -> Self {
        Self::String(value.into())
    }

    pub fn number(value: f64) -> Self {
        Self::Number(value)
    }

    pub fn boolean(value: bool) -> Self {
        Self::Boolean(value)
    }
}

impl From<f64> for FilterValue {
    fn from(value: f64) -> Self {
        Self::Number(value)
    }
}

impl From<bool> for FilterValue {
    fn from(value: bool) -> Self {
        Self::Boolean(value)
    }
}

impl From<Arc<str>> for FilterValue {
    fn from(value: Arc<str>) -> Self {
        Self::String(value)
    }
}

impl From<String> for FilterValue {
    fn from(value: String) -> Self {
        Self::String(Arc::<str>::from(value))
    }
}

impl From<&str> for FilterValue {
    fn from(value: &str) -> Self {
        Self::String(Arc::<str>::from(value))
    }
}

/// A small predicate AST that can be evaluated column-wise against a [`ColumnarTable`].
#[derive(Clone, Debug, PartialEq)]
pub enum FilterExpr {
    And(Box<FilterExpr>, Box<FilterExpr>),
    Or(Box<FilterExpr>, Box<FilterExpr>),
    Not(Box<FilterExpr>),
    Cmp {
        col: usize,
        op: CmpOp,
        value: FilterValue,
    },
    /// Case-insensitive string comparison (ASCII-only).
    ///
    /// This is an optional extension over the default case-sensitive dictionary comparisons.
    /// Equality can match multiple dictionary entries (e.g. "A" and "a"), so evaluation uses a
    /// precomputed set of matching dictionary indices.
    CmpStringCI {
        col: usize,
        op: CmpOp,
        value: Arc<str>,
    },
    IsNull {
        col: usize,
    },
    IsNotNull {
        col: usize,
    },
}

impl FilterExpr {
    pub fn and(self, rhs: FilterExpr) -> Self {
        Self::And(Box::new(self), Box::new(rhs))
    }

    pub fn or(self, rhs: FilterExpr) -> Self {
        Self::Or(Box::new(self), Box::new(rhs))
    }

    pub fn not(self) -> Self {
        Self::Not(Box::new(self))
    }

    pub fn cmp(col: usize, op: CmpOp, value: impl Into<FilterValue>) -> Self {
        Self::Cmp {
            col,
            op,
            value: value.into(),
        }
    }

    pub fn cmp_string_ci(col: usize, op: CmpOp, value: impl Into<Arc<str>>) -> Self {
        Self::CmpStringCI {
            col,
            op,
            value: value.into(),
        }
    }

    pub fn is_null(col: usize) -> Self {
        Self::IsNull { col }
    }

    pub fn is_not_null(col: usize) -> Self {
        Self::IsNotNull { col }
    }
}

/// Evaluate a filter expression and return a [`BitVec`] mask of matching rows.
///
/// Notes:
/// - Comparisons treat NULL as "unknown" and therefore evaluate to `false` for filtering.
/// - `IS NULL` / `IS NOT NULL` explicitly test the encoded validity bits.
pub fn filter_mask(table: &ColumnarTable, expr: &FilterExpr) -> Result<BitVec, QueryError> {
    if expr_contains_not(expr) {
        Ok(eval_filter_expr_tri(table, expr)?.true_mask)
    } else {
        eval_filter_expr(table, expr)
    }
}

/// Evaluate a filter expression and return the matching row indices.
///
/// This is a convenience helper for callers that need an explicit row mapping (e.g. interop with
/// APIs that accept `&[usize]`). Prefer [`filter_mask`] when possible to avoid materializing an
/// index vector.
pub fn filter_indices(table: &ColumnarTable, expr: &FilterExpr) -> Result<Vec<usize>, QueryError> {
    let mask = filter_mask(table, expr)?;
    if mask.count_ones() == 0 {
        return Ok(Vec::new());
    }
    if mask.all_true() {
        return Ok((0..table.row_count()).collect());
    }
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(mask.count_ones());
    out.extend(mask.iter_ones());
    Ok(out)
}

/// Materialize a filtered table using a previously computed row mask.
pub fn filter_table(table: &ColumnarTable, mask: &BitVec) -> Result<ColumnarTable, QueryError> {
    if mask.len() != table.row_count() {
        return Err(QueryError::InternalInvariant("filter mask length must match table"));
    }

    if mask.count_ones() == 0 {
        return Ok(ColumnarTableBuilder::new(table.schema().to_vec(), table.options()).finalize());
    }

    if mask.all_true() {
        return Ok(table.clone());
    }

    let schema: Vec<ColumnSchema> = table.schema().to_vec();
    let mut builder = ColumnarTableBuilder::new(schema, table.options());
    let mut scratch_row: Vec<Value> = vec![Value::Null; table.column_count()];

    for row in mask.iter_ones() {
        for col in 0..table.column_count() {
            scratch_row[col] = table.get_cell(row, col);
        }
        builder.append_row(&scratch_row);
    }

    Ok(builder.finalize())
}

fn expr_contains_not(expr: &FilterExpr) -> bool {
    match expr {
        FilterExpr::Not(_) => true,
        FilterExpr::And(left, right) | FilterExpr::Or(left, right) => {
            expr_contains_not(left) || expr_contains_not(right)
        }
        FilterExpr::Cmp { .. }
        | FilterExpr::CmpStringCI { .. }
        | FilterExpr::IsNull { .. }
        | FilterExpr::IsNotNull { .. } => false,
    }
}

#[derive(Clone)]
struct TriMask {
    true_mask: BitVec,
    unknown_mask: BitVec,
}

impl TriMask {
    fn false_mask(&self) -> BitVec {
        let mut out = self.true_mask.clone();
        out.or_inplace(&self.unknown_mask);
        out.not_inplace();
        out
    }
}

fn eval_filter_expr_tri(table: &ColumnarTable, expr: &FilterExpr) -> Result<TriMask, QueryError> {
    match expr {
        FilterExpr::And(left, right) => {
            let left = eval_filter_expr_tri(table, left)?;
            if left.true_mask.count_ones() == 0 && left.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_false(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            if left.true_mask.all_true() && left.unknown_mask.count_ones() == 0 {
                return eval_filter_expr_tri(table, right);
            }

            let right = eval_filter_expr_tri(table, right)?;
            if right.true_mask.count_ones() == 0 && right.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_false(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            if right.true_mask.all_true() && right.unknown_mask.count_ones() == 0 {
                return Ok(left);
            }

            let mut true_mask = left.true_mask.clone();
            true_mask.and_inplace(&right.true_mask);

            let mut false_mask = left.false_mask();
            false_mask.or_inplace(&right.false_mask());

            let mut unknown_mask = true_mask.clone();
            unknown_mask.or_inplace(&false_mask);
            unknown_mask.not_inplace();

            Ok(TriMask {
                true_mask,
                unknown_mask,
            })
        }
        FilterExpr::Or(left, right) => {
            let left = eval_filter_expr_tri(table, left)?;
            if left.true_mask.all_true() && left.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_true(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            if left.true_mask.count_ones() == 0 && left.unknown_mask.count_ones() == 0 {
                return eval_filter_expr_tri(table, right);
            }

            let right = eval_filter_expr_tri(table, right)?;
            if right.true_mask.all_true() && right.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_true(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            if right.true_mask.count_ones() == 0 && right.unknown_mask.count_ones() == 0 {
                return Ok(left);
            }

            let mut true_mask = left.true_mask.clone();
            true_mask.or_inplace(&right.true_mask);

            let mut false_mask = left.false_mask();
            false_mask.and_inplace(&right.false_mask());

            let mut unknown_mask = true_mask.clone();
            unknown_mask.or_inplace(&false_mask);
            unknown_mask.not_inplace();

            Ok(TriMask {
                true_mask,
                unknown_mask,
            })
        }
        FilterExpr::Not(inner) => {
            let inner = eval_filter_expr_tri(table, inner)?;
            if inner.true_mask.all_true() && inner.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_false(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            if inner.true_mask.count_ones() == 0 && inner.unknown_mask.count_ones() == 0 {
                let rows = table.row_count();
                return Ok(TriMask {
                    true_mask: BitVec::with_len_all_true(rows),
                    unknown_mask: BitVec::with_len_all_false(rows),
                });
            }
            Ok(TriMask {
                true_mask: inner.false_mask(),
                unknown_mask: inner.unknown_mask,
            })
        }
        FilterExpr::Cmp { col, op, value } => {
            let true_mask = eval_filter_cmp(table, *col, *op, value)?;
            let unknown_mask = eval_filter_is_null(table, *col, true)?;
            Ok(TriMask {
                true_mask,
                unknown_mask,
            })
        }
        FilterExpr::CmpStringCI { col, op, value } => {
            let true_mask = eval_filter_string_ci(table, *col, *op, value.as_ref())?;
            let unknown_mask = eval_filter_is_null(table, *col, true)?;
            Ok(TriMask {
                true_mask,
                unknown_mask,
            })
        }
        FilterExpr::IsNull { col } => Ok(TriMask {
            true_mask: eval_filter_is_null(table, *col, true)?,
            unknown_mask: BitVec::with_len_all_false(table.row_count()),
        }),
        FilterExpr::IsNotNull { col } => Ok(TriMask {
            true_mask: eval_filter_is_null(table, *col, false)?,
            unknown_mask: BitVec::with_len_all_false(table.row_count()),
        }),
    }
}

fn eval_filter_expr(table: &ColumnarTable, expr: &FilterExpr) -> Result<BitVec, QueryError> {
    match expr {
        FilterExpr::And(left, right) => {
            let left_mask = eval_filter_expr(table, left)?;
            if left_mask.count_ones() == 0 {
                return Ok(left_mask);
            }
            if left_mask.all_true() {
                return eval_filter_expr(table, right);
            }
            let mut left_mask = left_mask;
            let right_mask = eval_filter_expr(table, right)?;
            if right_mask.count_ones() == 0 {
                return Ok(right_mask);
            }
            if right_mask.all_true() {
                return Ok(left_mask);
            }
            left_mask.and_inplace(&right_mask);
            Ok(left_mask)
        }
        FilterExpr::Or(left, right) => {
            let left_mask = eval_filter_expr(table, left)?;
            if left_mask.all_true() {
                return Ok(left_mask);
            }
            if left_mask.count_ones() == 0 {
                return eval_filter_expr(table, right);
            }
            let mut left_mask = left_mask;
            let right_mask = eval_filter_expr(table, right)?;
            if right_mask.all_true() {
                return Ok(right_mask);
            }
            if right_mask.count_ones() == 0 {
                return Ok(left_mask);
            }
            left_mask.or_inplace(&right_mask);
            Ok(left_mask)
        }
        FilterExpr::Not(inner) => {
            let mut mask = eval_filter_expr(table, inner)?;
            mask.not_inplace();
            Ok(mask)
        }
        FilterExpr::Cmp { col, op, value } => eval_filter_cmp(table, *col, *op, value),
        FilterExpr::CmpStringCI { col, op, value } => eval_filter_string_ci(table, *col, *op, value.as_ref()),
        FilterExpr::IsNull { col } => eval_filter_is_null(table, *col, true),
        FilterExpr::IsNotNull { col } => eval_filter_is_null(table, *col, false),
    }
}

fn eval_filter_cmp(
    table: &ColumnarTable,
    col: usize,
    op: CmpOp,
    value: &FilterValue,
) -> Result<BitVec, QueryError> {
    match value {
        FilterValue::Number(v) => eval_filter_f64(table, col, op, *v),
        FilterValue::Boolean(v) => eval_filter_bool(table, col, op, *v),
        FilterValue::String(v) => eval_filter_string(table, col, op, v.as_ref()),
    }
}

fn eval_filter_f64(table: &ColumnarTable, col: usize, op: CmpOp, rhs: f64) -> Result<BitVec, QueryError> {
    let column_type = table
        .schema()
        .get(col)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    if column_type != ColumnType::Number {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter numeric comparison",
        });
    }

    let chunks = table
        .encoded_chunks(col)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    let rows = table.row_count();
    let page = table.page_size_rows();
    let rhs_is_zero = rhs == 0.0;
    let rhs_is_nan = rhs.is_nan();

    if rhs_is_nan && matches!(op, CmpOp::Lt | CmpOp::Lte | CmpOp::Gt | CmpOp::Gte) {
        // All comparisons with NaN (other than =/!= which we special-case) are false.
        return Ok(BitVec::with_len_all_false(rows));
    }

    if let Some(stats) = table.stats(col) {
        let nulls = stats.null_count as usize;
        if nulls == rows {
            return Ok(BitVec::with_len_all_false(rows));
        }

        if let (Some(Value::Number(min)), Some(Value::Number(max))) = (&stats.min, &stats.max) {
            match op {
                CmpOp::Eq if !rhs.is_nan() => {
                    // If `rhs` is outside the observed range it can't match any non-null value.
                    if rhs < *min || rhs > *max {
                        return Ok(BitVec::with_len_all_false(rows));
                    }
                }
                CmpOp::Ne if !rhs.is_nan() => {
                    // If `rhs` is outside the observed range then all non-null values are != rhs.
                    if rhs < *min || rhs > *max {
                        return eval_filter_is_null(table, col, false);
                    }
                }
                CmpOp::Lt => {
                    if *min >= rhs {
                        return Ok(BitVec::with_len_all_false(rows));
                    }
                }
                CmpOp::Lte => {
                    if *min > rhs {
                        return Ok(BitVec::with_len_all_false(rows));
                    }
                }
                CmpOp::Gt => {
                    if *max <= rhs {
                        return Ok(BitVec::with_len_all_false(rows));
                    }
                }
                CmpOp::Gte => {
                    if *max < rhs {
                        return Ok(BitVec::with_len_all_false(rows));
                    }
                }
                _ => {}
            }
        }
    }

    let mut out = BitVec::with_capacity_bits(rows);
    for (chunk_idx, chunk) in chunks.iter().enumerate() {
        let base = chunk_idx * page;
        if base >= rows {
            break;
        }
        let chunk_rows = (rows - base).min(chunk.len());

        let EncodedChunk::Float(c) = chunk else {
            return Err(QueryError::UnsupportedColumnType {
                col,
                column_type,
                operation: "filter numeric comparison",
            });
        };

        let eval = |v: f64| -> bool {
            match op {
                CmpOp::Eq => {
                    if rhs_is_zero {
                        v == 0.0
                    } else if rhs_is_nan {
                        v.is_nan()
                    } else {
                        v == rhs
                    }
                }
                CmpOp::Ne => {
                    if rhs_is_zero {
                        v != 0.0
                    } else if rhs_is_nan {
                        !v.is_nan()
                    } else {
                        v != rhs
                    }
                }
                CmpOp::Lt => v < rhs,
                CmpOp::Lte => v <= rhs,
                CmpOp::Gt => v > rhs,
                CmpOp::Gte => v >= rhs,
            }
        };

        match c.validity.as_ref() {
            None => {
                for i in 0..chunk_rows {
                    out.push(eval(c.values[i]));
                }
            }
            Some(validity) => {
                if validity.count_ones() == 0 {
                    out.extend_constant(false, chunk_rows);
                    continue;
                }
                if validity.all_true() {
                    for i in 0..chunk_rows {
                        out.push(eval(c.values[i]));
                    }
                } else {
                    for i in 0..chunk_rows {
                        if !validity.get(i) {
                            out.push(false);
                            continue;
                        }
                        out.push(eval(c.values[i]));
                    }
                }
            }
        }
    }

    Ok(out)
}

fn eval_filter_bool(
    table: &ColumnarTable,
    col: usize,
    op: CmpOp,
    rhs: bool,
) -> Result<BitVec, QueryError> {
    let column_type = table
        .schema()
        .get(col)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    if column_type != ColumnType::Boolean {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter boolean comparison",
        });
    }

    if !matches!(op, CmpOp::Eq | CmpOp::Ne) {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter boolean comparison",
        });
    }

    let chunks = table
        .encoded_chunks(col)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    let rows = table.row_count();
    let page = table.page_size_rows();

    if let Some(stats) = table.stats(col) {
        let nulls = stats.null_count as usize;
        if nulls == rows {
            return Ok(BitVec::with_len_all_false(rows));
        }
        if let Some(sum) = stats.sum {
            let true_count = sum.round().max(0.0) as usize;
            let non_null = rows.saturating_sub(nulls);
            let want = match op {
                CmpOp::Eq => rhs,
                CmpOp::Ne => !rhs,
                _ => rhs,
            };

            // Use the true-count statistics to quickly answer fully-satisfied or fully-unsatisfied
            // predicates without scanning encoded chunks.
            if want {
                if true_count == 0 {
                    return Ok(BitVec::with_len_all_false(rows));
                }
                if true_count == non_null {
                    return eval_filter_is_null(table, col, false);
                }
            } else {
                if true_count == 0 {
                    return eval_filter_is_null(table, col, false);
                }
                if true_count == non_null {
                    return Ok(BitVec::with_len_all_false(rows));
                }
            }
        }
    }

    let mut out = BitVec::with_capacity_bits(rows);
    for (chunk_idx, chunk) in chunks.iter().enumerate() {
        let base = chunk_idx * page;
        if base >= rows {
            break;
        }
        let chunk_rows = (rows - base).min(chunk.len());

        let EncodedChunk::Bool(c) = chunk else {
            return Err(QueryError::UnsupportedColumnType {
                col,
                column_type,
                operation: "filter boolean comparison",
            });
        };

        if c.validity.as_ref().is_some_and(|v| v.count_ones() == 0) {
            // All-null chunk: comparisons evaluate to false.
            out.extend_constant(false, chunk_rows);
            continue;
        }

        if c.validity.is_none() || c.validity.as_ref().is_some_and(|v| v.all_true()) {
            // Byte-wise fast path for non-null boolean chunks.
            let invert = match op {
                CmpOp::Eq => !rhs,
                CmpOp::Ne => rhs,
                _ => false,
            };

            let full_bytes = chunk_rows / 8;
            let rem_bits = chunk_rows % 8;

            for b in c.data.iter().take(full_bytes) {
                let byte = if invert { !*b } else { *b };
                match byte {
                    0x00 => out.extend_constant(false, 8),
                    0xFF => out.extend_constant(true, 8),
                    _ => {
                        for bit in 0..8 {
                            out.push(((byte >> bit) & 1) == 1);
                        }
                    }
                }
            }

            if rem_bits > 0 {
                let byte_idx = full_bytes;
                if let Some(b) = c.data.get(byte_idx) {
                    let byte = if invert { !*b } else { *b };
                    for bit in 0..rem_bits {
                        out.push(((byte >> bit) & 1) == 1);
                    }
                }
            }

            continue;
        }

        for i in 0..chunk_rows {
            if c.validity.as_ref().is_some_and(|v| !v.get(i)) {
                out.push(false);
                continue;
            }
            let byte = c.data[i / 8];
            let bit = i % 8;
            let v = ((byte >> bit) & 1) == 1;
            let passed = match op {
                CmpOp::Eq => v == rhs,
                CmpOp::Ne => v != rhs,
                _ => false,
            };
            out.push(passed);
        }
    }

    Ok(out)
}

fn eval_filter_string(
    table: &ColumnarTable,
    col: usize,
    op: CmpOp,
    rhs: &str,
) -> Result<BitVec, QueryError> {
    let column_type = table
        .schema()
        .get(col)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    if column_type != ColumnType::String {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter string comparison",
        });
    }

    if !matches!(op, CmpOp::Eq | CmpOp::Ne) {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter string comparison",
        });
    }

    let dict = table.dictionary(col).ok_or(QueryError::MissingDictionary { col })?;
    if let Some(stats) = table.stats(col) {
        let rows = table.row_count();
        let nulls = stats.null_count as usize;
        if nulls == rows {
            // Comparisons treat NULL as false, so this is always false when the column is entirely null.
            return Ok(BitVec::with_len_all_false(rows));
        }

        // Quick reject using lexicographic min/max. If rhs is outside the observed range, it cannot
        // be present in the dictionary.
        if let (Some(Value::String(min)), Some(Value::String(max))) = (&stats.min, &stats.max) {
            if rhs < min.as_ref() || rhs > max.as_ref() {
                return Ok(match op {
                    CmpOp::Eq => BitVec::with_len_all_false(rows),
                    CmpOp::Ne => eval_filter_is_null(table, col, false)?,
                    _ => BitVec::with_len_all_false(rows),
                });
            }
        }
    }
    let target = dict
        .iter()
        .enumerate()
        .find_map(|(idx, s)| (s.as_ref() == rhs).then_some(idx as u32));

    let chunks = table
        .encoded_chunks(col)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    let rows = table.row_count();
    let page = table.page_size_rows();

    let target = match (op, target) {
        (CmpOp::Eq, None) => return Ok(BitVec::with_len_all_false(rows)),
        (CmpOp::Ne, None) => return eval_filter_is_null(table, col, false),
        (_, Some(t)) => t,
        _ => return Err(QueryError::InternalInvariant("unexpected string comparison op")),
    };

    let mut out = BitVec::with_capacity_bits(rows);
    for (chunk_idx, chunk) in chunks.iter().enumerate() {
        let base = chunk_idx * page;
        if base >= rows {
            break;
        }
        let chunk_rows = (rows - base).min(chunk.len());

        let EncodedChunk::Dict(c) = chunk else {
            return Err(QueryError::UnsupportedColumnType {
                col,
                column_type,
                operation: "filter string comparison",
            });
        };

        if c.validity.as_ref().is_some_and(|v| v.count_ones() == 0) {
            // All-null chunk: comparisons evaluate to false.
            out.extend_constant(false, chunk_rows);
            continue;
        }

        // When the chunk has no nulls and uses RLE indices we can preserve some compression
        // benefits by emitting whole runs at once.
        if c.validity.is_none() || c.validity.as_ref().is_some_and(|v| v.all_true()) {
            match &c.indices {
                U32SequenceEncoding::Rle(rle) => {
                    let mut start: usize = 0;
                    for (&run_value, &end) in rle.values.iter().zip(rle.ends.iter()) {
                        let end = end as usize;
                        if start >= chunk_rows {
                            break;
                        }
                        let run_len = end.saturating_sub(start).min(chunk_rows - start);
                        let passed = match op {
                            CmpOp::Eq => run_value == target,
                            CmpOp::Ne => run_value != target,
                            _ => false,
                        };
                        out.extend_constant(passed, run_len);
                        start = end;
                    }
                    continue;
                }
                U32SequenceEncoding::Bitpacked { .. } => {}
            }
        }

        let mut cursor = U32SeqCursor::new(&c.indices);
        if c.validity.is_none() || c.validity.as_ref().is_some_and(|v| v.all_true()) {
            for _i in 0..chunk_rows {
                let ix = cursor.next();
                let passed = match op {
                    CmpOp::Eq => ix == target,
                    CmpOp::Ne => ix != target,
                    _ => false,
                };
                out.push(passed);
            }
        } else {
            let validity = c
                .validity
                .as_ref()
                .ok_or(QueryError::InternalInvariant("missing validity"))?;
            for i in 0..chunk_rows {
                let ix = cursor.next();
                if !validity.get(i) {
                    out.push(false);
                    continue;
                }
                let passed = match op {
                    CmpOp::Eq => ix == target,
                    CmpOp::Ne => ix != target,
                    _ => false,
                };
                out.push(passed);
            }
        }
    }

    Ok(out)
}

fn eval_filter_string_ci(
    table: &ColumnarTable,
    col: usize,
    op: CmpOp,
    rhs: &str,
) -> Result<BitVec, QueryError> {
    let column_type = table
        .schema()
        .get(col)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds {
            col,
            column_count: table.column_count(),
        })?;
    if column_type != ColumnType::String {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter case-insensitive string comparison",
        });
    }

    if !matches!(op, CmpOp::Eq | CmpOp::Ne) {
        return Err(QueryError::UnsupportedColumnType {
            col,
            column_type,
            operation: "filter case-insensitive string comparison",
        });
    }

    if let Some(stats) = table.stats(col) {
        let rows = table.row_count();
        let nulls = stats.null_count as usize;
        if nulls == rows {
            // Comparisons treat NULL as false.
            return Ok(BitVec::with_len_all_false(rows));
        }
    }

    let dict = table.dictionary(col).ok_or(QueryError::MissingDictionary { col })?;
    let rows = table.row_count();
    let page = table.page_size_rows();

    // Build a membership bitmap of all dictionary indices that match `rhs` case-insensitively.
    // We use ASCII-only folding to avoid per-entry allocations.
    let mut targets = BitVec::with_len_all_false(dict.len());
    for (idx, s) in dict.iter().enumerate() {
        if s.as_ref().eq_ignore_ascii_case(rhs) {
            targets.set(idx, true);
        }
    }

    let target_count = targets.count_ones();
    if target_count == 0 {
        return match op {
            CmpOp::Eq => Ok(BitVec::with_len_all_false(rows)),
            CmpOp::Ne => eval_filter_is_null(table, col, false),
            other => {
                debug_assert!(
                    false,
                    "unexpected comparison op for case-insensitive string filter: {other:?}"
                );
                Err(QueryError::UnsupportedColumnType {
                    col,
                    column_type,
                    operation: "filter case-insensitive string comparison",
                })
            }
        };
    }

    if target_count == dict.len() {
        // All dictionary entries match; equality matches all non-null rows, inequality matches none.
        return match op {
            CmpOp::Eq => eval_filter_is_null(table, col, false),
            CmpOp::Ne => Ok(BitVec::with_len_all_false(rows)),
            other => {
                debug_assert!(
                    false,
                    "unexpected comparison op for case-insensitive string filter: {other:?}"
                );
                Err(QueryError::UnsupportedColumnType {
                    col,
                    column_type,
                    operation: "filter case-insensitive string comparison",
                })
            }
        };
    }

    let chunks = table
        .encoded_chunks(col)
        .ok_or(QueryError::ColumnOutOfBounds {
            col,
            column_count: table.column_count(),
        })?;

    let invert = matches!(op, CmpOp::Ne);

    let mut out = BitVec::with_capacity_bits(rows);
    for (chunk_idx, chunk) in chunks.iter().enumerate() {
        let base = chunk_idx * page;
        if base >= rows {
            break;
        }
        let chunk_rows = (rows - base).min(chunk.len());

        let EncodedChunk::Dict(c) = chunk else {
            return Err(QueryError::UnsupportedColumnType {
                col,
                column_type,
                operation: "filter case-insensitive string comparison",
            });
        };

        if c.validity.as_ref().is_some_and(|v| v.count_ones() == 0) {
            out.extend_constant(false, chunk_rows);
            continue;
        }

        let non_null_chunk = c.validity.is_none() || c.validity.as_ref().is_some_and(|v| v.all_true());

        if non_null_chunk {
            // Run-level fast path when indices are RLE.
            if let U32SequenceEncoding::Rle(rle) = &c.indices {
                let mut start: usize = 0;
                for (&run_value, &end) in rle.values.iter().zip(rle.ends.iter()) {
                    let end = end as usize;
                    if start >= chunk_rows {
                        break;
                    }
                    let run_len = end.saturating_sub(start).min(chunk_rows - start);
                    let mut passed = targets.get(run_value as usize);
                    if invert {
                        passed = !passed;
                    }
                    out.extend_constant(passed, run_len);
                    start = end;
                }
                continue;
            }

            let mut cursor = U32SeqCursor::new(&c.indices);
            for _i in 0..chunk_rows {
                let ix = cursor.next();
                let mut passed = targets.get(ix as usize);
                if invert {
                    passed = !passed;
                }
                out.push(passed);
            }
        } else {
            let validity = c
                .validity
                .as_ref()
                .ok_or(QueryError::InternalInvariant("missing validity bitmap"))?;
            let mut cursor = U32SeqCursor::new(&c.indices);
            for i in 0..chunk_rows {
                let ix = cursor.next();
                if !validity.get(i) {
                    out.push(false);
                    continue;
                }
                let mut passed = targets.get(ix as usize);
                if invert {
                    passed = !passed;
                }
                out.push(passed);
            }
        }
    }

    Ok(out)
}

fn eval_filter_is_null(table: &ColumnarTable, col: usize, is_null: bool) -> Result<BitVec, QueryError> {
    if col >= table.column_count() {
        return Err(QueryError::ColumnOutOfBounds {
            col,
            column_count: table.column_count(),
        });
    }

    if let Some(stats) = table.stats(col) {
        let nulls = stats.null_count as usize;
        let rows = table.row_count();
        if nulls == 0 {
            return Ok(if is_null {
                BitVec::with_len_all_false(rows)
            } else {
                BitVec::with_len_all_true(rows)
            });
        }
        if nulls == rows {
            return Ok(if is_null {
                BitVec::with_len_all_true(rows)
            } else {
                BitVec::with_len_all_false(rows)
            });
        }
    }

    let chunks = table
        .encoded_chunks(col)
        .ok_or(QueryError::ColumnOutOfBounds { col, column_count: table.column_count() })?;
    let rows = table.row_count();
    let page = table.page_size_rows();

    // Fast paths when we don't have a validity bitmap (no nulls).
    if chunks.iter().all(|c| match c {
        EncodedChunk::Int(c) => c.validity.is_none(),
        EncodedChunk::Float(c) => c.validity.is_none(),
        EncodedChunk::Bool(c) => c.validity.is_none(),
        EncodedChunk::Dict(c) => c.validity.is_none(),
    }) {
        return Ok(if is_null {
            BitVec::with_len_all_false(rows)
        } else {
            BitVec::with_len_all_true(rows)
        });
    }

    let mut out = BitVec::with_capacity_bits(rows);
    for (chunk_idx, chunk) in chunks.iter().enumerate() {
        let base = chunk_idx * page;
        if base >= rows {
            break;
        }
        let chunk_rows = (rows - base).min(chunk.len());

        let validity = match chunk {
            EncodedChunk::Int(c) => c.validity.as_ref(),
            EncodedChunk::Float(c) => c.validity.as_ref(),
            EncodedChunk::Bool(c) => c.validity.as_ref(),
            EncodedChunk::Dict(c) => c.validity.as_ref(),
        };

        match validity {
            None => {
                // No validity bitmap => all values are non-null in this chunk.
                out.extend_constant(!is_null, chunk_rows);
            }
            Some(validity) => {
                if validity.count_ones() == 0 {
                    // All null.
                    out.extend_constant(is_null, chunk_rows);
                    continue;
                }
                if validity.all_true() {
                    out.extend_constant(!is_null, chunk_rows);
                    continue;
                }

                for i in 0..chunk_rows {
                    // Validity bitmap uses `true` for non-null.
                    out.push(if is_null { !validity.get(i) } else { validity.get(i) });
                }
            }
        }
    }

    Ok(out)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum KeyValue {
    Null,
    I64(i64),
    F64(u64),
    Bool(bool),
    Dict(u32),
}

fn canonical_f64_bits(v: f64) -> u64 {
    // Canonicalize `-0.0` to `0.0`, and canonicalize NaNs so they group together.
    if v == 0.0 {
        0.0f64.to_bits()
    } else if v.is_nan() {
        f64::NAN.to_bits()
    } else {
        v.to_bits()
    }
}

#[derive(Clone, Copy, Debug)]
enum Scalar {
    Null,
    I64(i64),
    F64(f64),
    Bool(bool),
    U32(u32),
}

#[derive(Clone, Copy, Debug)]
enum KeyKind {
    Int,
    Float,
    Bool,
    Dict,
}

fn key_kind_for_column_type(column_type: ColumnType) -> Option<KeyKind> {
    match column_type {
        ColumnType::Number => Some(KeyKind::Float),
        ColumnType::String => Some(KeyKind::Dict),
        ColumnType::Boolean => Some(KeyKind::Bool),
        ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {
            Some(KeyKind::Int)
        }
    }
}

fn scalar_to_key(kind: KeyKind, scalar: Scalar) -> KeyValue {
    match scalar {
        Scalar::Null => KeyValue::Null,
        Scalar::I64(v) => match kind {
            KeyKind::Int => KeyValue::I64(v),
            _ => KeyValue::Null,
        },
        Scalar::F64(v) => match kind {
            KeyKind::Float => KeyValue::F64(canonical_f64_bits(v)),
            _ => KeyValue::Null,
        },
        Scalar::Bool(v) => match kind {
            KeyKind::Bool => KeyValue::Bool(v),
            _ => KeyValue::Null,
        },
        Scalar::U32(v) => match kind {
            KeyKind::Dict => KeyValue::Dict(v),
            _ => KeyValue::Null,
        },
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
struct DistinctGroupKey {
    group: u64,
    value: u64,
}

fn distinct_value_bits(kind: KeyKind, scalar: Scalar) -> Option<u64> {
    match (kind, scalar) {
        (_, Scalar::Null) => None,
        (KeyKind::Int, Scalar::I64(v)) => Some(v as u64),
        (KeyKind::Float, Scalar::F64(v)) => Some(canonical_f64_bits(v)),
        (KeyKind::Bool, Scalar::Bool(v)) => Some(if v { 1 } else { 0 }),
        (KeyKind::Dict, Scalar::U32(v)) => Some(v as u64),
        _ => None,
    }
}

fn update_welford(counts: &mut [u64], means: &mut [f64], m2: &mut [f64], group: usize, x: f64) {
    let n0 = counts[group];
    let n1 = n0 + 1;
    counts[group] = n1;
    let mean0 = means[group];
    let delta = x - mean0;
    let mean1 = mean0 + delta / n1 as f64;
    means[group] = mean1;
    let delta2 = x - mean1;
    m2[group] += delta * delta2;
}

fn value_from_i64(column_type: ColumnType, value: i64) -> Value {
    match column_type {
        ColumnType::DateTime => Value::DateTime(value),
        ColumnType::Currency { .. } => Value::Currency(value),
        ColumnType::Percentage { .. } => Value::Percentage(value),
        _ => Value::Number(value as f64),
    }
}

fn default_output_name(table: &ColumnarTable, spec: &AggSpec) -> String {
    let col_name = spec
        .column
        .and_then(|idx| table.schema().get(idx))
        .map(|s| s.name.as_str());

    match (spec.op, col_name) {
        (AggOp::Count, None) => "count".to_owned(),
        (AggOp::Count, Some(name)) => format!("count_{name}"),
        (AggOp::CountNumbers, Some(name)) => format!("count_numbers_{name}"),
        (AggOp::SumF64, Some(name)) => format!("sum_{name}"),
        (AggOp::AvgF64, Some(name)) => format!("avg_{name}"),
        (AggOp::DistinctCount, Some(name)) => format!("distinct_count_{name}"),
        (AggOp::Var, Some(name)) => format!("var_{name}"),
        (AggOp::VarP, Some(name)) => format!("var_p_{name}"),
        (AggOp::StdDev, Some(name)) => format!("std_dev_{name}"),
        (AggOp::StdDevP, Some(name)) => format!("std_dev_p_{name}"),
        (AggOp::Min, Some(name)) => format!("min_{name}"),
        (AggOp::Max, Some(name)) => format!("max_{name}"),
        _ => "agg".to_owned(),
    }
}

#[derive(Default)]
struct FastHasher {
    hash: u64,
}

impl Hasher for FastHasher {
    fn finish(&self) -> u64 {
        self.hash
    }

    fn write(&mut self, bytes: &[u8]) {
        // FNV-1a style mixing for arbitrary bytes.
        const FNV_OFFSET: u64 = 0xcbf29ce484222325;
        const FNV_PRIME: u64 = 0x100000001b3;
        let mut h = if self.hash == 0 { FNV_OFFSET } else { self.hash };
        for b in bytes {
            h ^= *b as u64;
            h = h.wrapping_mul(FNV_PRIME);
        }
        self.hash = h;
    }

    fn write_u8(&mut self, i: u8) {
        self.write_u64(i as u64);
    }

    fn write_u16(&mut self, i: u16) {
        self.write_u64(i as u64);
    }

    fn write_u32(&mut self, i: u32) {
        self.write_u64(i as u64);
    }

    fn write_u64(&mut self, i: u64) {
        // A quick integer mixer (splitmix-ish).
        let mut x = i.wrapping_add(self.hash).wrapping_add(0x9E3779B97F4A7C15);
        x = (x ^ (x >> 30)).wrapping_mul(0xBF58476D1CE4E5B9);
        x = (x ^ (x >> 27)).wrapping_mul(0x94D049BB133111EB);
        self.hash = x ^ (x >> 31);
    }

    fn write_i64(&mut self, i: i64) {
        self.write_u64(i as u64);
    }

    fn write_usize(&mut self, i: usize) {
        self.write_u64(i as u64);
    }
}

type FastBuildHasher = BuildHasherDefault<FastHasher>;
type FastHashMap<K, V> = HashMap<K, V, FastBuildHasher>;

struct U32SeqCursor<'a> {
    pos: u32,
    inner: U32SeqCursorInner<'a>,
}

enum U32SeqCursorInner<'a> {
    Bitpacked { bit_width: u8, data: &'a [u8] },
    Rle {
        values: &'a [u32],
        ends: &'a [u32],
        run: usize,
        run_value: u32,
        run_end: u32,
    },
}

impl<'a> U32SeqCursor<'a> {
    fn new(encoding: &'a U32SequenceEncoding) -> Self {
        match encoding {
            U32SequenceEncoding::Bitpacked { bit_width, data } => Self {
                pos: 0,
                inner: U32SeqCursorInner::Bitpacked {
                    bit_width: *bit_width,
                    data,
                },
            },
            U32SequenceEncoding::Rle(rle) => {
                let (run_value, run_end) = rle
                    .values
                    .first()
                    .copied()
                    .zip(rle.ends.first().copied())
                    .unwrap_or((0, 0));
                Self {
                    pos: 0,
                    inner: U32SeqCursorInner::Rle {
                        values: &rle.values,
                        ends: &rle.ends,
                        run: 0,
                        run_value,
                        run_end,
                    },
                }
            }
        }
    }

    fn next(&mut self) -> u32 {
        let out = match &mut self.inner {
            U32SeqCursorInner::Bitpacked { bit_width, data } => {
                crate::bitpacking::get_u64_at(data, *bit_width, self.pos as usize) as u32
            }
            U32SeqCursorInner::Rle {
                values,
                ends,
                run,
                run_value,
                run_end,
            } => {
                while *run < ends.len() && self.pos >= *run_end {
                    *run += 1;
                    if *run < ends.len() {
                        *run_value = values[*run];
                        *run_end = ends[*run];
                    }
                }
                *run_value
            }
        };
        self.pos = self.pos.saturating_add(1);
        out
    }
}

struct U64SeqCursor<'a> {
    pos: u32,
    inner: U64SeqCursorInner<'a>,
}

enum U64SeqCursorInner<'a> {
    Bitpacked { bit_width: u8, data: &'a [u8] },
    Rle {
        values: &'a [u64],
        ends: &'a [u32],
        run: usize,
        run_value: u64,
        run_end: u32,
    },
}

impl<'a> U64SeqCursor<'a> {
    fn new(encoding: &'a U64SequenceEncoding) -> Self {
        match encoding {
            U64SequenceEncoding::Bitpacked { bit_width, data } => Self {
                pos: 0,
                inner: U64SeqCursorInner::Bitpacked {
                    bit_width: *bit_width,
                    data,
                },
            },
            U64SequenceEncoding::Rle(rle) => {
                let (run_value, run_end) = rle
                    .values
                    .first()
                    .copied()
                    .zip(rle.ends.first().copied())
                    .unwrap_or((0, 0));
                Self {
                    pos: 0,
                    inner: U64SeqCursorInner::Rle {
                        values: &rle.values,
                        ends: &rle.ends,
                        run: 0,
                        run_value,
                        run_end,
                    },
                }
            }
        }
    }

    fn next(&mut self) -> u64 {
        let out = match &mut self.inner {
            U64SeqCursorInner::Bitpacked { bit_width, data } => {
                crate::bitpacking::get_u64_at(data, *bit_width, self.pos as usize)
            }
            U64SeqCursorInner::Rle {
                values,
                ends,
                run,
                run_value,
                run_end,
            } => {
                while *run < ends.len() && self.pos >= *run_end {
                    *run += 1;
                    if *run < ends.len() {
                        *run_value = values[*run];
                        *run_end = ends[*run];
                    }
                }
                *run_value
            }
        };
        self.pos = self.pos.saturating_add(1);
        out
    }
}

enum ScalarChunkCursor<'a> {
    Int {
        min: i64,
        offsets: U64SeqCursor<'a>,
        validity: Option<&'a BitVec>,
        idx: usize,
    },
    Float {
        values: &'a [f64],
        validity: Option<&'a BitVec>,
        idx: usize,
    },
    Bool {
        data: &'a [u8],
        validity: Option<&'a BitVec>,
        idx: usize,
    },
    Dict {
        indices: U32SeqCursor<'a>,
        validity: Option<&'a BitVec>,
        idx: usize,
    },
}

impl<'a> ScalarChunkCursor<'a> {
    fn from_column_chunk(
        col: usize,
        column_type: ColumnType,
        chunk: &'a EncodedChunk,
    ) -> Result<Self, QueryError> {
        match (column_type, chunk) {
            (
                ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. },
                EncodedChunk::Int(c),
            ) => Ok(Self::Int {
                min: c.min,
                offsets: U64SeqCursor::new(&c.offsets),
                validity: c.validity.as_ref(),
                idx: 0,
            }),
            (ColumnType::Number, EncodedChunk::Float(c)) => Ok(Self::Float {
                values: &c.values,
                validity: c.validity.as_ref(),
                idx: 0,
            }),
            (ColumnType::Boolean, EncodedChunk::Bool(c)) => Ok(Self::Bool {
                data: &c.data,
                validity: c.validity.as_ref(),
                idx: 0,
            }),
            (ColumnType::String, EncodedChunk::Dict(c)) => Ok(Self::Dict {
                indices: U32SeqCursor::new(&c.indices),
                validity: c.validity.as_ref(),
                idx: 0,
            }),
            (ty, _) => Err(QueryError::UnsupportedColumnType {
                col,
                column_type: ty,
                operation: "chunk decode",
            }),
        }
    }

    fn next(&mut self) -> Scalar {
        match self {
            Self::Int {
                min,
                offsets,
                validity,
                idx,
            } => {
                let offset = offsets.next();
                let row = *idx;
                *idx += 1;
                if validity.as_ref().is_some_and(|v| !v.get(row)) {
                    return Scalar::Null;
                }
                let value = (*min as i128 + offset as i128) as i64;
                Scalar::I64(value)
            }
            Self::Float {
                values,
                validity,
                idx,
            } => {
                let row = *idx;
                *idx += 1;
                if validity.as_ref().is_some_and(|v| !v.get(row)) {
                    return Scalar::Null;
                }
                Scalar::F64(values[row])
            }
            Self::Bool {
                data,
                validity,
                idx,
            } => {
                let row = *idx;
                *idx += 1;
                if validity.as_ref().is_some_and(|v| !v.get(row)) {
                    return Scalar::Null;
                }
                let byte = data[row / 8];
                let bit = row % 8;
                Scalar::Bool(((byte >> bit) & 1) == 1)
            }
            Self::Dict {
                indices,
                validity,
                idx,
            } => {
                let row = *idx;
                *idx += 1;
                let ix = indices.next();
                if validity.as_ref().is_some_and(|v| !v.get(row)) {
                    return Scalar::Null;
                }
                Scalar::U32(ix)
            }
        }
    }
}

fn scalar_from_column_chunk_at(
    col: usize,
    column_type: ColumnType,
    chunk: &EncodedChunk,
    idx: usize,
) -> Result<Scalar, QueryError> {
    match (column_type, chunk) {
        (
            ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. },
            EncodedChunk::Int(c),
        ) => {
            if c.validity.as_ref().is_some_and(|v| !v.get(idx)) {
                return Ok(Scalar::Null);
            }
            let offset = c.offsets.get(idx);
            Ok(Scalar::I64((c.min as i128 + offset as i128) as i64))
        }
        (ColumnType::Number, EncodedChunk::Float(c)) => {
            if c.validity.as_ref().is_some_and(|v| !v.get(idx)) {
                return Ok(Scalar::Null);
            }
            Ok(Scalar::F64(c.values[idx]))
        }
        (ColumnType::Boolean, EncodedChunk::Bool(c)) => {
            if c.validity.as_ref().is_some_and(|v| !v.get(idx)) {
                return Ok(Scalar::Null);
            }
            let byte = c.data[idx / 8];
            let bit = idx % 8;
            Ok(Scalar::Bool(((byte >> bit) & 1) == 1))
        }
        (ColumnType::String, EncodedChunk::Dict(c)) => {
            if c.validity.as_ref().is_some_and(|v| !v.get(idx)) {
                return Ok(Scalar::Null);
            }
            Ok(Scalar::U32(c.indices.get(idx)))
        }
        (ty, _) => Err(QueryError::UnsupportedColumnType {
            col,
            column_type: ty,
            operation: "chunk scalar lookup",
        }),
    }
}

#[derive(Clone, Debug)]
enum ResultColumn {
    Int { values: Vec<i64>, validity: BitVec },
    Float { values: Vec<f64>, validity: BitVec },
    Bool { values: BitVec, validity: BitVec },
    Dict {
        indices: Vec<u32>,
        validity: BitVec,
        dictionary: Arc<Vec<Arc<str>>>,
    },
}

impl ResultColumn {
    fn len(&self) -> usize {
        match self {
            Self::Int { values, .. } => values.len(),
            Self::Float { values, .. } => values.len(),
            Self::Bool { values, .. } => values.len(),
            Self::Dict { indices, .. } => indices.len(),
        }
    }
}

/// Output of `GROUP BY`.
#[derive(Clone, Debug)]
pub struct GroupByResult {
    schema: Vec<ColumnSchema>,
    columns: Vec<ResultColumn>,
    rows: usize,
}

impl GroupByResult {
    pub fn schema(&self) -> &[ColumnSchema] {
        &self.schema
    }

    pub fn row_count(&self) -> usize {
        self.rows
    }

    pub fn column_count(&self) -> usize {
        self.schema.len()
    }

    pub fn to_values(&self) -> Vec<Vec<Value>> {
        let mut out: Vec<Vec<Value>> = Vec::new();
        let _ = out.try_reserve_exact(self.columns.len());
        for (col_idx, column) in self.columns.iter().enumerate() {
            let column_type = self.schema.get(col_idx).map(|s| s.column_type);
            let mut values: Vec<Value> = Vec::new();
            let _ = values.try_reserve_exact(self.rows);
            match (column, column_type) {
                (ResultColumn::Float { values: v, validity }, Some(_)) => {
                    for i in 0..self.rows {
                        if !validity.get(i) {
                            values.push(Value::Null);
                        } else {
                            values.push(Value::Number(v[i]));
                        }
                    }
                }
                (ResultColumn::Int { values: v, validity }, Some(ty)) => {
                    for i in 0..self.rows {
                        if !validity.get(i) {
                            values.push(Value::Null);
                        } else {
                            values.push(value_from_i64(ty, v[i]));
                        }
                    }
                }
                (ResultColumn::Bool { values: v, validity }, Some(_)) => {
                    for i in 0..self.rows {
                        if !validity.get(i) {
                            values.push(Value::Null);
                        } else {
                            values.push(Value::Boolean(v.get(i)));
                        }
                    }
                }
                (ResultColumn::Dict { indices, validity, dictionary }, Some(_)) => {
                    for i in 0..self.rows {
                        if !validity.get(i) {
                            values.push(Value::Null);
                        } else {
                            let idx = indices[i] as usize;
                            values.push(
                                dictionary
                                    .get(idx)
                                    .cloned()
                                    .map(Value::String)
                                    .unwrap_or(Value::Null),
                            );
                        }
                    }
                }
                _ => {
                    for _ in 0..self.rows {
                        values.push(Value::Null);
                    }
                }
            }
            out.push(values);
        }
        out
    }

    pub fn to_table(&self, options: TableOptions) -> ColumnarTable {
        let mut builder = ColumnarTableBuilder::new(self.schema.clone(), options);
        let mut row: Vec<Value> = vec![Value::Null; self.columns.len()];
        for r in 0..self.rows {
            for c in 0..self.columns.len() {
                row[c] = self.get_value(r, c);
            }
            builder.append_row(&row);
        }
        builder.finalize()
    }

    fn get_value(&self, row: usize, col: usize) -> Value {
        let column_type = self.schema.get(col).map(|s| s.column_type);
        match (self.columns.get(col), column_type) {
            (Some(ResultColumn::Float { values, validity }), Some(_)) => {
                if !validity.get(row) {
                    Value::Null
                } else {
                    Value::Number(values[row])
                }
            }
            (Some(ResultColumn::Int { values, validity }), Some(ty)) => {
                if !validity.get(row) {
                    Value::Null
                } else {
                    value_from_i64(ty, values[row])
                }
            }
            (Some(ResultColumn::Bool { values, validity }), Some(_)) => {
                if !validity.get(row) {
                    Value::Null
                } else {
                    Value::Boolean(values.get(row))
                }
            }
            (Some(ResultColumn::Dict { indices, validity, dictionary }), Some(_)) => {
                if !validity.get(row) {
                    Value::Null
                } else {
                    dictionary
                        .get(indices[row] as usize)
                        .cloned()
                        .map(Value::String)
                        .unwrap_or(Value::Null)
                }
            }
            _ => Value::Null,
        }
    }
}

enum KeyColumnBuilder {
    Int { values: Vec<i64>, validity: BitVec },
    Float { values: Vec<f64>, validity: BitVec },
    Bool { values: BitVec, validity: BitVec },
    Dict {
        indices: Vec<u32>,
        validity: BitVec,
        dictionary: Arc<Vec<Arc<str>>>,
    },
}

impl KeyColumnBuilder {
    fn new(
        col: usize,
        column_type: ColumnType,
        dict: Option<Arc<Vec<Arc<str>>>>,
    ) -> Result<Self, QueryError> {
        Ok(match column_type {
            ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {
                Self::Int {
                    values: Vec::new(),
                    validity: BitVec::new(),
                }
            }
            ColumnType::Number => Self::Float {
                values: Vec::new(),
                validity: BitVec::new(),
            },
            ColumnType::Boolean => Self::Bool {
                values: BitVec::new(),
                validity: BitVec::new(),
            },
            ColumnType::String => Self::Dict {
                indices: Vec::new(),
                validity: BitVec::new(),
                dictionary: dict.ok_or(QueryError::MissingDictionary { col })?,
            },
        })
    }

    fn push(&mut self, scalar: Scalar) {
        match (self, scalar) {
            (Self::Int { values, validity }, Scalar::I64(v)) => {
                values.push(v);
                validity.push(true);
            }
            (Self::Int { values, validity }, Scalar::Null) => {
                values.push(0);
                validity.push(false);
            }
            (Self::Float { values, validity }, Scalar::F64(v)) => {
                values.push(v);
                validity.push(true);
            }
            (Self::Float { values, validity }, Scalar::Null) => {
                values.push(0.0);
                validity.push(false);
            }
            (Self::Bool { values, validity }, Scalar::Bool(v)) => {
                values.push(v);
                validity.push(true);
            }
            (Self::Bool { values, validity }, Scalar::Null) => {
                values.push(false);
                validity.push(false);
            }
            (Self::Dict { indices, validity, .. }, Scalar::U32(v)) => {
                indices.push(v);
                validity.push(true);
            }
            (Self::Dict { indices, validity, .. }, Scalar::Null) => {
                indices.push(0);
                validity.push(false);
            }
            _ => {
                // Type mismatch should be prevented by planning.
                debug_assert!(false, "KeyColumnBuilder scalar type mismatch");
            }
        }
    }

    fn finish(self) -> ResultColumn {
        match self {
            Self::Int { values, validity } => ResultColumn::Int { values, validity },
            Self::Float { values, validity } => ResultColumn::Float { values, validity },
            Self::Bool { values, validity } => ResultColumn::Bool { values, validity },
            Self::Dict {
                indices,
                validity,
                dictionary,
            } => ResultColumn::Dict {
                indices,
                validity,
                dictionary,
            },
        }
    }
}

enum AggState {
    CountRows { counts: Vec<u64> },
    CountNonNull { counts: Vec<u64>, col: usize },
    CountNumbers { counts: Vec<u64>, col: usize },
    SumF64 {
        sums: Vec<f64>,
        non_null: Vec<u64>,
        col: usize,
    },
    AvgF64 {
        sums: Vec<f64>,
        counts: Vec<u64>,
        col: usize,
    },
    DistinctCount {
        counts: Vec<u64>,
        seen: FastHashMap<DistinctGroupKey, ()>,
        col: usize,
        kind: KeyKind,
    },
    Var {
        counts: Vec<u64>,
        means: Vec<f64>,
        m2: Vec<f64>,
        col: usize,
    },
    VarP {
        counts: Vec<u64>,
        means: Vec<f64>,
        m2: Vec<f64>,
        col: usize,
    },
    StdDev {
        counts: Vec<u64>,
        means: Vec<f64>,
        m2: Vec<f64>,
        col: usize,
    },
    StdDevP {
        counts: Vec<u64>,
        means: Vec<f64>,
        m2: Vec<f64>,
        col: usize,
    },
    MinI64 {
        values: Vec<i64>,
        validity: BitVec,
        col: usize,
    },
    MaxI64 {
        values: Vec<i64>,
        validity: BitVec,
        col: usize,
    },
    MinF64 {
        values: Vec<f64>,
        validity: BitVec,
        col: usize,
    },
    MaxF64 {
        values: Vec<f64>,
        validity: BitVec,
        col: usize,
    },
    MinBool {
        values: BitVec,
        validity: BitVec,
        col: usize,
    },
    MaxBool {
        values: BitVec,
        validity: BitVec,
        col: usize,
    },
}

impl AggState {
    fn input_col(&self) -> Option<usize> {
        match self {
            Self::CountRows { .. } => None,
            Self::CountNonNull { col, .. } => Some(*col),
            Self::CountNumbers { col, .. } => Some(*col),
            Self::SumF64 { col, .. } => Some(*col),
            Self::AvgF64 { col, .. } => Some(*col),
            Self::DistinctCount { col, .. } => Some(*col),
            Self::Var { col, .. }
            | Self::VarP { col, .. }
            | Self::StdDev { col, .. }
            | Self::StdDevP { col, .. } => Some(*col),
            Self::MinI64 { col, .. } => Some(*col),
            Self::MaxI64 { col, .. } => Some(*col),
            Self::MinF64 { col, .. } => Some(*col),
            Self::MaxF64 { col, .. } => Some(*col),
            Self::MinBool { col, .. } => Some(*col),
            Self::MaxBool { col, .. } => Some(*col),
        }
    }

    fn push_group(&mut self) {
        match self {
            Self::CountRows { counts } => counts.push(0),
            Self::CountNonNull { counts, .. } => counts.push(0),
            Self::CountNumbers { counts, .. } => counts.push(0),
            Self::SumF64 { sums, non_null, .. } => {
                sums.push(0.0);
                non_null.push(0);
            }
            Self::AvgF64 { sums, counts, .. } => {
                sums.push(0.0);
                counts.push(0);
            }
            Self::DistinctCount { counts, .. } => counts.push(0),
            Self::Var {
                counts,
                means,
                m2,
                ..
            }
            | Self::VarP {
                counts,
                means,
                m2,
                ..
            }
            | Self::StdDev {
                counts,
                means,
                m2,
                ..
            }
            | Self::StdDevP {
                counts,
                means,
                m2,
                ..
            } => {
                counts.push(0);
                means.push(0.0);
                m2.push(0.0);
            }
            Self::MinI64 { values, validity, .. } | Self::MaxI64 { values, validity, .. } => {
                values.push(0);
                validity.push(false);
            }
            Self::MinF64 { values, validity, .. } | Self::MaxF64 { values, validity, .. } => {
                values.push(0.0);
                validity.push(false);
            }
            Self::MinBool { values, validity, .. } | Self::MaxBool { values, validity, .. } => {
                values.push(false);
                validity.push(false);
            }
        }
    }

    fn update_count_row(&mut self, group: usize) {
        match self {
            Self::CountRows { counts } => counts[group] += 1,
            _ => {}
        }
    }

    fn update_from_scalar(&mut self, group: usize, scalar: Scalar) {
        match self {
            Self::CountNonNull { counts, .. } => {
                if !matches!(scalar, Scalar::Null) {
                    counts[group] += 1;
                }
            }
            Self::CountNumbers { counts, .. } => match scalar {
                Scalar::F64(_) | Scalar::I64(_) => {
                    counts[group] += 1;
                }
                Scalar::Null | Scalar::Bool(_) | Scalar::U32(_) => {}
            },
            Self::SumF64 { sums, non_null, .. } => match scalar {
                Scalar::F64(v) => {
                    sums[group] += v;
                    non_null[group] += 1;
                }
                Scalar::I64(v) => {
                    sums[group] += v as f64;
                    non_null[group] += 1;
                }
                Scalar::Bool(v) => {
                    sums[group] += if v { 1.0 } else { 0.0 };
                    non_null[group] += 1;
                }
                Scalar::Null | Scalar::U32(_) => {}
            },
            Self::AvgF64 { sums, counts, .. } => match scalar {
                Scalar::F64(v) => {
                    sums[group] += v;
                    counts[group] += 1;
                }
                Scalar::I64(v) => {
                    sums[group] += v as f64;
                    counts[group] += 1;
                }
                // Averages are only planned for numeric columns, but keep defensive behavior.
                Scalar::Null | Scalar::Bool(_) | Scalar::U32(_) => {}
            },
            Self::DistinctCount {
                counts,
                seen,
                kind,
                ..
            } => {
                let kind = *kind;
                let Some(value_bits) = distinct_value_bits(kind, scalar) else {
                    return;
                };
                let key = DistinctGroupKey {
                    group: group as u64,
                    value: value_bits,
                };
                if seen.insert(key, ()).is_none() {
                    counts[group] += 1;
                }
            }
            Self::Var {
                counts,
                means,
                m2,
                ..
            }
            | Self::VarP {
                counts,
                means,
                m2,
                ..
            }
            | Self::StdDev {
                counts,
                means,
                m2,
                ..
            }
            | Self::StdDevP {
                counts,
                means,
                m2,
                ..
            } => {
                let x = match scalar {
                    Scalar::F64(v) => Some(v),
                    Scalar::I64(v) => Some(v as f64),
                    _ => None,
                };
                if let Some(x) = x {
                    update_welford(counts, means, m2, group, x);
                }
            }
            Self::MinI64 { values, validity, .. } => match scalar {
                Scalar::I64(v) => {
                    if !validity.get(group) {
                        values[group] = v;
                        validity.set(group, true);
                    } else {
                        values[group] = values[group].min(v);
                    }
                }
                _ => {}
            },
            Self::MaxI64 { values, validity, .. } => match scalar {
                Scalar::I64(v) => {
                    if !validity.get(group) {
                        values[group] = v;
                        validity.set(group, true);
                    } else {
                        values[group] = values[group].max(v);
                    }
                }
                _ => {}
            },
            Self::MinF64 { values, validity, .. } => match scalar {
                Scalar::F64(v) => {
                    if !validity.get(group) {
                        values[group] = v;
                        validity.set(group, true);
                    } else if v.total_cmp(&values[group]).is_lt() {
                        values[group] = v;
                    }
                }
                _ => {}
            },
            Self::MaxF64 { values, validity, .. } => match scalar {
                Scalar::F64(v) => {
                    if !validity.get(group) {
                        values[group] = v;
                        validity.set(group, true);
                    } else if v.total_cmp(&values[group]).is_gt() {
                        values[group] = v;
                    }
                }
                _ => {}
            },
            Self::MinBool { values, validity, .. } => match scalar {
                Scalar::Bool(v) => {
                    if !validity.get(group) {
                        values.set(group, v);
                        validity.set(group, true);
                    } else if !v {
                        values.set(group, false);
                    }
                }
                _ => {}
            },
            Self::MaxBool { values, validity, .. } => match scalar {
                Scalar::Bool(v) => {
                    if !validity.get(group) {
                        values.set(group, v);
                        validity.set(group, true);
                    } else if v {
                        values.set(group, true);
                    }
                }
                _ => {}
            },
            Self::CountRows { .. } => {}
        }
    }

    fn finish(self) -> ResultColumn {
        match self {
            Self::CountRows { counts }
            | Self::CountNonNull { counts, .. }
            | Self::CountNumbers { counts, .. }
            | Self::DistinctCount { counts, .. } => {
                let mut validity = BitVec::new();
                let values: Vec<f64> = counts
                    .into_iter()
                    .map(|c| {
                        validity.push(true);
                        c as f64
                    })
                    .collect();
                ResultColumn::Float { values, validity }
            }
            Self::SumF64 { sums, non_null, .. } => {
                let mut validity = BitVec::with_capacity_bits(sums.len());
                for &cnt in &non_null {
                    validity.push(cnt > 0);
                }
                ResultColumn::Float {
                    values: sums,
                    validity,
                }
            }
            Self::AvgF64 { mut sums, counts, .. } => {
                let mut validity = BitVec::with_capacity_bits(sums.len());
                for (i, &cnt) in counts.iter().enumerate() {
                    if cnt == 0 {
                        sums[i] = 0.0;
                        validity.push(false);
                    } else {
                        sums[i] /= cnt as f64;
                        validity.push(true);
                    }
                }
                ResultColumn::Float {
                    values: sums,
                    validity,
                }
            }
            Self::Var { counts, mut m2, .. } => {
                let mut validity = BitVec::with_capacity_bits(m2.len());
                for (i, &cnt) in counts.iter().enumerate() {
                    if cnt > 1 {
                        m2[i] /= (cnt - 1) as f64;
                        validity.push(true);
                    } else {
                        m2[i] = 0.0;
                        validity.push(false);
                    }
                }
                ResultColumn::Float {
                    values: m2,
                    validity,
                }
            }
            Self::VarP { counts, mut m2, .. } => {
                let mut validity = BitVec::with_capacity_bits(m2.len());
                for (i, &cnt) in counts.iter().enumerate() {
                    if cnt > 0 {
                        m2[i] /= cnt as f64;
                        validity.push(true);
                    } else {
                        m2[i] = 0.0;
                        validity.push(false);
                    }
                }
                ResultColumn::Float {
                    values: m2,
                    validity,
                }
            }
            Self::StdDev { counts, mut m2, .. } => {
                let mut validity = BitVec::with_capacity_bits(m2.len());
                for (i, &cnt) in counts.iter().enumerate() {
                    if cnt > 1 {
                        m2[i] = (m2[i] / (cnt - 1) as f64).sqrt();
                        validity.push(true);
                    } else {
                        m2[i] = 0.0;
                        validity.push(false);
                    }
                }
                ResultColumn::Float {
                    values: m2,
                    validity,
                }
            }
            Self::StdDevP { counts, mut m2, .. } => {
                let mut validity = BitVec::with_capacity_bits(m2.len());
                for (i, &cnt) in counts.iter().enumerate() {
                    if cnt > 0 {
                        m2[i] = (m2[i] / cnt as f64).sqrt();
                        validity.push(true);
                    } else {
                        m2[i] = 0.0;
                        validity.push(false);
                    }
                }
                ResultColumn::Float {
                    values: m2,
                    validity,
                }
            }
            Self::MinI64 { values, validity, .. } | Self::MaxI64 { values, validity, .. } => {
                ResultColumn::Int { values, validity }
            }
            Self::MinF64 { values, validity, .. } | Self::MaxF64 { values, validity, .. } => {
                ResultColumn::Float { values, validity }
            }
            Self::MinBool { values, validity, .. } | Self::MaxBool { values, validity, .. } => {
                ResultColumn::Bool { values, validity }
            }
        }
    }
}

struct AggColumnPlan {
    col: usize,
    agg_indices: Vec<usize>,
    from_key_pos: Option<usize>,
}

/// A streaming `GROUP BY` engine.
///
/// `consume_chunks` lets callers process very large tables incrementally (page-by-page)
/// without decoding entire columns.
pub struct GroupByEngine {
    schema: Vec<ColumnSchema>,
    key_cols: Vec<usize>,
    key_kinds: Vec<KeyKind>,
    key_builders: Vec<KeyColumnBuilder>,
    agg_states: Vec<AggState>,
    count_row_aggs: Vec<usize>,
    agg_plans: Vec<AggColumnPlan>,
    groups: FastHashMap<Box<[KeyValue]>, usize>,
    scratch_keys: Vec<KeyValue>,
    scratch_key_scalars: Vec<Scalar>,
    groups_len: usize,
}

impl GroupByEngine {
    pub fn new(table: &ColumnarTable, keys: &[usize], aggs: &[AggSpec]) -> Result<Self, QueryError> {
        if keys.is_empty() {
            return Err(QueryError::EmptyKeys);
        }
        let column_count = table.column_count();
        for &k in keys {
            if k >= column_count {
                return Err(QueryError::ColumnOutOfBounds { col: k, column_count });
            }
        }
        for spec in aggs {
            if let Some(col) = spec.column {
                if col >= column_count {
                    return Err(QueryError::ColumnOutOfBounds { col, column_count });
                }
            }
        }

        let mut schema: Vec<ColumnSchema> = Vec::new();
        let _ = schema.try_reserve_exact(keys.len() + aggs.len());
        let mut key_cols: Vec<usize> = Vec::new();
        let _ = key_cols.try_reserve_exact(keys.len());
        let mut key_builders: Vec<KeyColumnBuilder> = Vec::new();
        let _ = key_builders.try_reserve_exact(keys.len());
        let mut key_kinds: Vec<KeyKind> = Vec::new();
        let _ = key_kinds.try_reserve_exact(keys.len());
        for &key_col in keys {
            let col_schema = table.schema()[key_col].clone();
            let kind = key_kind_for_column_type(col_schema.column_type).ok_or(
                QueryError::UnsupportedColumnType {
                    col: key_col,
                    column_type: col_schema.column_type,
                    operation: "GROUP BY key",
                },
            )?;
            let dict = if col_schema.column_type == ColumnType::String {
                Some(table.dictionary(key_col).ok_or(QueryError::MissingDictionary { col: key_col })?)
            } else {
                None
            };
            key_kinds.push(kind);
            key_builders.push(KeyColumnBuilder::new(key_col, col_schema.column_type, dict)?);
            key_cols.push(key_col);
            schema.push(col_schema);
        }

        let mut agg_states: Vec<AggState> = Vec::new();
        let _ = agg_states.try_reserve_exact(aggs.len());
        let mut count_row_aggs: Vec<usize> = Vec::new();
        for (idx, spec) in aggs.iter().enumerate() {
            let name = spec
                .name
                .clone()
                .unwrap_or_else(|| default_output_name(table, spec));

            match spec.op {
                AggOp::Count => {
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    match spec.column {
                        None => {
                            agg_states.push(AggState::CountRows { counts: Vec::new() });
                            count_row_aggs.push(idx);
                        }
                        Some(col) => {
                            agg_states.push(AggState::CountNonNull {
                                counts: Vec::new(),
                                col,
                            });
                        }
                    }
                }
                AggOp::CountNumbers => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "COUNTNUMBERS without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    match ty {
                        ColumnType::Number
                        | ColumnType::DateTime
                        | ColumnType::Currency { .. }
                        | ColumnType::Percentage { .. } => {}
                        ColumnType::String | ColumnType::Boolean => {
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "COUNTNUMBERS",
                            });
                        }
                    }
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    agg_states.push(AggState::CountNumbers {
                        counts: Vec::new(),
                        col,
                    });
                }
                AggOp::SumF64 => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "SUM without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    match ty {
                        ColumnType::Number
                        | ColumnType::DateTime
                        | ColumnType::Currency { .. }
                        | ColumnType::Percentage { .. }
                        | ColumnType::Boolean => {}
                        ColumnType::String => {
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "SUM",
                            });
                        }
                    }
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    agg_states.push(AggState::SumF64 {
                        sums: Vec::new(),
                        non_null: Vec::new(),
                        col,
                    });
                }
                AggOp::AvgF64 => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "AVG without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    match ty {
                        ColumnType::Number
                        | ColumnType::DateTime
                        | ColumnType::Currency { .. }
                        | ColumnType::Percentage { .. } => {}
                        ColumnType::String | ColumnType::Boolean => {
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "AVG",
                            });
                        }
                    }
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    agg_states.push(AggState::AvgF64 {
                        sums: Vec::new(),
                        counts: Vec::new(),
                        col,
                    });
                }
                AggOp::DistinctCount => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "DISTINCTCOUNT without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    let kind = key_kind_for_column_type(ty).ok_or(QueryError::UnsupportedColumnType {
                        col,
                        column_type: ty,
                        operation: "DISTINCTCOUNT",
                    })?;
                    let seen_capacity = table
                        .scan()
                        .stats(col)
                        .map(|s| s.distinct_count as usize)
                        .unwrap_or(0)
                        .min(table.row_count());
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    agg_states.push(AggState::DistinctCount {
                        counts: Vec::new(),
                        seen: FastHashMap::with_capacity_and_hasher(
                            seen_capacity,
                            FastBuildHasher::default(),
                        ),
                        col,
                        kind,
                    });
                }
                AggOp::Var | AggOp::VarP | AggOp::StdDev | AggOp::StdDevP => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "VAR/STDDEV without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    match ty {
                        ColumnType::Number
                        | ColumnType::DateTime
                        | ColumnType::Currency { .. }
                        | ColumnType::Percentage { .. } => {}
                        ColumnType::String | ColumnType::Boolean => {
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "VAR/STDDEV",
                            });
                        }
                    }
                    schema.push(ColumnSchema {
                        name,
                        column_type: ColumnType::Number,
                    });
                    let state = match spec.op {
                        AggOp::Var => AggState::Var {
                            counts: Vec::new(),
                            means: Vec::new(),
                            m2: Vec::new(),
                            col,
                        },
                        AggOp::VarP => AggState::VarP {
                            counts: Vec::new(),
                            means: Vec::new(),
                            m2: Vec::new(),
                            col,
                        },
                        AggOp::StdDev => AggState::StdDev {
                            counts: Vec::new(),
                            means: Vec::new(),
                            m2: Vec::new(),
                            col,
                        },
                        AggOp::StdDevP => AggState::StdDevP {
                            counts: Vec::new(),
                            means: Vec::new(),
                            m2: Vec::new(),
                            col,
                        },
                        other => {
                            debug_assert!(
                                false,
                                "unexpected agg op in VAR/STDDEV branch: {other:?}",
                            );
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "VAR/STDDEV",
                            });
                        }
                    };
                    agg_states.push(state);
                }
                AggOp::Min | AggOp::Max => {
                    let col = spec.column.ok_or(QueryError::UnsupportedColumnType {
                        col: 0,
                        column_type: ColumnType::String,
                        operation: "MIN/MAX without column",
                    })?;
                    let ty = table.schema()[col].column_type;
                    match ty {
                        ColumnType::String => {
                            return Err(QueryError::UnsupportedColumnType {
                                col,
                                column_type: ty,
                                operation: "MIN/MAX",
                            });
                        }
                        ColumnType::Number => {
                            schema.push(ColumnSchema { name, column_type: ty });
                            agg_states.push(if spec.op == AggOp::Min {
                                AggState::MinF64 {
                                    values: Vec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            } else {
                                AggState::MaxF64 {
                                    values: Vec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            });
                        }
                        ColumnType::Boolean => {
                            schema.push(ColumnSchema { name, column_type: ty });
                            agg_states.push(if spec.op == AggOp::Min {
                                AggState::MinBool {
                                    values: BitVec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            } else {
                                AggState::MaxBool {
                                    values: BitVec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            });
                        }
                        ColumnType::DateTime | ColumnType::Currency { .. } | ColumnType::Percentage { .. } => {
                            schema.push(ColumnSchema { name, column_type: ty });
                            agg_states.push(if spec.op == AggOp::Min {
                                AggState::MinI64 {
                                    values: Vec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            } else {
                                AggState::MaxI64 {
                                    values: Vec::new(),
                                    validity: BitVec::new(),
                                    col,
                                }
                            });
                        }
                    }
                }
            }
        }

        // Group aggs by input column for more cache-friendly scans.
        let mut key_pos_by_col: FastHashMap<usize, usize> = FastHashMap::default();
        for (pos, &col) in keys.iter().enumerate() {
            key_pos_by_col.insert(col, pos);
        }

        let mut by_col: FastHashMap<usize, Vec<usize>> = FastHashMap::default();
        for (agg_idx, state) in agg_states.iter().enumerate() {
            if let Some(col) = state.input_col() {
                by_col.entry(col).or_default().push(agg_idx);
            }
        }

        let mut agg_plans: Vec<AggColumnPlan> = Vec::new();
        let _ = agg_plans.try_reserve_exact(by_col.len());
        for (col, agg_indices) in by_col {
            let from_key_pos = key_pos_by_col.get(&col).copied();
            agg_plans.push(AggColumnPlan {
                col,
                agg_indices,
                from_key_pos,
            });
        }

        // A rough initial capacity estimate helps avoid early rehashing for common cases.
        let capacity_hint = table
            .scan()
            .stats(keys[0])
            .map(|s| s.distinct_count as usize)
            .unwrap_or(0)
            .min(table.row_count());

        Ok(Self {
            schema,
            key_cols,
            key_kinds,
            key_builders,
            agg_states,
            count_row_aggs,
            agg_plans,
            groups: FastHashMap::with_capacity_and_hasher(capacity_hint, FastBuildHasher::default()),
            scratch_keys: vec![KeyValue::Null; keys.len()],
            scratch_key_scalars: vec![Scalar::Null; keys.len()],
            groups_len: 0,
        })
    }

    pub fn consume_chunks(
        &mut self,
        table: &ColumnarTable,
        chunk_start: usize,
        chunk_end: usize,
    ) -> Result<(), QueryError> {
        let rows = table.row_count();
        let page = table.page_size_rows();
        let chunk_count = (rows + page - 1) / page;
        let chunk_end = chunk_end.min(chunk_count);

        for chunk_idx in chunk_start..chunk_end {
            let base = chunk_idx * page;
            if base >= rows {
                break;
            }
            let chunk_rows = (rows - base).min(page);

            let mut key_cursors: Vec<ScalarChunkCursor<'_>> = Vec::new();
            let _ = key_cursors.try_reserve_exact(self.key_kinds.len());
            for &col_idx in &self.key_cols {
                let chunks = table
                    .encoded_chunks(col_idx)
                    .ok_or(QueryError::ColumnOutOfBounds { col: col_idx, column_count: table.column_count() })?;
                let chunk = chunks.get(chunk_idx).ok_or(QueryError::ColumnOutOfBounds {
                    col: col_idx,
                    column_count: table.column_count(),
                })?;
                let ty = table.schema()[col_idx].column_type;
                key_cursors.push(ScalarChunkCursor::from_column_chunk(col_idx, ty, chunk)?);
            }

            let mut agg_cursors: Vec<Option<ScalarChunkCursor<'_>>> = Vec::new();
            let _ = agg_cursors.try_reserve_exact(self.agg_plans.len());
            for plan in &self.agg_plans {
                if plan.from_key_pos.is_some() {
                    agg_cursors.push(None);
                    continue;
                }
                let chunks = table.encoded_chunks(plan.col).ok_or(QueryError::ColumnOutOfBounds {
                    col: plan.col,
                    column_count: table.column_count(),
                })?;
                let chunk = chunks.get(chunk_idx).ok_or(QueryError::ColumnOutOfBounds {
                    col: plan.col,
                    column_count: table.column_count(),
                })?;
                let ty = table.schema()[plan.col].column_type;
                agg_cursors.push(Some(ScalarChunkCursor::from_column_chunk(plan.col, ty, chunk)?));
            }

            for _row_in_chunk in 0..chunk_rows {
                for (pos, cursor) in key_cursors.iter_mut().enumerate() {
                    let scalar = cursor.next();
                    self.scratch_key_scalars[pos] = scalar;
                    self.scratch_keys[pos] = scalar_to_key(self.key_kinds[pos], scalar);
                }

                let group_idx = if let Some(&idx) = self.groups.get(self.scratch_keys.as_slice()) {
                    idx
                } else {
                    let idx = self.groups_len;
                    self.groups_len += 1;
                    self.groups
                        .insert(self.scratch_keys.to_vec().into_boxed_slice(), idx);
                    for (pos, builder) in self.key_builders.iter_mut().enumerate() {
                        builder.push(self.scratch_key_scalars[pos]);
                    }
                    for state in &mut self.agg_states {
                        state.push_group();
                    }
                    idx
                };

                for &agg_idx in &self.count_row_aggs {
                    self.agg_states[agg_idx].update_count_row(group_idx);
                }

                for (plan_idx, plan) in self.agg_plans.iter().enumerate() {
                    let scalar = if let Some(key_pos) = plan.from_key_pos {
                        self.scratch_key_scalars[key_pos]
                    } else {
                        agg_cursors
                            .get_mut(plan_idx)
                            .and_then(|cursor| cursor.as_mut())
                            .ok_or(QueryError::InternalInvariant(
                                "cursor missing for non-key agg plan",
                            ))?
                            .next()
                    };
                    for &agg_idx in &plan.agg_indices {
                        self.agg_states[agg_idx].update_from_scalar(group_idx, scalar);
                    }
                }
            }
        }

        Ok(())
    }

    pub fn consume_rows(&mut self, table: &ColumnarTable, rows: &[usize]) -> Result<(), QueryError> {
        let row_count = table.row_count();
        let page = table.page_size_rows();

        // Gather chunk slices for each key column once.
        let mut key_chunks: Vec<&[EncodedChunk]> = Vec::new();
        let _ = key_chunks.try_reserve_exact(self.key_cols.len());
        for &col_idx in &self.key_cols {
            let chunks = table.encoded_chunks(col_idx).ok_or(QueryError::ColumnOutOfBounds {
                col: col_idx,
                column_count: table.column_count(),
            })?;
            key_chunks.push(chunks);
        }

        // Gather chunk slices for each agg plan column once.
        let mut agg_chunks: Vec<Option<&[EncodedChunk]>> = Vec::new();
        let _ = agg_chunks.try_reserve_exact(self.agg_plans.len());
        for plan in &self.agg_plans {
            if plan.from_key_pos.is_some() {
                agg_chunks.push(None);
                continue;
            }
            let chunks = table.encoded_chunks(plan.col).ok_or(QueryError::ColumnOutOfBounds {
                col: plan.col,
                column_count: table.column_count(),
            })?;
            agg_chunks.push(Some(chunks));
        }

        let mut current_chunk_idx = usize::MAX;
        let mut key_chunk_refs: Vec<&EncodedChunk> = Vec::new();
        let _ = key_chunk_refs.try_reserve_exact(self.key_cols.len());
        let mut agg_chunk_refs: Vec<Option<&EncodedChunk>> = Vec::new();
        let _ = agg_chunk_refs.try_reserve_exact(self.agg_plans.len());

        for &row in rows {
            if row >= row_count {
                return Err(QueryError::RowOutOfBounds { row, row_count });
            }

            let chunk_idx = row / page;
            let row_in_chunk = row % page;

            if chunk_idx != current_chunk_idx {
                current_chunk_idx = chunk_idx;

                key_chunk_refs.clear();
                for chunks in &key_chunks {
                    let chunk = chunks
                        .get(chunk_idx)
                        .ok_or(QueryError::RowOutOfBounds { row, row_count })?;
                    key_chunk_refs.push(chunk);
                }

                agg_chunk_refs.clear();
                for (plan_idx, plan) in self.agg_plans.iter().enumerate() {
                    if plan.from_key_pos.is_some() {
                        agg_chunk_refs.push(None);
                        continue;
                    }
                    let chunks = agg_chunks
                        .get(plan_idx)
                        .and_then(|chunks| chunks.as_ref().copied())
                        .ok_or(QueryError::InternalInvariant(
                            "agg chunks missing for non-key agg plan",
                        ))?;
                    let chunk = chunks
                        .get(chunk_idx)
                        .ok_or(QueryError::RowOutOfBounds { row, row_count })?;
                    agg_chunk_refs.push(Some(chunk));
                }
            }

            for pos in 0..self.key_cols.len() {
                let col_idx = self.key_cols[pos];
                let ty = table.schema()[col_idx].column_type;
                let scalar = scalar_from_column_chunk_at(
                    col_idx,
                    ty,
                    key_chunk_refs[pos],
                    row_in_chunk,
                )?;
                self.scratch_key_scalars[pos] = scalar;
                self.scratch_keys[pos] = scalar_to_key(self.key_kinds[pos], scalar);
            }

            let group_idx = if let Some(&idx) = self.groups.get(self.scratch_keys.as_slice()) {
                idx
            } else {
                let idx = self.groups_len;
                self.groups_len += 1;
                self.groups
                    .insert(self.scratch_keys.to_vec().into_boxed_slice(), idx);
                for (pos, builder) in self.key_builders.iter_mut().enumerate() {
                    builder.push(self.scratch_key_scalars[pos]);
                }
                for state in &mut self.agg_states {
                    state.push_group();
                }
                idx
            };

            for &agg_idx in &self.count_row_aggs {
                self.agg_states[agg_idx].update_count_row(group_idx);
            }

            for (plan_idx, plan) in self.agg_plans.iter().enumerate() {
                let scalar = if let Some(key_pos) = plan.from_key_pos {
                    self.scratch_key_scalars[key_pos]
                } else {
                    let chunk = agg_chunk_refs
                        .get(plan_idx)
                        .and_then(|chunk| chunk.as_ref().copied())
                        .ok_or(QueryError::InternalInvariant(
                            "chunk missing for non-key agg plan",
                        ))?;
                    let ty = table.schema()[plan.col].column_type;
                    scalar_from_column_chunk_at(plan.col, ty, chunk, row_in_chunk)?
                };
                for &agg_idx in &plan.agg_indices {
                    self.agg_states[agg_idx].update_from_scalar(group_idx, scalar);
                }
            }
        }

        Ok(())
    }

    pub fn consume_mask(&mut self, table: &ColumnarTable, mask: &BitVec) -> Result<(), QueryError> {
        if mask.len() != table.row_count() {
            return Err(QueryError::InternalInvariant(
                "filter mask length must match table row count",
            ));
        }

        if mask.all_true() {
            return self.consume_all(table);
        }

        let row_count = table.row_count();
        let page = table.page_size_rows();

        // Gather chunk slices for each key column once.
        let mut key_chunks: Vec<&[EncodedChunk]> = Vec::new();
        let _ = key_chunks.try_reserve_exact(self.key_cols.len());
        for &col_idx in &self.key_cols {
            let chunks = table.encoded_chunks(col_idx).ok_or(QueryError::ColumnOutOfBounds {
                col: col_idx,
                column_count: table.column_count(),
            })?;
            key_chunks.push(chunks);
        }

        // Gather chunk slices for each agg plan column once.
        let mut agg_chunks: Vec<Option<&[EncodedChunk]>> = Vec::new();
        let _ = agg_chunks.try_reserve_exact(self.agg_plans.len());
        for plan in &self.agg_plans {
            if plan.from_key_pos.is_some() {
                agg_chunks.push(None);
                continue;
            }
            let chunks = table.encoded_chunks(plan.col).ok_or(QueryError::ColumnOutOfBounds {
                col: plan.col,
                column_count: table.column_count(),
            })?;
            agg_chunks.push(Some(chunks));
        }

        let mut current_chunk_idx = usize::MAX;
        let mut key_chunk_refs: Vec<&EncodedChunk> = Vec::new();
        let _ = key_chunk_refs.try_reserve_exact(self.key_cols.len());
        let mut agg_chunk_refs: Vec<Option<&EncodedChunk>> = Vec::new();
        let _ = agg_chunk_refs.try_reserve_exact(self.agg_plans.len());

        for row in mask.iter_ones() {
            // Mask iteration should keep us in bounds, but keep defensive behavior.
            if row >= row_count {
                return Err(QueryError::RowOutOfBounds { row, row_count });
            }

            let chunk_idx = row / page;
            let row_in_chunk = row % page;

            if chunk_idx != current_chunk_idx {
                current_chunk_idx = chunk_idx;

                key_chunk_refs.clear();
                for chunks in &key_chunks {
                    let chunk = chunks
                        .get(chunk_idx)
                        .ok_or(QueryError::RowOutOfBounds { row, row_count })?;
                    key_chunk_refs.push(chunk);
                }

                agg_chunk_refs.clear();
                for (plan_idx, plan) in self.agg_plans.iter().enumerate() {
                    if plan.from_key_pos.is_some() {
                        agg_chunk_refs.push(None);
                        continue;
                    }
                    let chunks = agg_chunks
                        .get(plan_idx)
                        .and_then(|chunks| chunks.as_ref().copied())
                        .ok_or(QueryError::InternalInvariant(
                            "agg chunks missing for non-key agg plan",
                        ))?;
                    let chunk = chunks
                        .get(chunk_idx)
                        .ok_or(QueryError::RowOutOfBounds { row, row_count })?;
                    agg_chunk_refs.push(Some(chunk));
                }
            }

            for pos in 0..self.key_cols.len() {
                let col_idx = self.key_cols[pos];
                let ty = table.schema()[col_idx].column_type;
                let scalar = scalar_from_column_chunk_at(
                    col_idx,
                    ty,
                    key_chunk_refs[pos],
                    row_in_chunk,
                )?;
                self.scratch_key_scalars[pos] = scalar;
                self.scratch_keys[pos] = scalar_to_key(self.key_kinds[pos], scalar);
            }

            let group_idx = if let Some(&idx) = self.groups.get(self.scratch_keys.as_slice()) {
                idx
            } else {
                let idx = self.groups_len;
                self.groups_len += 1;
                self.groups
                    .insert(self.scratch_keys.to_vec().into_boxed_slice(), idx);
                for (pos, builder) in self.key_builders.iter_mut().enumerate() {
                    builder.push(self.scratch_key_scalars[pos]);
                }
                for state in &mut self.agg_states {
                    state.push_group();
                }
                idx
            };

            for &agg_idx in &self.count_row_aggs {
                self.agg_states[agg_idx].update_count_row(group_idx);
            }

            for (plan_idx, plan) in self.agg_plans.iter().enumerate() {
                let scalar = if let Some(key_pos) = plan.from_key_pos {
                    self.scratch_key_scalars[key_pos]
                } else {
                    let chunk = agg_chunk_refs
                        .get(plan_idx)
                        .and_then(|chunk| chunk.as_ref().copied())
                        .ok_or(QueryError::InternalInvariant(
                            "chunk missing for non-key agg plan",
                        ))?;
                    let ty = table.schema()[plan.col].column_type;
                    scalar_from_column_chunk_at(plan.col, ty, chunk, row_in_chunk)?
                };
                for &agg_idx in &plan.agg_indices {
                    self.agg_states[agg_idx].update_from_scalar(group_idx, scalar);
                }
            }
        }

        Ok(())
    }

    pub fn consume_all(&mut self, table: &ColumnarTable) -> Result<(), QueryError> {
        let rows = table.row_count();
        let page = table.page_size_rows();
        let chunk_count = (rows + page - 1) / page;
        self.consume_chunks(table, 0, chunk_count)
    }

    pub fn finish(self) -> GroupByResult {
        let mut columns: Vec<ResultColumn> = Vec::new();
        let _ = columns.try_reserve_exact(self.key_builders.len() + self.agg_states.len());
        for b in self.key_builders {
            columns.push(b.finish());
        }
        for s in self.agg_states {
            columns.push(s.finish());
        }

        debug_assert!(
            columns.iter().all(|c| c.len() == self.groups_len),
            "group-by column lengths should match number of groups"
        );

        GroupByResult {
            schema: self.schema,
            columns,
            rows: self.groups_len,
        }
    }
}

pub fn group_by(table: &ColumnarTable, keys: &[usize], aggs: &[AggSpec]) -> Result<GroupByResult, QueryError> {
    let mut engine = GroupByEngine::new(table, keys, aggs)?;
    engine.consume_all(table)?;
    Ok(engine.finish())
}

pub fn group_by_rows(
    table: &ColumnarTable,
    keys: &[usize],
    aggs: &[AggSpec],
    rows: &[usize],
) -> Result<GroupByResult, QueryError> {
    let mut engine = GroupByEngine::new(table, keys, aggs)?;
    engine.consume_rows(table, rows)?;
    Ok(engine.finish())
}

pub fn group_by_mask(
    table: &ColumnarTable,
    keys: &[usize],
    aggs: &[AggSpec],
    mask: &BitVec,
) -> Result<GroupByResult, QueryError> {
    let mut engine = GroupByEngine::new(table, keys, aggs)?;
    engine.consume_mask(table, mask)?;
    Ok(engine.finish())
}

/// Output of hash joins.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct JoinResult<L = usize, R = usize> {
    pub left_indices: Vec<L>,
    pub right_indices: Vec<R>,
}

impl<L, R> JoinResult<L, R> {
    pub fn len(&self) -> usize {
        self.left_indices.len()
    }

    pub fn is_empty(&self) -> bool {
        self.left_indices.is_empty()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum JoinType {
    /// Only emit pairs of matching rows.
    Inner,
    /// Emit all left rows. Unmatched rows have `None` for the right index.
    Left,
    /// Emit all right rows. Unmatched rows have `None` for the left index.
    Right,
    /// Emit all rows from both sides. Unmatched rows have `None` for the missing partner index.
    FullOuter,
}

fn build_dict_mapping(
    left_dict: &Arc<Vec<Arc<str>>>,
    right_dict: &Arc<Vec<Arc<str>>>,
) -> Vec<Option<u32>> {
    if Arc::ptr_eq(left_dict, right_dict) {
        return (0..right_dict.len()).map(|i| Some(i as u32)).collect();
    }

    let mut map: FastHashMap<&str, u32> =
        FastHashMap::with_capacity_and_hasher(left_dict.len(), FastBuildHasher::default());
    for (idx, s) in left_dict.iter().enumerate() {
        map.insert(s.as_ref(), idx as u32);
    }

    right_dict
        .iter()
        .map(|s| map.get(s.as_ref()).copied())
        .collect()
}

pub fn hash_join(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_on: usize,
    right_on: usize,
) -> Result<JoinResult, QueryError> {
    let left_type = left
        .schema()
        .get(left_on)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds {
            col: left_on,
            column_count: left.column_count(),
        })?;
    let right_type = right
        .schema()
        .get(right_on)
        .map(|s| s.column_type)
        .ok_or(QueryError::ColumnOutOfBounds {
            col: right_on,
            column_count: right.column_count(),
        })?;

    if left_type != right_type {
        return Err(QueryError::MismatchedJoinKeyTypes {
            left_type,
            right_type,
        });
    }

    let key_kind = key_kind_for_column_type(left_type).ok_or(QueryError::UnsupportedColumnType {
        col: left_on,
        column_type: left_type,
        operation: "JOIN key",
    })?;

    // For dictionary-encoded string keys, map right dictionary indices into the left dictionary
    // index space. When both sides share the same dictionary (common when joining cloned tables),
    // the mapping can be treated as an identity and we avoid allocating the vector entirely.
    let dict_mapping = if left_type == ColumnType::String {
        let left_dict = left
            .dictionary(left_on)
            .ok_or(QueryError::MissingDictionary { col: left_on })?;
        let right_dict = right
            .dictionary(right_on)
            .ok_or(QueryError::MissingDictionary { col: right_on })?;
        if Arc::ptr_eq(&left_dict, &right_dict) {
            None
        } else {
            Some(build_dict_mapping(&left_dict, &right_dict))
        }
    } else {
        None
    };

    let right_rows = right.row_count();
    let mut next: Vec<usize> = vec![usize::MAX; right_rows];

    let capacity_hint = right
        .scan()
        .stats(right_on)
        .map(|s| s.distinct_count as usize)
        .unwrap_or(0)
        .min(right_rows);

    let mut map: FastHashMap<KeyValue, usize> =
        FastHashMap::with_capacity_and_hasher(capacity_hint, FastBuildHasher::default());

    // Build phase (right).
    let right_chunks = right
        .encoded_chunks(right_on)
        .ok_or(QueryError::ColumnOutOfBounds {
            col: right_on,
            column_count: right.column_count(),
        })?;
    let page = right.page_size_rows();
    for (chunk_idx, chunk) in right_chunks.iter().enumerate() {
        let base = chunk_idx * page;
        let chunk_len = chunk.len();
        let mut cursor = ScalarChunkCursor::from_column_chunk(right_on, right_type, chunk)?;
        for i in 0..chunk_len {
            let row = base + i;
            if row >= right_rows {
                break;
            }
            let scalar = cursor.next();
            let key = match (key_kind, scalar) {
                (_, Scalar::Null) => continue,
                (KeyKind::Dict, Scalar::U32(ix)) => match dict_mapping.as_ref() {
                    Some(mapping) => {
                        let Some(mapped) = mapping.get(ix as usize).and_then(|m| *m) else {
                            continue;
                        };
                        KeyValue::Dict(mapped)
                    }
                    None => KeyValue::Dict(ix),
                },
                (kind, s) => scalar_to_key(kind, s),
            };
            if matches!(key, KeyValue::Null) {
                continue;
            }
            match map.get(&key).copied() {
                Some(head) => {
                    next[row] = head;
                    map.insert(key, row);
                }
                None => {
                    map.insert(key, row);
                }
            }
        }
    }

    // Probe phase (left).
    let mut out = JoinResult {
        left_indices: Vec::new(),
        right_indices: Vec::new(),
    };
    out.left_indices.reserve(left.row_count().min(right.row_count()));
    out.right_indices.reserve(left.row_count().min(right.row_count()));

    let left_chunks = left
        .encoded_chunks(left_on)
        .ok_or(QueryError::ColumnOutOfBounds {
            col: left_on,
            column_count: left.column_count(),
        })?;
    let page = left.page_size_rows();
    for (chunk_idx, chunk) in left_chunks.iter().enumerate() {
        let base = chunk_idx * page;
        let chunk_len = chunk.len();
        let mut cursor = ScalarChunkCursor::from_column_chunk(left_on, left_type, chunk)?;
        for i in 0..chunk_len {
            let row = base + i;
            if row >= left.row_count() {
                break;
            }
            let scalar = cursor.next();
            if matches!(scalar, Scalar::Null) {
                continue;
            }
            let key = match (key_kind, scalar) {
                (KeyKind::Dict, Scalar::U32(ix)) => KeyValue::Dict(ix),
                (kind, s) => scalar_to_key(kind, s),
            };
            if matches!(key, KeyValue::Null) {
                continue;
            }
            let Some(&head) = map.get(&key) else {
                continue;
            };
            let mut r = head;
            while r != usize::MAX {
                out.left_indices.push(row);
                out.right_indices.push(r);
                r = next[r];
            }
        }
    }

    Ok(out)
}

/// Hash join on a single key column (left join).
///
/// Rows from the left table with no match (or NULL key) are included with `None` for the right
/// index.
pub fn hash_left_join(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_on: usize,
    right_on: usize,
) -> Result<JoinResult<usize, Option<usize>>, QueryError> {
    hash_left_join_multi(left, right, &[left_on], &[right_on])
}

/// Hash join on a single key column (right join).
///
/// Rows from the right table with no match (or NULL key) are included with `None` for the left
/// index.
pub fn hash_right_join(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_on: usize,
    right_on: usize,
) -> Result<JoinResult<Option<usize>, usize>, QueryError> {
    hash_right_join_multi(left, right, &[left_on], &[right_on])
}

/// Hash join on a single key column (full outer join).
///
/// Unmatched rows from either side are included with `None` for the missing partner index.
pub fn hash_full_outer_join(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_on: usize,
    right_on: usize,
) -> Result<JoinResult<Option<usize>, Option<usize>>, QueryError> {
    hash_full_outer_join_multi(left, right, &[left_on], &[right_on])
}

/// Hash join on a single key column with a runtime join type.
///
/// This is a convenience API that always returns optional indices, regardless of join type.
pub fn hash_join_with_type(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_on: usize,
    right_on: usize,
    join_type: JoinType,
) -> Result<JoinResult<Option<usize>, Option<usize>>, QueryError> {
    hash_join_multi_with_type(left, right, &[left_on], &[right_on], join_type)
}

struct JoinKeyPlan {
    left_col: usize,
    right_col: usize,
    column_type: ColumnType,
    kind: KeyKind,
    /// For string keys, maps right dictionary indices into the left dictionary index space.
    /// Missing entries mean the right value is not present in the left dictionary and can never match.
    ///
    /// `None` is used for two cases:
    /// - non-string join keys (`kind != KeyKind::Dict`)
    /// - string join keys where both sides share the same dictionary (identity mapping)
    right_dict_to_left: Option<Vec<Option<u32>>>,
}

fn plan_join_keys(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
) -> Result<Vec<JoinKeyPlan>, QueryError> {
    if left_keys.is_empty() {
        return Err(QueryError::EmptyKeys);
    }
    if left_keys.len() != right_keys.len() {
        return Err(QueryError::MismatchedJoinKeyCount {
            left: left_keys.len(),
            right: right_keys.len(),
        });
    }

    let mut plans = Vec::new();
    let _ = plans.try_reserve_exact(left_keys.len());
    for (&left_col, &right_col) in left_keys.iter().zip(right_keys.iter()) {
        let left_type = left
            .schema()
            .get(left_col)
            .map(|s| s.column_type)
            .ok_or(QueryError::ColumnOutOfBounds {
                col: left_col,
                column_count: left.column_count(),
            })?;
        let right_type = right
            .schema()
            .get(right_col)
            .map(|s| s.column_type)
            .ok_or(QueryError::ColumnOutOfBounds {
                col: right_col,
                column_count: right.column_count(),
            })?;

        if left_type != right_type {
            return Err(QueryError::MismatchedJoinKeyTypes {
                left_type,
                right_type,
            });
        }

        let kind = key_kind_for_column_type(left_type).ok_or(QueryError::UnsupportedColumnType {
            col: left_col,
            column_type: left_type,
            operation: "JOIN key",
        })?;

        let right_dict_to_left = if left_type == ColumnType::String {
            let left_dict = left
                .dictionary(left_col)
                .ok_or(QueryError::MissingDictionary { col: left_col })?;
            let right_dict = right
                .dictionary(right_col)
                .ok_or(QueryError::MissingDictionary { col: right_col })?;
            if Arc::ptr_eq(&left_dict, &right_dict) {
                None
            } else {
                Some(build_dict_mapping(&left_dict, &right_dict))
            }
        } else {
            None
        };

        plans.push(JoinKeyPlan {
            left_col,
            right_col,
            column_type: left_type,
            kind,
            right_dict_to_left,
        });
    }

    Ok(plans)
}

fn join_key_from_scalar_for_right(
    plan: &JoinKeyPlan,
    scalar: Scalar,
) -> Option<KeyValue> {
    match (plan.kind, scalar) {
        (_, Scalar::Null) => None,
        (KeyKind::Dict, Scalar::U32(ix)) => match plan.right_dict_to_left.as_ref() {
            Some(mapping) => mapping.get(ix as usize).and_then(|m| *m).map(KeyValue::Dict),
            None => Some(KeyValue::Dict(ix)),
        },
        (kind, s) => {
            let key = scalar_to_key(kind, s);
            if matches!(key, KeyValue::Null) {
                None
            } else {
                Some(key)
            }
        }
    }
}

fn join_key_from_scalar_for_left(plan: &JoinKeyPlan, scalar: Scalar) -> Option<KeyValue> {
    match (plan.kind, scalar) {
        (_, Scalar::Null) => None,
        (KeyKind::Dict, Scalar::U32(ix)) => Some(KeyValue::Dict(ix)),
        (kind, s) => {
            let key = scalar_to_key(kind, s);
            if matches!(key, KeyValue::Null) {
                None
            } else {
                Some(key)
            }
        }
    }
}

fn hash_join_multi_core<L, R, FMatch, FLeft, FRight>(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
    mut push_match: FMatch,
    mut push_left_unmatched: FLeft,
    mut push_right_unmatched: FRight,
    track_unmatched_right: bool,
) -> Result<JoinResult<L, R>, QueryError>
where
    FMatch: FnMut(&mut JoinResult<L, R>, usize, usize),
    FLeft: FnMut(&mut JoinResult<L, R>, usize),
    FRight: FnMut(&mut JoinResult<L, R>, usize),
{
    let plans = plan_join_keys(left, right, left_keys, right_keys)?;

    let right_rows = right.row_count();
    let mut next: Vec<usize> = vec![usize::MAX; right_rows];

    // Capacity hint: when stats exist for all key columns, approximate distinct composite keys
    // as the product of per-column distinct counts (capped by row count).
    let capacity_hint = {
        let mut est: u128 = 1;
        for plan in &plans {
            let Some(stats) = right.scan().stats(plan.right_col) else {
                est = 0;
                break;
            };
            est = est.saturating_mul(stats.distinct_count as u128);
            est = est.min(right_rows as u128);
        }
        (est as usize).min(right_rows)
    };

    let mut map: FastHashMap<Box<[KeyValue]>, usize> =
        FastHashMap::with_capacity_and_hasher(capacity_hint, FastBuildHasher::default());

    // Build phase (right).
    let right_chunks_by_plan: Vec<&[EncodedChunk]> = plans
        .iter()
        .map(|plan| {
            right.encoded_chunks(plan.right_col)
                .ok_or(QueryError::ColumnOutOfBounds {
                    col: plan.right_col,
                    column_count: right.column_count(),
                })
        })
        .collect::<Result<_, _>>()?;
    let page = right.page_size_rows();
    let chunk_count = (right_rows + page - 1) / page;
    let mut scratch_keys: Vec<KeyValue> = vec![KeyValue::Null; plans.len()];
    let mut cursors: Vec<ScalarChunkCursor<'_>> = Vec::new();
    let _ = cursors.try_reserve_exact(plans.len());
    for chunk_idx in 0..chunk_count {
        let base = chunk_idx * page;
        if base >= right_rows {
            break;
        }
        let chunk_rows = (right_rows - base).min(page);

        cursors.clear();
        for (pos, plan) in plans.iter().enumerate() {
            let chunk = right_chunks_by_plan[pos].get(chunk_idx).ok_or(QueryError::RowOutOfBounds {
                row: base,
                row_count: right_rows,
            })?;
            cursors.push(ScalarChunkCursor::from_column_chunk(
                plan.right_col,
                plan.column_type,
                chunk,
            )?);
        }

        for i in 0..chunk_rows {
            let row = base + i;

            let mut valid = true;
            for (pos, plan) in plans.iter().enumerate() {
                let scalar = cursors[pos].next();
                if !valid {
                    continue;
                }
                match join_key_from_scalar_for_right(plan, scalar) {
                    Some(key) => scratch_keys[pos] = key,
                    None => valid = false,
                }
            }
            if !valid {
                continue;
            }

            let key_slice = scratch_keys.as_slice();
            if let Some(head) = map.get_mut(key_slice) {
                next[row] = *head;
                *head = row;
            } else {
                map.insert(scratch_keys.to_vec().into_boxed_slice(), row);
            }
        }
    }

    // Probe phase (left).
    let left_rows = left.row_count();
    let mut out: JoinResult<L, R> = JoinResult {
        left_indices: Vec::new(),
        right_indices: Vec::new(),
    };

    // Reserve something reasonable; outer joins can exceed this, but `reserve` is just a hint.
    let reserve_hint = match track_unmatched_right {
        true => left_rows.saturating_add(right_rows).min(left_rows.saturating_mul(2).max(1024)),
        false => left_rows.min(right_rows).max(1024),
    };
    out.left_indices.reserve(reserve_hint);
    out.right_indices.reserve(reserve_hint);

    let mut matched_right: Option<Vec<bool>> = track_unmatched_right.then(|| vec![false; right_rows]);

    let page = left.page_size_rows();
    let chunk_count = (left_rows + page - 1) / page;
    let left_chunks_by_plan: Vec<&[EncodedChunk]> = plans
        .iter()
        .map(|plan| {
            left.encoded_chunks(plan.left_col)
                .ok_or(QueryError::ColumnOutOfBounds {
                    col: plan.left_col,
                    column_count: left.column_count(),
                })
        })
        .collect::<Result<_, _>>()?;
    let mut cursors: Vec<ScalarChunkCursor<'_>> = Vec::new();
    let _ = cursors.try_reserve_exact(plans.len());
    for chunk_idx in 0..chunk_count {
        let base = chunk_idx * page;
        if base >= left_rows {
            break;
        }
        let chunk_rows = (left_rows - base).min(page);

        cursors.clear();
        for (pos, plan) in plans.iter().enumerate() {
            let chunk = left_chunks_by_plan[pos].get(chunk_idx).ok_or(QueryError::RowOutOfBounds {
                row: base,
                row_count: left_rows,
            })?;
            cursors.push(ScalarChunkCursor::from_column_chunk(
                plan.left_col,
                plan.column_type,
                chunk,
            )?);
        }

        for i in 0..chunk_rows {
            let row = base + i;

            let mut valid = true;
            for (pos, plan) in plans.iter().enumerate() {
                let scalar = cursors[pos].next();
                if !valid {
                    continue;
                }
                match join_key_from_scalar_for_left(plan, scalar) {
                    Some(key) => scratch_keys[pos] = key,
                    None => valid = false,
                }
            }
            if !valid {
                push_left_unmatched(&mut out, row);
                continue;
            }

            let Some(&head) = map.get(scratch_keys.as_slice()) else {
                push_left_unmatched(&mut out, row);
                continue;
            };

            let mut r = head;
            while r != usize::MAX {
                push_match(&mut out, row, r);
                if let Some(ref mut matched) = matched_right {
                    matched[r] = true;
                }
                r = next[r];
            }
        }
    }

    // Emit unmatched right rows (full outer join).
    if let Some(matched) = matched_right {
        for r in 0..right_rows {
            if !matched[r] {
                push_right_unmatched(&mut out, r);
            }
        }
    }

    Ok(out)
}

/// Hash join on multiple key columns (inner join).
pub fn hash_join_multi(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
) -> Result<JoinResult, QueryError> {
    hash_join_multi_core(
        left,
        right,
        left_keys,
        right_keys,
        |out: &mut JoinResult, l, r| {
            out.left_indices.push(l);
            out.right_indices.push(r);
        },
        |_out: &mut JoinResult, _l| {},
        |_out: &mut JoinResult, _r| {},
        false,
    )
}

/// Hash join on multiple key columns (left join).
///
/// Rows from the left table with no match (or NULL in any join key) are included with `None`
/// for the right index.
pub fn hash_left_join_multi(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
) -> Result<JoinResult<usize, Option<usize>>, QueryError> {
    hash_join_multi_core(
        left,
        right,
        left_keys,
        right_keys,
        |out: &mut JoinResult<usize, Option<usize>>, l, r| {
            out.left_indices.push(l);
            out.right_indices.push(Some(r));
        },
        |out: &mut JoinResult<usize, Option<usize>>, l| {
            out.left_indices.push(l);
            out.right_indices.push(None);
        },
        |_out: &mut JoinResult<usize, Option<usize>>, _r| {},
        false,
    )
}

/// Hash join on multiple key columns (right join).
///
/// Rows from the right table with no match (or NULL in any join key) are included with `None`
/// for the left index.
pub fn hash_right_join_multi(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
) -> Result<JoinResult<Option<usize>, usize>, QueryError> {
    hash_join_multi_core(
        left,
        right,
        left_keys,
        right_keys,
        |out: &mut JoinResult<Option<usize>, usize>, l, r| {
            out.left_indices.push(Some(l));
            out.right_indices.push(r);
        },
        |_out: &mut JoinResult<Option<usize>, usize>, _l| {},
        |out: &mut JoinResult<Option<usize>, usize>, r| {
            out.left_indices.push(None);
            out.right_indices.push(r);
        },
        true,
    )
}

/// Hash join on multiple key columns (full outer join).
///
/// Unmatched rows from either side are included with `None` for the missing partner index.
pub fn hash_full_outer_join_multi(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
) -> Result<JoinResult<Option<usize>, Option<usize>>, QueryError> {
    hash_join_multi_core(
        left,
        right,
        left_keys,
        right_keys,
        |out: &mut JoinResult<Option<usize>, Option<usize>>, l, r| {
            out.left_indices.push(Some(l));
            out.right_indices.push(Some(r));
        },
        |out: &mut JoinResult<Option<usize>, Option<usize>>, l| {
            out.left_indices.push(Some(l));
            out.right_indices.push(None);
        },
        |out: &mut JoinResult<Option<usize>, Option<usize>>, r| {
            out.left_indices.push(None);
            out.right_indices.push(Some(r));
        },
        true,
    )
}

/// Hash join on multiple key columns with a runtime join type.
///
/// This is a convenience API that always returns optional indices, regardless of join type.
pub fn hash_join_multi_with_type(
    left: &ColumnarTable,
    right: &ColumnarTable,
    left_keys: &[usize],
    right_keys: &[usize],
    join_type: JoinType,
) -> Result<JoinResult<Option<usize>, Option<usize>>, QueryError> {
    match join_type {
        JoinType::Inner => hash_join_multi_core(
            left,
            right,
            left_keys,
            right_keys,
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l, r| {
                out.left_indices.push(Some(l));
                out.right_indices.push(Some(r));
            },
            |_out: &mut JoinResult<Option<usize>, Option<usize>>, _l| {},
            |_out: &mut JoinResult<Option<usize>, Option<usize>>, _r| {},
            false,
        ),
        JoinType::Left => hash_join_multi_core(
            left,
            right,
            left_keys,
            right_keys,
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l, r| {
                out.left_indices.push(Some(l));
                out.right_indices.push(Some(r));
            },
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l| {
                out.left_indices.push(Some(l));
                out.right_indices.push(None);
            },
            |_out: &mut JoinResult<Option<usize>, Option<usize>>, _r| {},
            false,
        ),
        JoinType::Right => hash_join_multi_core(
            left,
            right,
            left_keys,
            right_keys,
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l, r| {
                out.left_indices.push(Some(l));
                out.right_indices.push(Some(r));
            },
            |_out: &mut JoinResult<Option<usize>, Option<usize>>, _l| {},
            |out: &mut JoinResult<Option<usize>, Option<usize>>, r| {
                out.left_indices.push(None);
                out.right_indices.push(Some(r));
            },
            true,
        ),
        JoinType::FullOuter => hash_join_multi_core(
            left,
            right,
            left_keys,
            right_keys,
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l, r| {
                out.left_indices.push(Some(l));
                out.right_indices.push(Some(r));
            },
            |out: &mut JoinResult<Option<usize>, Option<usize>>, l| {
                out.left_indices.push(Some(l));
                out.right_indices.push(None);
            },
            |out: &mut JoinResult<Option<usize>, Option<usize>>, r| {
                out.left_indices.push(None);
                out.right_indices.push(Some(r));
            },
            true,
        ),
    }
}
