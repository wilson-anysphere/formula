use crate::engine::{DaxError, DaxResult};
use crate::model::normalize_ident;
use crate::value::Value;
use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum AggregationKind {
    Sum,
    Average,
    Min,
    Max,
    CountRows,
    /// Count non-blank (non-null) values in a column.
    CountNonBlank,
    /// Count non-blank numeric values in a column (DAX `COUNT` semantics).
    CountNumbers,
    DistinctCount,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct AggregationSpec {
    pub kind: AggregationKind,
    pub column_idx: Option<usize>,
}

/// Storage abstraction for tables used by the DAX engine.
///
/// The engine relies on this trait to:
/// - resolve column indices by name
/// - access scalar values by row/column
/// - optionally use backend-specific accelerations (column stats, dictionary lookups, scans)
pub trait TableBackend: fmt::Debug + Send + Sync {
    fn columns(&self) -> &[String];
    fn row_count(&self) -> usize;
    fn column_index(&self, column: &str) -> Option<usize>;
    fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value>;

    fn value(&self, row: usize, column: &str) -> Option<Value> {
        let idx = self.column_index(column)?;
        self.value_by_idx(row, idx)
    }

    /// Precomputed sum for a numeric column (ignoring blanks/nulls), if available.
    fn stats_sum(&self, _idx: usize) -> Option<f64> {
        None
    }

    /// Number of non-blank values in a column, if available.
    fn stats_non_blank_count(&self, _idx: usize) -> Option<usize> {
        None
    }

    fn stats_min(&self, _idx: usize) -> Option<Value> {
        None
    }

    fn stats_max(&self, _idx: usize) -> Option<Value> {
        None
    }

    /// Distinct count for a column excluding blanks/nulls, if available.
    fn stats_distinct_count(&self, _idx: usize) -> Option<u64> {
        None
    }

    /// Whether the column has any blanks/nulls, if available.
    fn stats_has_blank(&self, _idx: usize) -> Option<bool> {
        None
    }

    /// If the backend has a dictionary of distinct values for the column, return them.
    ///
    /// The returned values exclude blanks/nulls (dictionary-encoded columns do not store null
    /// entries in the dictionary).
    fn dictionary_values(&self, _idx: usize) -> Option<Vec<Value>> {
        None
    }

    /// If the backend can efficiently find rows where `column == value`, return those rows.
    fn filter_eq(&self, _idx: usize, _value: &Value) -> Option<Vec<usize>> {
        None
    }

    /// Enumerate distinct values for a column within a set of rows.
    ///
    /// When `rows` is `None`, the backend may assume all rows are included.
    fn distinct_values_filtered(&self, _idx: usize, _rows: Option<&[usize]>) -> Option<Vec<Value>> {
        None
    }

    /// Enumerate distinct values for a column within a set of rows selected by a bit mask.
    ///
    /// When `mask` is `None`, the backend may assume all rows are included.
    ///
    /// This is primarily an optimization for columnar-backed tables: it avoids materializing a
    /// potentially huge `Vec<usize>` of row indices for large filtered datasets.
    fn distinct_values_filtered_mask(
        &self,
        _idx: usize,
        _mask: Option<&formula_columnar::BitVec>,
    ) -> Option<Vec<Value>> {
        None
    }

    /// Group the table by `group_by` column indices and compute aggregations for each group.
    ///
    /// The returned rows must contain `group_by.len()` key values followed by one value per
    /// aggregation spec, in the same order as `aggs`.
    ///
    /// When `rows` is `None`, the backend may assume all rows are included.
    fn group_by_aggregations(
        &self,
        _group_by: &[usize],
        _aggs: &[AggregationSpec],
        _rows: Option<&[usize]>,
    ) -> Option<Vec<Vec<Value>>> {
        None
    }

    /// Group the table by `group_by` column indices and compute aggregations for each group,
    /// restricted to rows where `mask` is true.
    ///
    /// This is primarily an optimization for columnar-backed tables: it avoids materializing a
    /// potentially huge `Vec<usize>` of row indices for large filtered datasets.
    ///
    /// When `mask` is `None`, the backend may assume all rows are included.
    fn group_by_aggregations_mask(
        &self,
        _group_by: &[usize],
        _aggs: &[AggregationSpec],
        _mask: Option<&formula_columnar::BitVec>,
    ) -> Option<Vec<Vec<Value>>> {
        None
    }

    /// If the backend can efficiently find rows where `column IN (values...)`, return those rows.
    ///
    /// This is useful for relationship propagation and multi-value filters.
    fn filter_in(&self, _idx: usize, _values: &[Value]) -> Option<Vec<usize>> {
        None
    }

    /// If this backend is backed by a `formula_columnar::ColumnarTable`, return a reference to it.
    ///
    /// This enables optional join/group-by accelerations that require access to the columnar
    /// encoded representation.
    fn columnar_table(&self) -> Option<&formula_columnar::ColumnarTable> {
        None
    }

    /// If both tables are columnar, perform a hash join on a single key column.
    ///
    /// The result is a list of matching row index pairs (left indices refer to `self`, right to
    /// `right`). This is primarily intended as a building block for relationship lookups and
    /// filter propagation.
    fn hash_join(
        &self,
        right: &dyn TableBackend,
        left_on: usize,
        right_on: usize,
    ) -> Option<formula_columnar::JoinResult> {
        let left = self.columnar_table()?;
        let right = right.columnar_table()?;
        left.hash_join(right, left_on, right_on).ok()
    }
}

#[derive(Clone, Debug)]
pub struct InMemoryTableBackend {
    pub(crate) columns: Vec<String>,
    pub(crate) column_index: HashMap<String, usize>,
    pub(crate) rows: Vec<Vec<Value>>,
}

impl InMemoryTableBackend {
    pub fn new(columns: Vec<String>) -> Self {
        let column_index = columns
            .iter()
            .enumerate()
            .map(|(idx, c)| (normalize_ident(c), idx))
            .collect();

        Self {
            columns,
            column_index,
            rows: Vec::new(),
        }
    }

    pub fn push_row(&mut self, table: &str, row: Vec<Value>) -> DaxResult<()> {
        if row.len() != self.columns.len() {
            return Err(DaxError::SchemaMismatch {
                table: table.to_string(),
                expected: self.columns.len(),
                actual: row.len(),
            });
        }
        self.rows.push(row);
        Ok(())
    }

    pub fn add_column(&mut self, table: &str, name: String, values: Vec<Value>) -> DaxResult<()> {
        let key = normalize_ident(&name);
        if self.column_index.contains_key(&key) {
            return Err(DaxError::DuplicateColumn {
                table: table.to_string(),
                column: name,
            });
        }
        if values.len() != self.rows.len() {
            return Err(DaxError::ColumnLengthMismatch {
                table: table.to_string(),
                column: name,
                expected: self.rows.len(),
                actual: values.len(),
            });
        }

        let idx = self.columns.len();
        self.columns.push(name.clone());
        self.column_index.insert(key, idx);
        for (row, value) in self.rows.iter_mut().zip(values) {
            row.push(value);
        }
        Ok(())
    }

    pub fn set_value_by_idx(&mut self, row: usize, idx: usize, value: Value) -> DaxResult<()> {
        let row_ref = self
            .rows
            .get_mut(row)
            .ok_or_else(|| DaxError::Eval("row out of bounds".into()))?;
        let slot = row_ref
            .get_mut(idx)
            .ok_or_else(|| DaxError::Eval("column out of bounds".into()))?;
        *slot = value;
        Ok(())
    }
}

impl TableBackend for InMemoryTableBackend {
    fn columns(&self) -> &[String] {
        &self.columns
    }

    fn row_count(&self) -> usize {
        self.rows.len()
    }

    fn column_index(&self, column: &str) -> Option<usize> {
        self.column_index.get(&normalize_ident(column)).copied()
    }

    fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value> {
        self.rows.get(row)?.get(idx).cloned()
    }
}

#[derive(Clone)]
pub struct ColumnarTableBackend {
    pub(crate) columns: Vec<String>,
    pub(crate) column_index: HashMap<String, usize>,
    pub(crate) table: Arc<formula_columnar::ColumnarTable>,
}

#[derive(Clone, Copy)]
enum RowSelection<'a> {
    All,
    Rows(&'a [usize]),
    Mask(&'a formula_columnar::BitVec),
}

impl fmt::Debug for ColumnarTableBackend {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("ColumnarTableBackend")
            .field("columns", &self.columns)
            .field("rows", &self.table.row_count())
            .finish()
    }
}

impl ColumnarTableBackend {
    pub fn new(table: formula_columnar::ColumnarTable) -> Self {
        let columns: Vec<String> = table.schema().iter().map(|c| c.name.clone()).collect();
        let mut column_index = HashMap::new();
        for (idx, c) in columns.iter().enumerate() {
            column_index.entry(normalize_ident(c)).or_insert(idx);
        }

        Self {
            columns,
            column_index,
            table: Arc::new(table),
        }
    }

    pub fn from_arc(table: Arc<formula_columnar::ColumnarTable>) -> Self {
        let columns: Vec<String> = table.schema().iter().map(|c| c.name.clone()).collect();
        let mut column_index = HashMap::new();
        for (idx, c) in columns.iter().enumerate() {
            column_index.entry(normalize_ident(c)).or_insert(idx);
        }

        Self {
            columns,
            column_index,
            table,
        }
    }

    fn scale_factor(scale: u8) -> f64 {
        10_f64.powi(scale as i32)
    }

    fn dax_from_columnar_typed(
        value: formula_columnar::Value,
        ty: formula_columnar::ColumnType,
    ) -> Value {
        use formula_columnar::{ColumnType, Value as ColValue};

        match (ty, value) {
            (_, ColValue::Null) => Value::Blank,
            (_, ColValue::Number(n)) => Value::from(n),
            (_, ColValue::Boolean(b)) => Value::from(b),
            (_, ColValue::String(s)) => Value::from(s),

            // MVP: treat DateTime as an integer-backed numeric type (matching previous behavior).
            (ColumnType::DateTime, ColValue::DateTime(raw)) => Value::from(raw as f64),

            (ColumnType::Currency { scale }, ColValue::Currency(raw)) => {
                Value::from(raw as f64 / Self::scale_factor(scale))
            }
            (ColumnType::Percentage { scale }, ColValue::Percentage(raw)) => {
                Value::from(raw as f64 / Self::scale_factor(scale))
            }

            // Defensive fallback: if a value's variant doesn't match the schema metadata,
            // preserve the old "raw as f64" behavior rather than returning BLANK.
            (_, ColValue::DateTime(raw) | ColValue::Currency(raw) | ColValue::Percentage(raw)) => {
                Value::from(raw as f64)
            }
        }
    }

    fn numeric_from_columnar(value: &formula_columnar::Value) -> Option<f64> {
        match value {
            formula_columnar::Value::Number(n) => Some(*n),
            formula_columnar::Value::DateTime(v)
            | formula_columnar::Value::Currency(v)
            | formula_columnar::Value::Percentage(v) => Some(*v as f64),
            _ => None,
        }
    }

    fn numeric_from_columnar_typed(
        value: &formula_columnar::Value,
        ty: formula_columnar::ColumnType,
    ) -> Option<f64> {
        use formula_columnar::{ColumnType, Value as ColValue};

        match (ty, value) {
            (_, ColValue::Null) => None,
            (_, ColValue::Number(n)) => Some(*n),
            (ColumnType::DateTime, ColValue::DateTime(raw)) => Some(*raw as f64),
            (ColumnType::Currency { scale }, ColValue::Currency(raw)) => {
                Some(*raw as f64 / Self::scale_factor(scale))
            }
            (ColumnType::Percentage { scale }, ColValue::Percentage(raw)) => {
                Some(*raw as f64 / Self::scale_factor(scale))
            }
            // Defensive fallback (see `dax_from_columnar_typed`).
            (_, ColValue::DateTime(raw) | ColValue::Currency(raw) | ColValue::Percentage(raw)) => {
                Some(*raw as f64)
            }
            _ => None,
        }
    }

    fn is_dax_numeric_column(column_type: formula_columnar::ColumnType) -> bool {
        matches!(
            column_type,
            formula_columnar::ColumnType::Number
                | formula_columnar::ColumnType::DateTime
                | formula_columnar::ColumnType::Currency { .. }
                | formula_columnar::ColumnType::Percentage { .. }
        )
    }

    fn group_by_aggregations_query(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        selection: RowSelection<'_>,
    ) -> Option<Vec<Vec<Value>>> {
        use formula_columnar::{AggOp, AggSpec};
        use std::collections::HashMap;

        let key_len = group_by.len();
        if key_len == 0 {
            return None;
        }

        #[derive(Clone, Debug)]
        enum Plan {
            Direct {
                col: usize,
            },
            /// A numeric aggregation (SUM/AVERAGE) over an int-backed logical type that needs
            /// post-scaling (Currency/Percentage).
            DirectScaled { col: usize, scale: u8 },
            DistinctCount {
                distinct_non_blank_col: usize,
                count_rows_col: usize,
                count_non_blank_col: usize,
            },
            Constant(Value),
        }

        let schema = self.table.schema();

        let mut planned: Vec<AggSpec> = Vec::new();
        let mut planned_pos: HashMap<(AggOp, Option<usize>), usize> = HashMap::new();

        let mut ensure = |op: AggOp, column: Option<usize>| -> usize {
            if let Some(&pos) = planned_pos.get(&(op, column)) {
                return key_len + pos;
            }
            let pos = planned.len();
            planned.push(AggSpec {
                op,
                column,
                name: None,
            });
            planned_pos.insert((op, column), pos);
            key_len + pos
        };

        let mut plans: Vec<Plan> = Vec::new();
        let _ = plans.try_reserve_exact(aggs.len());
        for spec in aggs {
            let plan = match spec.kind {
                AggregationKind::CountRows => Plan::Direct {
                    col: ensure(AggOp::Count, None),
                },
                AggregationKind::CountNonBlank => {
                    let col_idx = spec.column_idx?;
                    // Fast path: if the column has no blanks at all, COUNTNONBLANK is equivalent
                    // to COUNTROWS(group) for every group.
                    if !self.stats_has_blank(col_idx).unwrap_or(true) {
                        Plan::Direct {
                            col: ensure(AggOp::Count, None),
                        }
                    } else {
                        Plan::Direct {
                            col: ensure(AggOp::Count, Some(col_idx)),
                        }
                    }
                }
                AggregationKind::CountNumbers => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        // Fast path: for numeric columns with no blanks, COUNT(column) is just
                        // COUNTROWS(group).
                        if !self.stats_has_blank(col_idx).unwrap_or(true) {
                            Plan::Direct {
                                col: ensure(AggOp::Count, None),
                            }
                        } else {
                            Plan::Direct {
                                col: ensure(AggOp::CountNumbers, Some(col_idx)),
                            }
                        }
                    } else {
                        Plan::Constant(Value::from(0))
                    }
                }
                AggregationKind::Sum => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        match column_type {
                            formula_columnar::ColumnType::Currency { scale }
                            | formula_columnar::ColumnType::Percentage { scale } => {
                                Plan::DirectScaled {
                                    col: ensure(AggOp::SumF64, Some(col_idx)),
                                    scale,
                                }
                            }
                            _ => Plan::Direct {
                                col: ensure(AggOp::SumF64, Some(col_idx)),
                            },
                        }
                    } else {
                        Plan::Constant(Value::Blank)
                    }
                }
                AggregationKind::Min => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        Plan::Direct {
                            col: ensure(AggOp::Min, Some(col_idx)),
                        }
                    } else {
                        Plan::Constant(Value::Blank)
                    }
                }
                AggregationKind::Max => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        Plan::Direct {
                            col: ensure(AggOp::Max, Some(col_idx)),
                        }
                    } else {
                        Plan::Constant(Value::Blank)
                    }
                }
                AggregationKind::Average => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        match column_type {
                            formula_columnar::ColumnType::Currency { scale }
                            | formula_columnar::ColumnType::Percentage { scale } => {
                                Plan::DirectScaled {
                                    col: ensure(AggOp::AvgF64, Some(col_idx)),
                                    scale,
                                }
                            }
                            _ => Plan::Direct {
                                col: ensure(AggOp::AvgF64, Some(col_idx)),
                            },
                        }
                    } else {
                        Plan::Constant(Value::Blank)
                    }
                }
                AggregationKind::DistinctCount => {
                    let col_idx = spec.column_idx?;
                    // Fast path: if the column has no blanks at all, DAX DISTINCTCOUNT is
                    // equivalent to counting distinct non-null values, so we can avoid the extra
                    // COUNTROWS/COUNTNONBLANK plumbing.
                    if !self.stats_has_blank(col_idx).unwrap_or(true) {
                        Plan::Direct {
                            col: ensure(AggOp::DistinctCount, Some(col_idx)),
                        }
                    } else {
                        Plan::DistinctCount {
                            distinct_non_blank_col: ensure(AggOp::DistinctCount, Some(col_idx)),
                            // DAX DISTINCTCOUNT includes BLANK if any rows are blank.
                            // Our columnar `DistinctCount` ignores nulls/blanks, so we add 1 when
                            // `COUNTROWS(group) > COUNTNONBLANK(column)` for the group.
                            count_rows_col: ensure(AggOp::Count, None),
                            count_non_blank_col: ensure(AggOp::Count, Some(col_idx)),
                        }
                    }
                }
            };
            plans.push(plan);
        }

        let row_count = self.table.row_count();
        if let RowSelection::Mask(mask) = selection {
            if mask.len() != row_count {
                return None;
            }
        }

        let selection = match selection {
            RowSelection::Rows(rows)
                if rows.len() == row_count
                    && rows.first().copied() == Some(0)
                    && rows.last().copied() == row_count.checked_sub(1) =>
            {
                RowSelection::All
            }
            RowSelection::Mask(mask) if mask.len() == row_count && mask.all_true() => {
                RowSelection::All
            }
            other => other,
        };

        let grouped = match selection {
            RowSelection::All => self.table.group_by(group_by, &planned).ok()?,
            RowSelection::Rows(rows) => self.table.group_by_rows(group_by, &planned, rows).ok()?,
            RowSelection::Mask(mask) => self.table.group_by_mask(group_by, &planned, mask).ok()?,
        };
        let grouped_schema = grouped.schema();
        let grouped_cols = grouped.to_values();
        let group_count = grouped.row_count();

        let mut out: Vec<Vec<Value>> = Vec::new();
        let _ = out.try_reserve_exact(group_count);
        for row_idx in 0..group_count {
            let mut row: Vec<Value> = Vec::new();
            let _ = row.try_reserve_exact(key_len + plans.len());
            for (col, &col_idx) in group_by.iter().enumerate().take(key_len) {
                let ty = schema.get(col_idx)?.column_type;
                let value = grouped_cols
                    .get(col)
                    .and_then(|c| c.get(row_idx))
                    .cloned()
                    .unwrap_or(formula_columnar::Value::Null);
                row.push(Self::dax_from_columnar_typed(value, ty));
            }
            for plan in &plans {
                let value = match plan {
                    Plan::Direct { col } => {
                        let col = *col;
                        let ty = grouped_schema.get(col).map(|s| s.column_type)?;
                        let v = grouped_cols
                            .get(col)
                            .and_then(|c| c.get(row_idx))
                            .cloned()
                            .unwrap_or(formula_columnar::Value::Null);
                        Self::dax_from_columnar_typed(v, ty)
                    }
                    Plan::DirectScaled { col, scale } => {
                        let col = *col;
                        let v = grouped_cols
                            .get(col)
                            .and_then(|c| c.get(row_idx))
                            .cloned()
                            .unwrap_or(formula_columnar::Value::Null);
                        match Self::numeric_from_columnar(&v) {
                            Some(raw) => Value::from(raw / Self::scale_factor(*scale)),
                            None => Value::Blank,
                        }
                    }
                    Plan::DistinctCount {
                        distinct_non_blank_col,
                        count_rows_col,
                        count_non_blank_col,
                    } => {
                        let distinct_non_blank = grouped_cols
                            .get(*distinct_non_blank_col)
                            .and_then(|c| c.get(row_idx))
                            .and_then(Self::numeric_from_columnar)
                            .unwrap_or(0.0) as i64;
                        let count_rows = grouped_cols
                            .get(*count_rows_col)
                            .and_then(|c| c.get(row_idx))
                            .and_then(Self::numeric_from_columnar)
                            .unwrap_or(0.0) as i64;
                        let count_non_blank = grouped_cols
                            .get(*count_non_blank_col)
                            .and_then(|c| c.get(row_idx))
                            .and_then(Self::numeric_from_columnar)
                            .unwrap_or(0.0) as i64;
                        let mut out = distinct_non_blank;
                        if count_rows > count_non_blank {
                            out += 1;
                        }
                        Value::from(out)
                    }
                    Plan::Constant(v) => v.clone(),
                };
                row.push(value);
            }
            out.push(row);
        }

        Some(out)
    }

    fn group_by_aggregations_scan(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        selection: RowSelection<'_>,
    ) -> Option<Vec<Vec<Value>>> {
        use std::collections::HashMap;
        use std::collections::HashSet;

        let row_count = self.table.row_count();
        if let RowSelection::Mask(mask) = selection {
            if mask.len() != row_count {
                return None;
            }
        }
        let schema = self.table.schema();

        #[derive(Clone)]
        enum AggState {
            Sum { sum: f64, count: usize },
            Avg { sum: f64, count: usize },
            Min { best: Option<f64> },
            Max { best: Option<f64> },
            CountRows { count: usize },
            CountNonBlank { count: usize },
            CountNumbers { count: usize },
            DistinctCount { set: HashSet<Value> },
        }

        impl AggState {
            fn new(spec: &AggregationSpec) -> Option<Self> {
                Some(match spec.kind {
                    AggregationKind::Sum => AggState::Sum { sum: 0.0, count: 0 },
                    AggregationKind::Average => AggState::Avg { sum: 0.0, count: 0 },
                    AggregationKind::Min => AggState::Min { best: None },
                    AggregationKind::Max => AggState::Max { best: None },
                    AggregationKind::CountRows => AggState::CountRows { count: 0 },
                    AggregationKind::CountNonBlank => AggState::CountNonBlank { count: 0 },
                    AggregationKind::CountNumbers => AggState::CountNumbers { count: 0 },
                    AggregationKind::DistinctCount => AggState::DistinctCount {
                        set: HashSet::new(),
                    },
                })
            }

            fn update(
                &mut self,
                spec: &AggregationSpec,
                column_type: Option<formula_columnar::ColumnType>,
                value: Option<&formula_columnar::Value>,
            ) {
                match (self, spec.kind) {
                    (AggState::CountRows { count }, AggregationKind::CountRows) => {
                        *count += 1;
                    }
                    (AggState::CountNonBlank { count }, AggregationKind::CountNonBlank) => {
                        if let Some(v) = value {
                            if !matches!(v, formula_columnar::Value::Null) {
                                *count += 1;
                            }
                        }
                    }
                    (AggState::CountNumbers { count }, AggregationKind::CountNumbers) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        if value
                            .and_then(|v| ColumnarTableBackend::numeric_from_columnar_typed(v, column_type))
                            .is_some()
                        {
                            *count += 1;
                        }
                    }
                    (AggState::Sum { sum, count }, AggregationKind::Sum) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        if let Some(v) =
                            value.and_then(|v| ColumnarTableBackend::numeric_from_columnar_typed(v, column_type))
                        {
                            *sum += v;
                            *count += 1;
                        }
                    }
                    (AggState::Avg { sum, count }, AggregationKind::Average) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        if let Some(v) =
                            value.and_then(|v| ColumnarTableBackend::numeric_from_columnar_typed(v, column_type))
                        {
                            *sum += v;
                            *count += 1;
                        }
                    }
                    (AggState::Min { best }, AggregationKind::Min) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        if let Some(v) =
                            value.and_then(|v| ColumnarTableBackend::numeric_from_columnar_typed(v, column_type))
                        {
                            *best = Some(best.map_or(v, |current| current.min(v)));
                        }
                    }
                    (AggState::Max { best }, AggregationKind::Max) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        if let Some(v) =
                            value.and_then(|v| ColumnarTableBackend::numeric_from_columnar_typed(v, column_type))
                        {
                            *best = Some(best.map_or(v, |current| current.max(v)));
                        }
                    }
                    (AggState::DistinctCount { set }, AggregationKind::DistinctCount) => {
                        let Some(column_type) = column_type else {
                            return;
                        };
                        let Some(v) = value else {
                            return;
                        };
                        set.insert(ColumnarTableBackend::dax_from_columnar_typed(
                            v.clone(),
                            column_type,
                        ));
                    }
                    _ => {}
                }
            }

            fn finalize(self) -> Value {
                match self {
                    AggState::Sum { sum, count } => {
                        if count == 0 {
                            Value::Blank
                        } else {
                            Value::from(sum)
                        }
                    }
                    AggState::Avg { sum, count } => {
                        if count == 0 {
                            Value::Blank
                        } else {
                            Value::from(sum / count as f64)
                        }
                    }
                    AggState::Min { best } => best.map(Value::from).unwrap_or(Value::Blank),
                    AggState::Max { best } => best.map(Value::from).unwrap_or(Value::Blank),
                    AggState::CountRows { count } => Value::from(count as i64),
                    AggState::CountNonBlank { count } => Value::from(count as i64),
                    AggState::CountNumbers { count } => Value::from(count as i64),
                    AggState::DistinctCount { set } => Value::from(set.len() as i64),
                }
            }
        }

        let key_len = group_by.len();
        let key_types: Vec<formula_columnar::ColumnType> = group_by
            .iter()
            .map(|&idx| schema.get(idx).map(|c| c.column_type))
            .collect::<Option<Vec<_>>>()?;
        let agg_types: Vec<Option<formula_columnar::ColumnType>> = aggs
            .iter()
            .map(|spec| spec.column_idx.and_then(|idx| schema.get(idx).map(|c| c.column_type)))
            .collect();
        let mut key_buf: Vec<Value> = Vec::new();
        let _ = key_buf.try_reserve_exact(key_len);

        let mut states_template = Vec::new();
        let _ = states_template.try_reserve_exact(aggs.len());
        for spec in aggs {
            states_template.push(AggState::new(spec)?);
        }

        let mut groups: HashMap<Vec<Value>, Vec<AggState>> = HashMap::new();

        let needed_min = group_by
            .iter()
            .chain(aggs.iter().filter_map(|a| a.column_idx.as_ref()))
            .copied()
            .min()
            .unwrap_or(0);
        let needed_max = group_by
            .iter()
            .chain(aggs.iter().filter_map(|a| a.column_idx.as_ref()))
            .copied()
            .max()
            .unwrap_or(0);
        let col_start = needed_min;
        let col_end = needed_max + 1;

        const CHUNK_ROWS: usize = 65_536;
        let mut process_row = |row_offset: usize, range: &formula_columnar::ColumnarRange| {
            key_buf.clear();
            for (pos, idx) in group_by.iter().enumerate() {
                let col = *idx - col_start;
                let value = range
                    .columns
                    .get(col)
                    .and_then(|c| c.get(row_offset))
                    .cloned()
                    .unwrap_or(formula_columnar::Value::Null);
                key_buf.push(Self::dax_from_columnar_typed(value, key_types[pos]));
            }

            if let Some(existing) = groups.get_mut(key_buf.as_slice()) {
                for ((state, spec), col_type) in existing.iter_mut().zip(aggs).zip(&agg_types) {
                    let value = spec.column_idx.and_then(|idx| {
                        let col = idx - col_start;
                        range.columns.get(col).and_then(|c| c.get(row_offset))
                    });
                    state.update(spec, *col_type, value);
                }
                return;
            }

            let mut states = states_template.clone();
            for ((state, spec), col_type) in states.iter_mut().zip(aggs).zip(&agg_types) {
                let value = spec.column_idx.and_then(|idx| {
                    let col = idx - col_start;
                    range.columns.get(col).and_then(|c| c.get(row_offset))
                });
                state.update(spec, *col_type, value);
            }
            groups.insert(key_buf.clone(), states);
        };

        match selection {
            RowSelection::All => {
                let mut start = 0;
                while start < row_count {
                    let end = (start + CHUNK_ROWS).min(row_count);
                    let range = self.table.get_range(start, end, col_start, col_end);
                    for row_offset in 0..range.rows() {
                        process_row(row_offset, &range);
                    }
                    start = end;
                }
            }
            RowSelection::Rows(rows) => {
                let mut pos = 0;
                while pos < rows.len() {
                    let row = rows[pos];
                    let chunk_start = (row / CHUNK_ROWS) * CHUNK_ROWS;
                    let chunk_end = (chunk_start + CHUNK_ROWS).min(row_count);
                    let range = self
                        .table
                        .get_range(chunk_start, chunk_end, col_start, col_end);
                    while pos < rows.len() {
                        let row = rows[pos];
                        if row >= chunk_end {
                            break;
                        }
                        process_row(row - chunk_start, &range);
                        pos += 1;
                    }
                }
            }
            RowSelection::Mask(mask) => {
                let mut current_chunk_start = usize::MAX;
                let mut current_range: Option<formula_columnar::ColumnarRange> = None;
                for row in mask.iter_ones() {
                    if row >= row_count {
                        return None;
                    }
                    let chunk_start = (row / CHUNK_ROWS) * CHUNK_ROWS;
                    if chunk_start != current_chunk_start {
                        current_chunk_start = chunk_start;
                        current_range = Some(
                            self.table
                                .get_range(chunk_start, (chunk_start + CHUNK_ROWS).min(row_count), col_start, col_end),
                        );
                    }

                    let Some(range) = current_range.as_ref() else {
                        debug_assert!(false, "current_range missing for chunk_start={chunk_start}");
                        return None;
                    };
                    process_row(row - current_chunk_start, range);
                }
            }
        }

        let mut out = Vec::new();
        let _ = out.try_reserve_exact(groups.len());
        for (key, states) in groups {
            let mut row = key;
            for state in states {
                row.push(state.finalize());
            }
            out.push(row);
        }
        Some(out)
    }

    fn distinct_values_filtered_selection(
        &self,
        idx: usize,
        selection: RowSelection<'_>,
    ) -> Option<Vec<Value>> {
        let row_count = self.table.row_count();
        if idx >= self.table.column_count() {
            return None;
        }
        let column_type = self.table.schema().get(idx)?.column_type;

        // If the mask is sparse, materializing row indices can be cheaper than scanning a full
        // bitmask (same heuristic as `UnmatchedFactRowsBuilder`).
        if let RowSelection::Mask(mask) = selection {
            if mask.len() != row_count {
                return None;
            }
            let visible = mask.count_ones();
            if visible == 0 {
                return Some(Vec::new());
            }
            let sparse_to_dense_threshold = row_count / 64;
            if visible <= sparse_to_dense_threshold {
                let rows: Vec<usize> = mask.iter_ones().collect();
                return self.distinct_values_filtered_selection(idx, RowSelection::Rows(rows.as_slice()));
            }
        }

        let selection = match selection {
            RowSelection::Rows(rows)
                if rows.len() == row_count
                    && rows.first().copied() == Some(0)
                    && rows.last().copied() == row_count.checked_sub(1) =>
            {
                RowSelection::All
            }
            RowSelection::Mask(mask) if mask.len() == row_count && mask.all_true() => RowSelection::All,
            other => other,
        };

        match selection {
            RowSelection::All => {
                if let Some(values) = self.dictionary_values(idx) {
                    let mut out: Vec<Value> = values;
                    if self.stats_has_blank(idx).unwrap_or(false) {
                        out.push(Value::Blank);
                    }
                    return Some(out);
                }

                if let Ok(result) = self.table.group_by(&[idx], &[]) {
                    let values = result.to_values();
                    let mut out = Vec::new();
                    let _ = out.try_reserve_exact(result.row_count());
                    if let Some(col) = values.get(0) {
                        for v in col {
                            out.push(Self::dax_from_columnar_typed(v.clone(), column_type));
                        }
                    }
                    return Some(out);
                }
            }
            RowSelection::Rows(rows) => {
                if rows.is_empty() {
                    return Some(Vec::new());
                }
                if let Ok(result) = self.table.group_by_rows(&[idx], &[], rows) {
                    let values = result.to_values();
                    let mut out = Vec::new();
                    let _ = out.try_reserve_exact(result.row_count());
                    if let Some(col) = values.get(0) {
                        for v in col {
                            out.push(Self::dax_from_columnar_typed(v.clone(), column_type));
                        }
                    }
                    return Some(out);
                }
            }
            RowSelection::Mask(mask) => {
                if mask.len() != row_count {
                    return None;
                }
                if mask.count_ones() == 0 {
                    return Some(Vec::new());
                }
                if let Ok(result) = self.table.group_by_mask(&[idx], &[], mask) {
                    let values = result.to_values();
                    let mut out = Vec::new();
                    let _ = out.try_reserve_exact(result.row_count());
                    if let Some(col) = values.get(0) {
                        for v in col {
                            out.push(Self::dax_from_columnar_typed(v.clone(), column_type));
                        }
                    }
                    return Some(out);
                }
            }
        }

        // Fallback: decode selected pages and de-duplicate in memory.
        use std::collections::HashSet;

        let mut seen: HashSet<Value> = HashSet::new();
        let mut out = Vec::new();

        const CHUNK_ROWS: usize = 65_536;
        let mut push_value = |value: Value| {
            if seen.insert(value.clone()) {
                out.push(value);
            }
        };

        match selection {
            RowSelection::All => {
                let mut start = 0;
                while start < row_count {
                    let end = (start + CHUNK_ROWS).min(row_count);
                    let range = self.table.get_range(start, end, idx, idx + 1);
                    for v in range.columns.get(0).into_iter().flatten() {
                        push_value(Self::dax_from_columnar_typed(v.clone(), column_type));
                    }
                    start = end;
                }
            }
            RowSelection::Rows(rows) => {
                let mut pos = 0;
                while pos < rows.len() {
                    let row = rows[pos];
                    let chunk_start = (row / CHUNK_ROWS) * CHUNK_ROWS;
                    let chunk_end = (chunk_start + CHUNK_ROWS).min(row_count);
                    let range = self.table.get_range(chunk_start, chunk_end, idx, idx + 1);
                    while pos < rows.len() {
                        let row = rows[pos];
                        if row >= chunk_end {
                            break;
                        }
                        let in_chunk = row - chunk_start;
                        if let Some(v) = range.columns.get(0).and_then(|c| c.get(in_chunk)) {
                            push_value(Self::dax_from_columnar_typed(v.clone(), column_type));
                        }
                        pos += 1;
                    }
                }
            }
            RowSelection::Mask(mask) => {
                if mask.len() != row_count {
                    return None;
                }
                let mut rows = mask.iter_ones().peekable();
                while let Some(&row) = rows.peek() {
                    let chunk_start = (row / CHUNK_ROWS) * CHUNK_ROWS;
                    let chunk_end = (chunk_start + CHUNK_ROWS).min(row_count);
                    let range = self.table.get_range(chunk_start, chunk_end, idx, idx + 1);
                    while let Some(&row) = rows.peek() {
                        if row >= chunk_end {
                            break;
                        }
                        let in_chunk = row - chunk_start;
                        if let Some(v) = range.columns.get(0).and_then(|c| c.get(in_chunk)) {
                            push_value(Self::dax_from_columnar_typed(v.clone(), column_type));
                        }
                        rows.next();
                    }
                }
            }
        }

        Some(out)
    }
}

impl TableBackend for ColumnarTableBackend {
    fn columns(&self) -> &[String] {
        &self.columns
    }

    fn row_count(&self) -> usize {
        self.table.row_count()
    }

    fn column_index(&self, column: &str) -> Option<usize> {
        self.column_index.get(&normalize_ident(column)).copied()
    }

    fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value> {
        if row >= self.table.row_count() || idx >= self.table.column_count() {
            return None;
        }
        let ty = self.table.schema().get(idx)?.column_type;
        let value = self.table.get_cell(row, idx);
        Some(Self::dax_from_columnar_typed(value, ty))
    }

    fn stats_sum(&self, idx: usize) -> Option<f64> {
        let stats = self.table.scan().stats(idx)?;
        let sum = stats.sum?;
        let ty = self.table.schema().get(idx)?.column_type;
        match ty {
            formula_columnar::ColumnType::Currency { scale }
            | formula_columnar::ColumnType::Percentage { scale } => {
                Some(sum / Self::scale_factor(scale))
            }
            _ => Some(sum),
        }
    }

    fn stats_non_blank_count(&self, idx: usize) -> Option<usize> {
        let stats = self.table.scan().stats(idx)?;
        let non_blank = self
            .table
            .row_count()
            .saturating_sub(stats.null_count as usize);
        Some(non_blank)
    }

    fn stats_min(&self, idx: usize) -> Option<Value> {
        let stats = self.table.scan().stats(idx)?;
        let ty = self.table.schema().get(idx)?.column_type;
        Some(Self::dax_from_columnar_typed(stats.min.clone()?, ty))
    }

    fn stats_max(&self, idx: usize) -> Option<Value> {
        let stats = self.table.scan().stats(idx)?;
        let ty = self.table.schema().get(idx)?.column_type;
        Some(Self::dax_from_columnar_typed(stats.max.clone()?, ty))
    }

    fn stats_distinct_count(&self, idx: usize) -> Option<u64> {
        Some(self.table.scan().stats(idx)?.distinct_count)
    }

    fn stats_has_blank(&self, idx: usize) -> Option<bool> {
        let stats = self.table.scan().stats(idx)?;
        Some(stats.null_count > 0)
    }

    fn dictionary_values(&self, idx: usize) -> Option<Vec<Value>> {
        let ty = self.table.schema().get(idx)?.column_type;
        let dict = self.table.dictionary(idx)?;
        Some(
            dict.iter()
                .cloned()
                .map(|s| Self::dax_from_columnar_typed(formula_columnar::Value::String(s), ty))
                .collect(),
        )
    }

    fn filter_eq(&self, idx: usize, value: &Value) -> Option<Vec<usize>> {
        if idx >= self.table.column_count() {
            return None;
        }

        let column_type = self.table.schema().get(idx)?.column_type;

        fn safe_i64_key(v: f64) -> Option<i64> {
            // DAX stores numeric values as `f64`. For int-backed columnar logical types we only
            // use the i64 scan acceleration when the value is a "safe integer", so that the fast
            // path matches the fallback semantics (`i64 as f64` conversion) for typical keys.
            const MAX_SAFE_INT: f64 = 9_007_199_254_740_992.0; // 2^53
            if !v.is_finite() {
                return None;
            }
            if v.fract() != 0.0 {
                return None;
            }
            if v.abs() > MAX_SAFE_INT {
                return None;
            }
            Some(v as i64)
        }

        fn scaled_i64_key(v: f64, scale: u8) -> Option<i64> {
            // Match the fallback semantics for fixed-point logical types (Currency/Percentage):
            // convert `raw: i64` to `f64` via `raw as f64 / 10^scale`, and compare as `f64`.
            // For correctness we only use the i64 scan fast path when the computed raw value
            // round-trips exactly back to the original `f64`.
            const MAX_SAFE_INT: f64 = 9_007_199_254_740_992.0; // 2^53
            if !v.is_finite() {
                return None;
            }
            let factor = ColumnarTableBackend::scale_factor(scale);
            let raw_f = v * factor;
            if !raw_f.is_finite() {
                return None;
            }
            let raw_round = raw_f.round();
            if raw_round.abs() > MAX_SAFE_INT {
                return None;
            }
            if raw_round < i64::MIN as f64 || raw_round > i64::MAX as f64 {
                return None;
            }
            let raw = raw_round as i64;
            ((raw as f64) / factor == v).then_some(raw)
        }

        match value {
            Value::Text(s) => Some(self.table.scan().filter_eq_string(idx, s.as_ref())),
            Value::Number(n) => match column_type {
                formula_columnar::ColumnType::Number => Some(self.table.scan().filter_eq_number(idx, n.0)),
                formula_columnar::ColumnType::DateTime => {
                    let Some(v) = safe_i64_key(n.0) else {
                        return self.filter_in(idx, std::slice::from_ref(value));
                    };
                    Some(self.table.scan().filter_eq_i64(idx, v))
                }
                formula_columnar::ColumnType::Currency { scale }
                | formula_columnar::ColumnType::Percentage { scale } => {
                    let Some(v) = scaled_i64_key(n.0, scale) else {
                        return self.filter_in(idx, std::slice::from_ref(value));
                    };
                    Some(self.table.scan().filter_eq_i64(idx, v))
                }
                _ => self.filter_in(idx, std::slice::from_ref(value)),
            },
            Value::Boolean(b) => match column_type {
                formula_columnar::ColumnType::Boolean => Some(self.table.scan().filter_eq_bool(idx, *b)),
                _ => self.filter_in(idx, std::slice::from_ref(value)),
            },
            _ => self.filter_in(idx, std::slice::from_ref(value)),
        }
    }

    fn distinct_values_filtered(&self, idx: usize, rows: Option<&[usize]>) -> Option<Vec<Value>> {
        let selection = rows.map_or(RowSelection::All, RowSelection::Rows);
        self.distinct_values_filtered_selection(idx, selection)
    }

    fn distinct_values_filtered_mask(
        &self,
        idx: usize,
        mask: Option<&formula_columnar::BitVec>,
    ) -> Option<Vec<Value>> {
        let selection = mask.map_or(RowSelection::All, RowSelection::Mask);
        self.distinct_values_filtered_selection(idx, selection)
    }

    fn group_by_aggregations(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        rows: Option<&[usize]>,
    ) -> Option<Vec<Vec<Value>>> {
        if group_by.is_empty() {
            return None;
        }
        if group_by
            .iter()
            .chain(aggs.iter().filter_map(|a| a.column_idx.as_ref()))
            .any(|idx| *idx >= self.table.column_count())
        {
            return None;
        }

        let selection = rows.map_or(RowSelection::All, RowSelection::Rows);

        if let Some(out) = self.group_by_aggregations_query(group_by, aggs, selection) {
            return Some(out);
        }

        self.group_by_aggregations_scan(group_by, aggs, selection)
    }

    fn group_by_aggregations_mask(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        mask: Option<&formula_columnar::BitVec>,
    ) -> Option<Vec<Vec<Value>>> {
        if group_by.is_empty() {
            return None;
        }
        if group_by
            .iter()
            .chain(aggs.iter().filter_map(|a| a.column_idx.as_ref()))
            .any(|idx| *idx >= self.table.column_count())
        {
            return None;
        }

        let selection = match mask {
            Some(mask) if mask.len() == self.table.row_count() && mask.all_true() => {
                RowSelection::All
            }
            Some(mask) => RowSelection::Mask(mask),
            None => RowSelection::All,
        };

        if let Some(out) = self.group_by_aggregations_query(group_by, aggs, selection) {
            return Some(out);
        }

        self.group_by_aggregations_scan(group_by, aggs, selection)
    }

    fn filter_in(&self, idx: usize, values: &[Value]) -> Option<Vec<usize>> {
        if values.is_empty() {
            return Some(Vec::new());
        }
        if idx >= self.table.column_count() {
            return None;
        }

        let column_type = self.table.schema().get(idx)?.column_type;

        if values.iter().all(|v| matches!(v, Value::Text(_))) {
            let strs: Vec<&str> = values
                .iter()
                .filter_map(|v| match v {
                    Value::Text(s) => Some(s.as_ref()),
                    _ => None,
                })
                .collect();
            return Some(self.table.scan().filter_in_string(idx, &strs));
        }

        if values.iter().all(|v| matches!(v, Value::Number(_))) {
            let nums: Vec<f64> = values
                .iter()
                .filter_map(|v| match v {
                    Value::Number(n) => Some(n.0),
                    _ => None,
                })
                .collect();

            match column_type {
                formula_columnar::ColumnType::Number => {
                    return Some(self.table.scan().filter_in_number(idx, &nums));
                }
                formula_columnar::ColumnType::DateTime => {
                    const MAX_SAFE_INT: f64 = 9_007_199_254_740_992.0; // 2^53
                    let mut ints = Vec::new();
                    let _ = ints.try_reserve_exact(nums.len());
                    for v in &nums {
                        if !v.is_finite() || v.fract() != 0.0 || v.abs() > MAX_SAFE_INT {
                            ints.clear();
                            break;
                        }
                        ints.push(*v as i64);
                    }

                    if !ints.is_empty() {
                        return Some(self.table.scan().filter_in_i64(idx, &ints));
                    }
                }
                formula_columnar::ColumnType::Currency { scale }
                | formula_columnar::ColumnType::Percentage { scale } => {
                    const MAX_SAFE_INT: f64 = 9_007_199_254_740_992.0; // 2^53
                    let factor = Self::scale_factor(scale);
                    let mut ints = Vec::new();
                    let _ = ints.try_reserve_exact(nums.len());
                    for v in &nums {
                        if !v.is_finite() {
                            ints.clear();
                            break;
                        }
                        let raw_f = *v * factor;
                        if !raw_f.is_finite() {
                            ints.clear();
                            break;
                        }
                        let raw_round = raw_f.round();
                        if raw_round.abs() > MAX_SAFE_INT
                            || raw_round < i64::MIN as f64
                            || raw_round > i64::MAX as f64
                        {
                            ints.clear();
                            break;
                        }
                        let raw = raw_round as i64;
                        if (raw as f64) / factor != *v {
                            ints.clear();
                            break;
                        }
                        ints.push(raw);
                    }

                    if !ints.is_empty() {
                        return Some(self.table.scan().filter_in_i64(idx, &ints));
                    }
                }
                _ => {}
            }
        }

        if values.iter().all(|v| matches!(v, Value::Boolean(_))) {
            let bools: Vec<bool> = values
                .iter()
                .filter_map(|v| match v {
                    Value::Boolean(b) => Some(*b),
                    _ => None,
                })
                .collect();
            if column_type == formula_columnar::ColumnType::Boolean {
                return Some(self.table.scan().filter_in_bool(idx, &bools));
            }
        }

        use std::collections::HashSet;
        let targets: HashSet<Value> = values.iter().cloned().collect();
        let row_count = self.table.row_count();

        const CHUNK_ROWS: usize = 65_536;
        let mut out = Vec::new();
        let mut start = 0;
        while start < row_count {
            let end = (start + CHUNK_ROWS).min(row_count);
            let range = self.table.get_range(start, end, idx, idx + 1);
            if let Some(col) = range.columns.get(0) {
                for (offset, v) in col.iter().enumerate() {
                    let dax_value = Self::dax_from_columnar_typed(v.clone(), column_type);
                    if targets.contains(&dax_value) {
                        out.push(start + offset);
                    }
                }
            }
            start = end;
        }
        Some(out)
    }

    fn columnar_table(&self) -> Option<&formula_columnar::ColumnarTable> {
        Some(self.table.as_ref())
    }
}
