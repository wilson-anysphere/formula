use crate::engine::{DaxError, DaxResult};
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
            .map(|(idx, c)| (c.clone(), idx))
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
        if self.column_index.contains_key(&name) {
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
        self.column_index.insert(name, idx);
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
        self.column_index.get(column).copied()
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
        let column_index = columns
            .iter()
            .enumerate()
            .map(|(idx, c)| (c.clone(), idx))
            .collect();

        Self {
            columns,
            column_index,
            table: Arc::new(table),
        }
    }

    pub fn from_arc(table: Arc<formula_columnar::ColumnarTable>) -> Self {
        let columns: Vec<String> = table.schema().iter().map(|c| c.name.clone()).collect();
        let column_index = columns
            .iter()
            .enumerate()
            .map(|(idx, c)| (c.clone(), idx))
            .collect();

        Self {
            columns,
            column_index,
            table,
        }
    }

    fn dax_from_columnar(value: formula_columnar::Value) -> Value {
        match value {
            formula_columnar::Value::Null => Value::Blank,
            formula_columnar::Value::Number(n) => Value::from(n),
            formula_columnar::Value::Boolean(b) => Value::from(b),
            formula_columnar::Value::String(s) => Value::from(s),
            formula_columnar::Value::DateTime(v) => Value::from(v as f64),
            formula_columnar::Value::Currency(v) => Value::from(v as f64),
            formula_columnar::Value::Percentage(v) => Value::from(v as f64),
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
        rows: Option<&[usize]>,
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

        let mut plans: Vec<Plan> = Vec::with_capacity(aggs.len());
        for spec in aggs {
            let plan = match spec.kind {
                AggregationKind::CountRows => Plan::Direct {
                    col: ensure(AggOp::Count, None),
                },
                AggregationKind::CountNonBlank => {
                    let col_idx = spec.column_idx?;
                    Plan::Direct {
                        col: ensure(AggOp::Count, Some(col_idx)),
                    }
                }
                AggregationKind::CountNumbers => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        Plan::Direct {
                            col: ensure(AggOp::CountNumbers, Some(col_idx)),
                        }
                    } else {
                        Plan::Constant(Value::from(0))
                    }
                }
                AggregationKind::Sum => {
                    let col_idx = spec.column_idx?;
                    let column_type = schema.get(col_idx)?.column_type;
                    if Self::is_dax_numeric_column(column_type) {
                        Plan::Direct {
                            col: ensure(AggOp::SumF64, Some(col_idx)),
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
                        Plan::Direct {
                            col: ensure(AggOp::AvgF64, Some(col_idx)),
                        }
                    } else {
                        Plan::Constant(Value::Blank)
                    }
                }
                AggregationKind::DistinctCount => {
                    let col_idx = spec.column_idx?;
                    Plan::DistinctCount {
                        distinct_non_blank_col: ensure(AggOp::DistinctCount, Some(col_idx)),
                        // DAX DISTINCTCOUNT includes BLANK if any rows are blank.
                        // Our columnar `DistinctCount` ignores nulls/blanks, so we add 1 when
                        // `COUNTROWS(group) > COUNTNONBLANK(column)` for the group.
                        count_rows_col: ensure(AggOp::Count, None),
                        count_non_blank_col: ensure(AggOp::Count, Some(col_idx)),
                    }
                }
            };
            plans.push(plan);
        }

        let row_count = self.table.row_count();
        let rows = match rows {
            Some(rows)
                if rows.len() == row_count
                    && rows.first().copied() == Some(0)
                    && rows.last().copied() == row_count.checked_sub(1) =>
            {
                None
            }
            other => other,
        };

        let grouped = match rows {
            Some(rows) => self.table.group_by_rows(group_by, &planned, rows).ok()?,
            None => self.table.group_by(group_by, &planned).ok()?,
        };
        let grouped_cols = grouped.to_values();
        let group_count = grouped.row_count();

        let mut out: Vec<Vec<Value>> = Vec::with_capacity(group_count);
        for row_idx in 0..group_count {
            let mut row: Vec<Value> = Vec::with_capacity(key_len + plans.len());
            for col in 0..key_len {
                let value = grouped_cols
                    .get(col)
                    .and_then(|c| c.get(row_idx))
                    .cloned()
                    .unwrap_or(formula_columnar::Value::Null);
                row.push(Self::dax_from_columnar(value));
            }
            for plan in &plans {
                let value = match plan {
                    Plan::Direct { col } => {
                        let col = *col;
                        let v = grouped_cols
                            .get(col)
                            .and_then(|c| c.get(row_idx))
                            .cloned()
                            .unwrap_or(formula_columnar::Value::Null);
                        Self::dax_from_columnar(v)
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
        rows: Option<&[usize]>,
    ) -> Option<Vec<Vec<Value>>> {
        use std::collections::HashMap;
        use std::collections::HashSet;

        let row_count = self.table.row_count();

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

            fn update(&mut self, spec: &AggregationSpec, value: Option<&formula_columnar::Value>) {
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
                        if value
                            .and_then(ColumnarTableBackend::numeric_from_columnar)
                            .is_some()
                        {
                            *count += 1;
                        }
                    }
                    (AggState::Sum { sum, count }, AggregationKind::Sum) => {
                        if let Some(v) = value.and_then(ColumnarTableBackend::numeric_from_columnar)
                        {
                            *sum += v;
                            *count += 1;
                        }
                    }
                    (AggState::Avg { sum, count }, AggregationKind::Average) => {
                        if let Some(v) = value.and_then(ColumnarTableBackend::numeric_from_columnar)
                        {
                            *sum += v;
                            *count += 1;
                        }
                    }
                    (AggState::Min { best }, AggregationKind::Min) => {
                        if let Some(v) = value.and_then(ColumnarTableBackend::numeric_from_columnar)
                        {
                            *best = Some(best.map_or(v, |current| current.min(v)));
                        }
                    }
                    (AggState::Max { best }, AggregationKind::Max) => {
                        if let Some(v) = value.and_then(ColumnarTableBackend::numeric_from_columnar)
                        {
                            *best = Some(best.map_or(v, |current| current.max(v)));
                        }
                    }
                    (AggState::DistinctCount { set }, AggregationKind::DistinctCount) => {
                        let Some(v) = value else {
                            return;
                        };
                        set.insert(ColumnarTableBackend::dax_from_columnar(v.clone()));
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
        let mut key_buf: Vec<Value> = Vec::with_capacity(key_len);

        let mut states_template = Vec::with_capacity(aggs.len());
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
            for idx in group_by {
                let col = *idx - col_start;
                let value = range
                    .columns
                    .get(col)
                    .and_then(|c| c.get(row_offset))
                    .cloned()
                    .unwrap_or(formula_columnar::Value::Null);
                key_buf.push(Self::dax_from_columnar(value));
            }

            if let Some(existing) = groups.get_mut(key_buf.as_slice()) {
                for (state, spec) in existing.iter_mut().zip(aggs) {
                    let value = spec.column_idx.and_then(|idx| {
                        let col = idx - col_start;
                        range.columns.get(col).and_then(|c| c.get(row_offset))
                    });
                    state.update(spec, value);
                }
                return;
            }

            let mut states = states_template.clone();
            for (state, spec) in states.iter_mut().zip(aggs) {
                let value = spec.column_idx.and_then(|idx| {
                    let col = idx - col_start;
                    range.columns.get(col).and_then(|c| c.get(row_offset))
                });
                state.update(spec, value);
            }
            groups.insert(key_buf.clone(), states);
        };

        match rows {
            None => {
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
            Some(rows) => {
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
        }

        let mut out = Vec::with_capacity(groups.len());
        for (key, states) in groups {
            let mut row = key;
            for state in states {
                row.push(state.finalize());
            }
            out.push(row);
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
        self.column_index.get(column).copied()
    }

    fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value> {
        if row >= self.table.row_count() || idx >= self.table.column_count() {
            return None;
        }
        let value = self.table.get_cell(row, idx);
        Some(Self::dax_from_columnar(value))
    }

    fn stats_sum(&self, idx: usize) -> Option<f64> {
        self.table.scan().stats(idx)?.sum
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
        Some(Self::dax_from_columnar(stats.min.clone()?))
    }

    fn stats_max(&self, idx: usize) -> Option<Value> {
        let stats = self.table.scan().stats(idx)?;
        Some(Self::dax_from_columnar(stats.max.clone()?))
    }

    fn stats_distinct_count(&self, idx: usize) -> Option<u64> {
        Some(self.table.scan().stats(idx)?.distinct_count)
    }

    fn stats_has_blank(&self, idx: usize) -> Option<bool> {
        let stats = self.table.scan().stats(idx)?;
        Some(stats.null_count > 0)
    }

    fn dictionary_values(&self, idx: usize) -> Option<Vec<Value>> {
        let dict = self.table.dictionary(idx)?;
        Some(dict.iter().cloned().map(Value::from).collect())
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

        match value {
            Value::Text(s) => Some(self.table.scan().filter_eq_string(idx, s.as_ref())),
            Value::Number(n) => match column_type {
                formula_columnar::ColumnType::Number => Some(self.table.scan().filter_eq_number(idx, n.0)),
                formula_columnar::ColumnType::DateTime
                | formula_columnar::ColumnType::Currency { .. }
                | formula_columnar::ColumnType::Percentage { .. } => {
                    let Some(v) = safe_i64_key(n.0) else {
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
        let row_count = self.table.row_count();
        if idx >= self.table.column_count() {
            return None;
        }

        let rows = match rows {
            Some(rows)
                if rows.len() == row_count
                    && rows.first().copied() == Some(0)
                    && rows.last().copied() == row_count.checked_sub(1) =>
            {
                None
            }
            other => other,
        };

        if rows.is_none() {
            if let Some(values) = self.dictionary_values(idx) {
                let mut out: Vec<Value> = values;
                if self.stats_has_blank(idx).unwrap_or(false) {
                    out.push(Value::Blank);
                }
                return Some(out);
            }

            if let Ok(result) = self.table.group_by(&[idx], &[]) {
                let values = result.to_values();
                let mut out = Vec::with_capacity(result.row_count());
                if let Some(col) = values.get(0) {
                    for v in col {
                        out.push(Self::dax_from_columnar(v.clone()));
                    }
                }
                return Some(out);
            }
        }

        if let Some(rows) = rows {
            if let Ok(result) = self.table.group_by_rows(&[idx], &[], rows) {
                let values = result.to_values();
                let mut out = Vec::with_capacity(result.row_count());
                if let Some(col) = values.get(0) {
                    for v in col {
                        out.push(Self::dax_from_columnar(v.clone()));
                    }
                }
                return Some(out);
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

        match rows {
            None => {
                let mut start = 0;
                while start < row_count {
                    let end = (start + CHUNK_ROWS).min(row_count);
                    let range = self.table.get_range(start, end, idx, idx + 1);
                    for v in range.columns.get(0).into_iter().flatten() {
                        push_value(Self::dax_from_columnar(v.clone()));
                    }
                    start = end;
                }
            }
            Some(rows) => {
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
                            push_value(Self::dax_from_columnar(v.clone()));
                        }
                        pos += 1;
                    }
                }
            }
        }

        Some(out)
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

        if let Some(out) = self.group_by_aggregations_query(group_by, aggs, rows) {
            return Some(out);
        }

        self.group_by_aggregations_scan(group_by, aggs, rows)
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
                formula_columnar::ColumnType::DateTime
                | formula_columnar::ColumnType::Currency { .. }
                | formula_columnar::ColumnType::Percentage { .. } => {
                    const MAX_SAFE_INT: f64 = 9_007_199_254_740_992.0; // 2^53
                    let mut ints = Vec::with_capacity(nums.len());
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
                    let dax_value = Self::dax_from_columnar(v.clone());
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
