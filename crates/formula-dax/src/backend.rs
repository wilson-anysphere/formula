use crate::engine::{DaxError, DaxResult};
use crate::value::Value;
use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

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
        let non_blank = self.table.row_count().saturating_sub(stats.null_count as usize);
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
        match value {
            Value::Text(s) => Some(self.table.scan().filter_eq_string(idx, s.as_ref())),
            _ => None,
        }
    }
}
