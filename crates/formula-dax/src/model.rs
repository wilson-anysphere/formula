use crate::backend::{AggregationSpec, ColumnarTableBackend, InMemoryTableBackend, TableBackend};
use crate::engine::{DaxError, DaxResult, FilterContext, RowContext};
use crate::parser::Expr;
use crate::value::Value;
use formula_columnar::{
    BitVec, ColumnSchema as ColumnarColumnSchema, ColumnType as ColumnarColumnType,
    Value as ColumnarValue,
};
use std::collections::hash_map::Entry;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

pub(crate) fn normalize_ident(s: &str) -> String {
    let s = s.trim();
    if s.is_ascii() {
        s.to_ascii_uppercase()
    } else {
        // Use Unicode-aware uppercasing to approximate Excel/Tabular-style case-insensitive
        // identifier matching for non-ASCII names (e.g. ß -> SS).
        s.chars().flat_map(|c| c.to_uppercase()).collect()
    }
}

/// Relationship cardinality between two tables.
///
/// `formula-dax` models relationships in the same oriented way as Tabular/Power Pivot: every
/// relationship has a `from_*` side and a `to_*` side. That orientation is meaningful even for
/// [`Cardinality::ManyToMany`] relationships (see [`Relationship`] for details).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Cardinality {
    OneToMany,
    OneToOne,
    ManyToMany,
}

/// Controls how filters propagate across a relationship.
///
/// In `formula-dax`, `from_table`/`to_table` are oriented such that the default propagation
/// direction is always `to_table → from_table`:
///
/// - [`CrossFilterDirection::Single`]: propagate filters only from `to_table` to `from_table`.
/// - [`CrossFilterDirection::Both`]: propagate filters in both directions.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CrossFilterDirection {
    Single,
    Both,
}

/// A relationship between two tables.
///
/// ## Orientation (`from_*` / `to_*`)
/// `from_table[from_column]` and `to_table[to_column]` form an *oriented* relationship.
///
/// - For [`Cardinality::OneToMany`], `to_*` is the lookup/"one" side (unique key) and `from_*` is
///   the fact/"many" side (foreign key).
/// - For [`Cardinality::OneToOne`], both sides are unique.
/// - For [`Cardinality::ManyToMany`], neither side is required to be unique. The orientation is
///   still meaningful: it defines the default direction for filter propagation (see
///   [`cross_filter_direction`](Relationship::cross_filter_direction)) and which side `RELATED` /
///   `RELATEDTABLE` navigate.
///
/// ## Row-context navigation (`RELATED` / `RELATEDTABLE`)
/// `RELATED` navigates from a row on the `from_table` side to a row on the `to_table` side.
///
/// - For 1:* / 1:1 relationships, this is a single-row lookup.
/// - For many-to-many relationships, the lookup can be ambiguous: if the key matches more than one
///   row in `to_table`, `RELATED` raises an error.
///
/// `RELATEDTABLE` navigates from a row on the `to_table` side to the set of matching rows in
/// `from_table` (and may return multiple rows for both 1:* and *:* relationships).
///
/// ## Filter propagation
/// Relationship propagation is handled by the evaluation engine ([`crate::DaxEngine`]) by
/// repeatedly applying relationship constraints until reaching a fixed point.
///
/// - With [`CrossFilterDirection::Single`], filters propagate from `to_table` to `from_table`.
/// - With [`CrossFilterDirection::Both`], filters propagate in both directions.
/// - For [`Cardinality::ManyToMany`], propagation is based on the **distinct set of visible key
///   values** on the source side (conceptually similar to `TREATAS(VALUES(source[key]),
///   target[key])`), rather than relying on a unique lookup row.
///
/// ## Referential integrity and the implicit blank/unknown member
/// Tabular models treat fact-side rows whose key is BLANK (or has no match in the related table) as
/// belonging to a virtual "(blank)" / "unknown" member on the `to_table` side.
///
/// **Important:** fact-side BLANK values always belong to this relationship-generated blank member,
/// even if a *physical* row exists on the `to_table` side whose key is BLANK. In other words, BLANK
/// is treated as an *unmatchable* relationship key during join/filter propagation and row-context
/// navigation.
///
/// When [`enforce_referential_integrity`](Relationship::enforce_referential_integrity) is `true`,
/// `formula-dax` rejects non-BLANK values on the `from_*` side that have no match on the `to_*`
/// side. This prevents *unmatched* keys from contributing to the virtual blank member, but BLANK
/// values are still allowed and can still make the virtual blank member visible.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Relationship {
    pub name: String,
    /// The "from" table in the oriented relationship.
    ///
    /// For 1:* relationships this is the fact/"many" side (foreign key). For many-to-many this is
    /// still the side that is filtered by default when
    /// [`cross_filter_direction`](Relationship::cross_filter_direction) is
    /// [`CrossFilterDirection::Single`].
    pub from_table: String,
    /// The column in [`from_table`](Relationship::from_table) participating in the relationship.
    pub from_column: String,
    /// The "to" table in the oriented relationship.
    ///
    /// For 1:* relationships this is the lookup/"one" side (unique key). For many-to-many this is
    /// still the default source of filter propagation and the table navigated to by `RELATED`.
    pub to_table: String,
    /// The column in [`to_table`](Relationship::to_table) participating in the relationship.
    pub to_column: String,
    pub cardinality: Cardinality,
    pub cross_filter_direction: CrossFilterDirection,
    pub is_active: bool,
    /// If true, ensure that every non-BLANK `from_column` value exists in `to_column`.
    ///
    /// When this is `false`, keys on the `from_*` side that do not exist on the `to_*` side are
    /// treated as belonging to the implicit blank/unknown member of `to_table` during filter
    /// propagation.
    pub enforce_referential_integrity: bool,
}

#[derive(Clone, Debug)]
pub struct Measure {
    pub name: String,
    pub expression: String,
    pub(crate) parsed: Expr,
}

#[derive(Clone, Debug)]
pub struct CalculatedColumn {
    pub table: String,
    pub name: String,
    pub expression: String,
    pub parsed: Expr,
}

#[derive(Clone, Debug)]
pub struct Table {
    name: String,
    storage: TableStorage,
}

impl Table {
    pub fn new(name: impl Into<String>, columns: Vec<impl Into<String>>) -> Self {
        let name = name.into();
        let columns: Vec<String> = columns.into_iter().map(Into::into).collect();
        Self {
            name,
            storage: TableStorage::InMemory(InMemoryTableBackend::new(columns)),
        }
    }

    pub fn from_columnar(name: impl Into<String>, table: formula_columnar::ColumnarTable) -> Self {
        Self {
            name: name.into(),
            storage: TableStorage::Columnar(ColumnarTableBackend::new(table)),
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    /// If this table is backed by a [`formula_columnar::ColumnarTable`], return the underlying
    /// storage.
    pub fn columnar_table(&self) -> Option<&Arc<formula_columnar::ColumnarTable>> {
        match &self.storage {
            TableStorage::Columnar(backend) => Some(&backend.table),
            _ => None,
        }
    }

    pub fn columns(&self) -> &[String] {
        self.backend().columns()
    }

    pub fn row_count(&self) -> usize {
        self.backend().row_count()
    }

    pub fn push_row(&mut self, row: Vec<Value>) -> DaxResult<()> {
        match &mut self.storage {
            TableStorage::InMemory(backend) => backend.push_row(&self.name, row),
            TableStorage::Columnar(_) => Err(DaxError::Eval(format!(
                "cannot push rows into columnar table {}",
                self.name
            ))),
        }
    }

    pub(crate) fn column_idx(&self, column: &str) -> Option<usize> {
        self.backend().column_index(column)
    }

    pub fn value(&self, row: usize, column: &str) -> Option<Value> {
        let idx = self.column_idx(column)?;
        self.value_by_idx(row, idx)
    }

    pub(crate) fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value> {
        self.backend().value_by_idx(row, idx)
    }

    pub(crate) fn add_column(
        &mut self,
        name: impl Into<String>,
        values: Vec<Value>,
    ) -> DaxResult<()> {
        let name = name.into();
        match &mut self.storage {
            TableStorage::InMemory(backend) => backend.add_column(&self.name, name, values),
            TableStorage::Columnar(backend) => {
                let key = normalize_ident(&name);
                if backend.column_index.contains_key(&key) {
                    return Err(DaxError::DuplicateColumn {
                        table: self.name.clone(),
                        column: name,
                    });
                }

                let expected = backend.table.row_count();
                let actual = values.len();
                if actual != expected {
                    return Err(DaxError::ColumnLengthMismatch {
                        table: self.name.clone(),
                        column: name,
                        expected,
                        actual,
                    });
                }

                // Columnar columns have a single physical type. For now we require that all
                // non-blank values in the calculated column share the same type.
                let column_type = Self::infer_columnar_type(&self.name, &name, &values)?;
                let column_values: Vec<ColumnarValue> =
                    values.iter().map(Self::dax_to_columnar_value).collect();
                let schema = ColumnarColumnSchema {
                    name: name.clone(),
                    column_type,
                };
                let map_append_err = |err| match err {
                    formula_columnar::ColumnAppendError::LengthMismatch { expected, actual } => {
                        DaxError::ColumnLengthMismatch {
                            table: self.name.clone(),
                            column: name.clone(),
                            expected,
                            actual,
                        }
                    }
                    formula_columnar::ColumnAppendError::DuplicateColumn { name: column } => {
                        DaxError::DuplicateColumn {
                            table: self.name.clone(),
                            column,
                        }
                    }
                    other => DaxError::Eval(format!(
                        "failed to append column {}[{}] to columnar table: {other}",
                        self.name, name
                    )),
                };
                // The columnar backend is stored behind an `Arc`.
                //
                // When the `Arc` is uniquely owned, try to unwrap it to avoid cloning the existing
                // columns. When shared, operate on a clone so errors (though unexpected after
                // pre-validation) leave the original table unchanged.
                let updated = if Arc::strong_count(&backend.table) == 1 {
                    let options = backend.table.options();
                    let placeholder = formula_columnar::ColumnarTable::from_encoded(
                        Vec::new(),
                        Vec::new(),
                        0,
                        options,
                    );
                    let table_arc = std::mem::replace(&mut backend.table, Arc::new(placeholder));
                    let table = match Arc::try_unwrap(table_arc) {
                        Ok(table) => table,
                        Err(table) => {
                            debug_assert!(
                                false,
                                "Arc::strong_count == 1 but Arc::try_unwrap failed"
                            );
                            table.as_ref().clone()
                        }
                    };
                    table
                        .with_appended_column(schema, column_values)
                        .map_err(map_append_err)?
                } else {
                    backend
                        .table
                        .as_ref()
                        .clone()
                        .with_appended_column(schema, column_values)
                        .map_err(map_append_err)?
                };

                backend.table = Arc::new(updated);
                backend.columns = backend
                    .table
                    .schema()
                    .iter()
                    .map(|c| c.name.clone())
                    .collect();
                backend.column_index = backend
                    .columns
                    .iter()
                    .enumerate()
                    .map(|(idx, c)| (normalize_ident(c), idx))
                    .collect();
                Ok(())
            }
        }
    }

    pub(crate) fn set_value_by_idx(
        &mut self,
        row: usize,
        idx: usize,
        value: Value,
    ) -> DaxResult<()> {
        match &mut self.storage {
            TableStorage::InMemory(backend) => backend.set_value_by_idx(row, idx, value),
            TableStorage::Columnar(_) => Err(DaxError::Eval(format!(
                "cannot mutate columnar table {}",
                self.name
            ))),
        }
    }

    pub(crate) fn pop_row(&mut self) -> Option<Vec<Value>> {
        match &mut self.storage {
            TableStorage::InMemory(backend) => backend.rows.pop(),
            TableStorage::Columnar(_) => None,
        }
    }

    pub(crate) fn pop_last_column(&mut self) -> DaxResult<Option<String>> {
        match &mut self.storage {
            TableStorage::InMemory(backend) => {
                let name = match backend.columns.pop() {
                    Some(name) => name,
                    None => return Ok(None),
                };
                backend.column_index.remove(&normalize_ident(&name));
                for row in &mut backend.rows {
                    row.pop();
                }
                Ok(Some(name))
            }
            TableStorage::Columnar(_) => Err(DaxError::Eval(format!(
                "cannot remove columns from columnar table {}",
                self.name
            ))),
        }
    }

    fn backend(&self) -> &dyn TableBackend {
        match &self.storage {
            TableStorage::InMemory(backend) => backend,
            TableStorage::Columnar(backend) => backend,
        }
    }

    fn infer_columnar_type(
        table: &str,
        column: &str,
        values: &[Value],
    ) -> DaxResult<ColumnarColumnType> {
        let mut ty: Option<ColumnarColumnType> = None;
        for v in values {
            let this = match v {
                Value::Blank => continue,
                Value::Number(_) => ColumnarColumnType::Number,
                Value::Text(_) => ColumnarColumnType::String,
                Value::Boolean(_) => ColumnarColumnType::Boolean,
            };
            match ty {
                Some(existing) if existing != this => {
                    return Err(DaxError::Type(format!(
                        "calculated column {table}[{column}] must have a single type; saw both {existing:?} and {this:?}"
                    )));
                }
                None => ty = Some(this),
                _ => {}
            }
        }

        // If the column is entirely blank, the physical type is unobservable. Default to `Number`
        // since it tends to be the most permissive for aggregations.
        Ok(ty.unwrap_or(ColumnarColumnType::Number))
    }

    fn dax_to_columnar_value(value: &Value) -> ColumnarValue {
        match value {
            Value::Blank => ColumnarValue::Null,
            Value::Number(n) => ColumnarValue::Number(n.0),
            Value::Text(s) => ColumnarValue::String(s.clone()),
            Value::Boolean(b) => ColumnarValue::Boolean(*b),
        }
    }
}

impl TableBackend for Table {
    fn columns(&self) -> &[String] {
        self.backend().columns()
    }

    fn row_count(&self) -> usize {
        self.backend().row_count()
    }

    fn column_index(&self, column: &str) -> Option<usize> {
        self.backend().column_index(column)
    }

    fn value_by_idx(&self, row: usize, idx: usize) -> Option<Value> {
        self.backend().value_by_idx(row, idx)
    }

    fn stats_sum(&self, idx: usize) -> Option<f64> {
        self.backend().stats_sum(idx)
    }

    fn stats_non_blank_count(&self, idx: usize) -> Option<usize> {
        self.backend().stats_non_blank_count(idx)
    }

    fn stats_min(&self, idx: usize) -> Option<Value> {
        self.backend().stats_min(idx)
    }

    fn stats_max(&self, idx: usize) -> Option<Value> {
        self.backend().stats_max(idx)
    }

    fn stats_distinct_count(&self, idx: usize) -> Option<u64> {
        self.backend().stats_distinct_count(idx)
    }

    fn stats_has_blank(&self, idx: usize) -> Option<bool> {
        self.backend().stats_has_blank(idx)
    }

    fn dictionary_values(&self, idx: usize) -> Option<Vec<Value>> {
        self.backend().dictionary_values(idx)
    }

    fn filter_eq(&self, idx: usize, value: &Value) -> Option<Vec<usize>> {
        self.backend().filter_eq(idx, value)
    }

    fn distinct_values_filtered(&self, idx: usize, rows: Option<&[usize]>) -> Option<Vec<Value>> {
        self.backend().distinct_values_filtered(idx, rows)
    }

    fn distinct_values_filtered_mask(
        &self,
        idx: usize,
        mask: Option<&BitVec>,
    ) -> Option<Vec<Value>> {
        self.backend().distinct_values_filtered_mask(idx, mask)
    }

    fn group_by_aggregations(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        rows: Option<&[usize]>,
    ) -> Option<Vec<Vec<Value>>> {
        self.backend().group_by_aggregations(group_by, aggs, rows)
    }

    fn group_by_aggregations_mask(
        &self,
        group_by: &[usize],
        aggs: &[AggregationSpec],
        mask: Option<&BitVec>,
    ) -> Option<Vec<Vec<Value>>> {
        self.backend()
            .group_by_aggregations_mask(group_by, aggs, mask)
    }

    fn filter_in(&self, idx: usize, values: &[Value]) -> Option<Vec<usize>> {
        self.backend().filter_in(idx, values)
    }

    fn columnar_table(&self) -> Option<&formula_columnar::ColumnarTable> {
        self.backend().columnar_table()
    }

    fn hash_join(
        &self,
        right: &dyn TableBackend,
        left_on: usize,
        right_on: usize,
    ) -> Option<formula_columnar::JoinResult> {
        self.backend().hash_join(right, left_on, right_on)
    }
}

#[derive(Clone, Debug)]
enum TableStorage {
    InMemory(InMemoryTableBackend),
    Columnar(ColumnarTableBackend),
}

#[derive(Clone, Debug)]
pub struct DataModel {
    pub(crate) tables: HashMap<String, Table>,
    pub(crate) relationships: Vec<RelationshipInfo>,
    pub(crate) measures: HashMap<String, Measure>,
    pub(crate) calculated_columns: Vec<CalculatedColumn>,
    pub(crate) calculated_column_order: HashMap<String, Vec<usize>>,
}

/// A compact representation of the set of rows that share the same relationship key on the
/// [`Relationship::to_table`] side of a relationship.
///
/// - For [`Cardinality::OneToMany`] and [`Cardinality::OneToOne`], keys are unique on the `to_table`
///   side and we store a single row index via [`RowSet::One`] without allocating.
/// - For [`Cardinality::ManyToMany`], keys can map to multiple rows on the `to_table` side and we
///   store all matching rows via [`RowSet::Many`].
#[derive(Clone, Debug)]
pub(crate) enum RowSet {
    One(usize),
    Many(Vec<usize>),
}

/// Lookup structure for the relationship key on the [`Relationship::to_table`] side.
///
/// For many-to-many relationships the `to_table` side may contain a very large number of duplicate
/// keys (especially when backed by a columnar table). Storing `Vec<usize>` row lists per key can
/// become a major memory sink.
///
/// When the `to_table` is columnar and the relationship is many-to-many, we store only the set of
/// distinct keys and rely on [`crate::backend::TableBackend::filter_eq`] /
/// [`crate::backend::TableBackend::filter_in`] to retrieve matching row indices on demand.
#[derive(Clone, Debug)]
pub(crate) enum ToIndex {
    /// Full mapping from key -> row set(s).
    RowSets {
        map: HashMap<Value, RowSet>,
        /// Whether any **non-blank** key maps to more than one row.
        ///
        /// Physical BLANK keys do not participate in relationship joins (fact-side BLANK foreign
        /// keys map to the relationship-generated virtual blank member). We therefore ignore
        /// duplicate BLANK keys when deciding whether relationship traversal requires a
        /// many-to-many expansion algorithm.
        has_duplicates: bool,
    },
    /// Scalable representation for columnar many-to-many `to_table` lookups.
    KeySet {
        keys: HashSet<Value>,
        /// Whether any **non-blank** key occurs more than once in the `to_table`.
        ///
        /// Physical BLANK keys do not participate in relationship joins (fact-side BLANK foreign
        /// keys map to the relationship-generated virtual blank member). We therefore ignore
        /// duplicate BLANK keys when deciding whether relationship traversal requires a
        /// many-to-many expansion algorithm.
        has_duplicates: bool,
    },
}

impl ToIndex {
    pub(crate) fn contains_key(&self, key: &Value) -> bool {
        match self {
            ToIndex::RowSets { map, .. } => map.contains_key(key),
            ToIndex::KeySet { keys, .. } => keys.contains(key),
        }
    }

    pub(crate) fn has_duplicates(&self) -> bool {
        match self {
            ToIndex::RowSets { has_duplicates, .. } => *has_duplicates,
            ToIndex::KeySet { has_duplicates, .. } => *has_duplicates,
        }
    }
}

impl RowSet {
    pub(crate) fn push(&mut self, row: usize) {
        match self {
            RowSet::One(existing) => {
                let mut rows = Vec::new();
                let _ = rows.try_reserve_exact(2);
                rows.push(*existing);
                rows.push(row);
                *self = RowSet::Many(rows);
            }
            RowSet::Many(rows) => rows.push(row),
        }
    }

    pub(crate) fn any_allowed(&self, allowed: &BitVec) -> bool {
        match self {
            RowSet::One(row) => *row < allowed.len() && allowed.get(*row),
            RowSet::Many(rows) => rows
                .iter()
                .any(|row| *row < allowed.len() && allowed.get(*row)),
        }
    }

    pub(crate) fn for_each_row(&self, mut f: impl FnMut(usize)) {
        match self {
            RowSet::One(row) => f(*row),
            RowSet::Many(rows) => {
                for &row in rows {
                    f(row);
                }
            }
        }
    }
}

/// A compact representation of fact-side rows that belong to the "virtual blank" dimension member:
/// rows whose foreign key is BLANK or does not map to any key in the related `to_table`.
///
/// For columnar fact tables we avoid materializing a full `FK -> Vec<row>` index; however, the
/// engine still needs to model Tabular's implicit blank/unknown member semantics. Storing a row
/// list for the blank member can still be too large when most rows are unmatched, so we switch
/// from a sparse list to a dense bitset when it becomes more memory-efficient.
#[derive(Clone, Debug)]
pub(crate) enum UnmatchedFactRows {
    Sparse(Vec<usize>),
    Dense {
        /// Bitset of length `len` (stored in 64-bit words).
        bits: Vec<u64>,
        len: usize,
        count: usize,
    },
}

impl UnmatchedFactRows {
    pub(crate) fn is_empty(&self) -> bool {
        match self {
            UnmatchedFactRows::Sparse(rows) => rows.is_empty(),
            UnmatchedFactRows::Dense { count, .. } => *count == 0,
        }
    }

    #[allow(dead_code)]
    pub(crate) fn retain(&mut self, mut keep: impl FnMut(&usize) -> bool) {
        match self {
            UnmatchedFactRows::Sparse(rows) => rows.retain(|row| keep(row)),
            UnmatchedFactRows::Dense { bits, len, count } => {
                let len = *len;
                let mut new_count = 0usize;
                for (word_idx, word) in bits.iter_mut().enumerate() {
                    let mut w = *word;
                    let mut new_word = 0u64;
                    while w != 0 {
                        let tz = w.trailing_zeros() as usize;
                        let row = word_idx * 64 + tz;
                        if row >= len {
                            break;
                        }
                        let mask = 1u64 << tz;
                        if keep(&row) {
                            new_word |= mask;
                            new_count += 1;
                        }
                        w &= w - 1;
                    }
                    *word = new_word;
                }
                *count = new_count;
            }
        }
    }

    pub(crate) fn push_row(&mut self, row: usize, new_len: usize, is_unmatched: bool) {
        match self {
            UnmatchedFactRows::Sparse(rows) => {
                if is_unmatched {
                    rows.push(row);
                }

                // Compare the approximate memory usage of:
                // - sparse list: `unmatched_count * size_of::<usize>()`
                // - dense bitset: `row_count / 8` bytes
                //
                // Switch to the dense representation once it becomes more memory-efficient:
                //   unmatched_count > row_count / 64.
                let sparse_to_dense_threshold = new_len / 64;
                if rows.len() > sparse_to_dense_threshold {
                    let word_len = (new_len + 63) / 64;
                    let mut bits = vec![0u64; word_len];
                    let mut count = 0usize;
                    for &row in rows.iter() {
                        let word = row / 64;
                        let bit = row % 64;
                        let mask = 1u64 << bit;
                        if (bits[word] & mask) == 0 {
                            bits[word] |= mask;
                            count += 1;
                        }
                    }
                    *self = UnmatchedFactRows::Dense {
                        bits,
                        len: new_len,
                        count,
                    };
                }
            }
            UnmatchedFactRows::Dense { bits, len, count } => {
                if new_len > *len {
                    let word_len = (new_len + 63) / 64;
                    if bits.len() < word_len {
                        bits.resize(word_len, 0u64);
                    }
                    *len = new_len;
                }
                if is_unmatched {
                    let word = row / 64;
                    let bit = row % 64;
                    let mask = 1u64 << bit;
                    if (bits[word] & mask) == 0 {
                        bits[word] |= mask;
                        *count += 1;
                    }
                }
            }
        }
    }

    pub(crate) fn clear_row(&mut self, row: usize) {
        match self {
            UnmatchedFactRows::Sparse(rows) => {
                if let Some(pos) = rows.iter().position(|&r| r == row) {
                    rows.swap_remove(pos);
                }
            }
            UnmatchedFactRows::Dense { bits, len, count } => {
                if row >= *len {
                    return;
                }
                let word = row / 64;
                let bit = row % 64;
                let mask = 1u64 << bit;
                if (bits[word] & mask) != 0 {
                    bits[word] &= !mask;
                    *count = count.saturating_sub(1);
                }
            }
        }
    }
    pub(crate) fn for_each_row(&self, mut f: impl FnMut(usize)) {
        match self {
            UnmatchedFactRows::Sparse(rows) => {
                for &row in rows {
                    f(row);
                }
            }
            UnmatchedFactRows::Dense { bits, len, .. } => {
                for (word_idx, &word) in bits.iter().enumerate() {
                    let mut w = word;
                    while w != 0 {
                        let tz = w.trailing_zeros() as usize;
                        let row = word_idx * 64 + tz;
                        if row >= *len {
                            break;
                        }
                        f(row);
                        w &= w - 1;
                    }
                }
            }
        }
    }
    pub(crate) fn extend_into(&self, out: &mut Vec<usize>) {
        match self {
            UnmatchedFactRows::Sparse(rows) => out.extend(rows.iter().copied()),
            UnmatchedFactRows::Dense { .. } => self.for_each_row(|row| out.push(row)),
        }
    }

    pub(crate) fn any_row_allowed(&self, allowed: &BitVec) -> bool {
        match self {
            UnmatchedFactRows::Sparse(rows) => rows
                .iter()
                .any(|row| *row < allowed.len() && allowed.get(*row)),
            UnmatchedFactRows::Dense { bits, .. } => {
                // Fast path: intersect the dense "unmatched" bitmap with the current allowed set.
                let allowed_words = allowed.as_words();
                let min_words = bits.len().min(allowed_words.len());
                for i in 0..min_words {
                    if (bits[i] & allowed_words[i]) != 0 {
                        return true;
                    }
                }
                false
            }
        }
    }
}

struct UnmatchedFactRowsBuilder {
    row_count: usize,
    sparse_to_dense_threshold: usize,
    rows: UnmatchedFactRows,
}

impl UnmatchedFactRowsBuilder {
    fn new(row_count: usize) -> Self {
        // Compare the approximate memory usage of:
        // - sparse list: `unmatched_count * size_of::<usize>()`
        // - dense bitset: `row_count / 8` bytes
        //
        // We switch to the dense representation once it becomes more memory-efficient:
        //   unmatched_count > row_count / 64.
        let sparse_to_dense_threshold = row_count / 64;
        Self {
            row_count,
            sparse_to_dense_threshold,
            rows: UnmatchedFactRows::Sparse(Vec::new()),
        }
    }

    fn push(&mut self, row: usize) {
        match &mut self.rows {
            UnmatchedFactRows::Sparse(rows) => {
                rows.push(row);
                if rows.len() > self.sparse_to_dense_threshold {
                    let word_len = (self.row_count + 63) / 64;
                    let mut bits = vec![0u64; word_len];
                    let mut count = 0usize;
                    for &row in rows.iter() {
                        let word = row / 64;
                        let bit = row % 64;
                        let mask = 1u64 << bit;
                        if (bits[word] & mask) == 0 {
                            bits[word] |= mask;
                            count += 1;
                        }
                    }
                    self.rows = UnmatchedFactRows::Dense {
                        bits,
                        len: self.row_count,
                        count,
                    };
                }
            }
            UnmatchedFactRows::Dense { bits, count, .. } => {
                let word = row / 64;
                let bit = row % 64;
                let mask = 1u64 << bit;
                if (bits[word] & mask) == 0 {
                    bits[word] |= mask;
                    *count += 1;
                }
            }
        }
    }

    fn finish(self) -> UnmatchedFactRows {
        self.rows
    }
}
/// Internal relationship representation with precomputed key indices used by filter propagation
/// and row-context navigation.
///
/// - `to_index` provides lookups for keys in `to_table[to_column]`.
///   - For most relationships, it maps each key value to the set of matching row indices in
///     `to_table` (see [`RowSet`] for how this is represented compactly).
///   - For many-to-many relationships where `to_table` is columnar, it stores only the set of
///     distinct keys and relies on backend filter primitives to retrieve row indices on demand.
/// - `from_index`, when present, maps each key value in `from_table[from_column]` to the list of
///   matching row indices in `from_table`. Columnar fact tables omit this index and rely on
///   backend filter primitives instead.
#[derive(Clone, Debug)]
pub(crate) struct RelationshipInfo {
    pub(crate) rel: Relationship,
    pub(crate) from_table_key: String,
    pub(crate) from_column_key: String,
    pub(crate) to_table_key: String,
    pub(crate) to_column_key: String,
    /// Column index of `rel.from_column` in the `from_table`.
    pub(crate) from_idx: usize,
    /// Column index of `rel.to_column` in the `to_table`.
    pub(crate) to_idx: usize,
    pub(crate) to_index: ToIndex,
    /// Relationship index for the fact-side (from_table) foreign key.
    ///
    /// For in-memory fact tables, we build an index of `FK -> fact row indices` to enable fast
    /// relationship navigation (e.g. `RELATEDTABLE`) and filter propagation.
    ///
    /// For columnar fact tables, storing row vectors for every key is prohibitively expensive.
    /// In that case this stays `None` and the engine relies on columnar primitives such as
    /// [`TableBackend::filter_eq`], [`TableBackend::filter_in`], and
    /// [`TableBackend::distinct_values_filtered`].
    pub(crate) from_index: Option<HashMap<Value, Vec<usize>>>,

    /// Fact-side row indices whose foreign key is BLANK or does not map to any key in
    /// [`Self::to_index`]. These rows belong to the relationship's "virtual blank" member on the
    /// dimension side.
    ///
    /// The engine uses this cache to implement Tabular's unknown-member semantics efficiently,
    /// without scanning `from_index` / the fact table to determine blank-row existence or to
    /// enumerate unmatched fact rows.
    pub(crate) unmatched_fact_rows: Option<UnmatchedFactRows>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum RelationshipPathDirection {
    /// Follow relationships in their defined direction: `from_table -> to_table`.
    ///
    /// The name reflects the common 1:* star-schema case where `from_table` is the many-side table
    /// and `to_table` is the one-side table, but the direction is meaningful for
    /// [`Cardinality::ManyToMany`] as well.
    ManyToOne,
    /// Follow relationships in reverse: `to_table -> from_table`.
    OneToMany,
}

impl DataModel {
    pub fn new() -> Self {
        Self {
            tables: HashMap::new(),
            relationships: Vec::new(),
            measures: HashMap::new(),
            calculated_columns: Vec::new(),
            calculated_column_order: HashMap::new(),
        }
    }

    pub fn table(&self, name: &str) -> Option<&Table> {
        let key = normalize_ident(name);
        self.tables.get(&key)
    }

    pub fn tables(&self) -> impl Iterator<Item = &Table> {
        self.tables.values()
    }

    pub fn relationships_definitions(&self) -> impl Iterator<Item = &Relationship> {
        self.relationships.iter().map(|r| &r.rel)
    }

    pub fn measures_definitions(&self) -> impl Iterator<Item = &Measure> {
        self.measures.values()
    }

    pub fn calculated_columns(&self) -> &[CalculatedColumn] {
        &self.calculated_columns
    }

    pub fn add_table(&mut self, table: Table) -> DaxResult<()> {
        let name = table.name.clone();
        let key = normalize_ident(&name);
        if self.tables.contains_key(&key) {
            return Err(DaxError::DuplicateTable { table: name });
        }

        // Tabular/DAX identifiers are case-insensitive. Ensure we don't allow two physical columns
        // that normalize to the same identifier, which would make subsequent column resolution
        // ambiguous.
        let mut seen_cols: HashSet<String> = HashSet::new();
        for column in table.columns() {
            let col_key = normalize_ident(column);
            if !seen_cols.insert(col_key) {
                return Err(DaxError::DuplicateColumn {
                    table: name.clone(),
                    column: column.clone(),
                });
            }
        }
        self.tables.insert(key, table);
        Ok(())
    }

    pub fn insert_row(&mut self, table: &str, row: Vec<Value>) -> DaxResult<()> {
        let table_key = normalize_ident(table);
        let table_ref = self
            .tables
            .get(&table_key)
            .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;

        let calc_count = self
            .calculated_columns
            .iter()
            .filter(|c| normalize_ident(&c.table) == table_key)
            .count();

        let total_columns = table_ref.columns().len();
        let base_columns = total_columns.saturating_sub(calc_count);

        let mut full_row = match row.len() {
            n if n == total_columns => row,
            n if n == base_columns => {
                // Insert values for non-calculated columns in schema order and leave calculated
                // column slots blank. This ensures `insert_row` works even when calculated columns
                // are not physically stored at the end of the table (e.g. in persisted models).
                let calc_names: HashSet<String> = self
                    .calculated_columns
                    .iter()
                    .filter(|c| normalize_ident(&c.table) == table_key)
                    .map(|c| normalize_ident(&c.name))
                    .collect();
                let mut iter = row.into_iter();
                let mut expanded = Vec::new();
                let _ = expanded.try_reserve_exact(total_columns);
                for col in table_ref.columns() {
                    if calc_names.contains(&normalize_ident(col)) {
                        expanded.push(Value::Blank);
                    } else {
                        expanded.push(iter.next().unwrap_or(Value::Blank));
                    }
                }
                expanded
            }
            actual => {
                return Err(DaxError::SchemaMismatch {
                    table: table_ref.name.clone(),
                    expected: base_columns,
                    actual,
                })
            }
        };

        let row_index = {
            let table_mut = self
                .tables
                .get_mut(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            table_mut.push_row(full_row.clone())?;
            table_mut.row_count().saturating_sub(1)
        };

        if calc_count > 0 {
            let calc_result: DaxResult<Vec<Value>> = (|| {
                let mut row_ctx = RowContext::default();
                row_ctx.push(&table_key, row_index);
                let engine = crate::engine::DaxEngine::new();

                let topo_order = match self.calculated_column_order.get(&table_key) {
                    Some(order) if order.len() == calc_count => order.clone(),
                    _ => {
                        let order = self.build_calculated_column_order(&table_key)?;
                        self.calculated_column_order
                            .insert(table_key.clone(), order.clone());
                        order
                    }
                };

                let calc_defs: Vec<CalculatedColumn> = topo_order
                    .into_iter()
                    .filter_map(|idx| self.calculated_columns.get(idx).cloned())
                    .collect();

                for calc in calc_defs {
                    let value = engine.evaluate_expr(
                        self,
                        &calc.parsed,
                        &FilterContext::default(),
                        &row_ctx,
                    )?;
                    let col_idx = {
                        let table_ref = self
                            .tables
                            .get(&table_key)
                            .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
                        table_ref
                            .column_idx(&calc.name)
                            .ok_or_else(|| DaxError::UnknownColumn {
                                table: table.to_string(),
                                column: calc.name.clone(),
                            })?
                    };

                    let table_mut = self
                        .tables
                        .get_mut(&table_key)
                        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
                    table_mut.set_value_by_idx(row_index, col_idx, value)?;
                }

                Ok(self
                    .tables
                    .get(&table_key)
                    .map(|t| {
                        (0..t.columns().len())
                            .map(|idx| t.value_by_idx(row_index, idx).unwrap_or(Value::Blank))
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_else(|| full_row.clone()))
            })();

            match calc_result {
                Ok(updated_row) => full_row = updated_row,
                Err(err) => {
                    // Keep insert_row atomic: if computing a calculated column fails, remove the
                    // appended row before returning the error.
                    if let Some(table_mut) = self.tables.get_mut(&table_key) {
                        table_mut.pop_row();
                    }
                    return Err(err);
                }
            }
        }

        // Collect updates to relationship indexes. We stage these so we can validate referential
        // integrity / uniqueness before mutating any relationship state.
        //
        // Each entry is:
        // - relationship index
        // - inserted key value on the `to_table` side
        // - whether that key was already present in the relationship's `to_index` before insert
        let mut to_index_updates: Vec<(usize, Value, bool)> = Vec::new();
        for (rel_idx, rel_info) in self.relationships.iter().enumerate() {
            let rel = &rel_info.rel;
            if rel_info.to_table_key == table_key {
                let key = full_row
                    .get(rel_info.to_idx)
                    .cloned()
                    .unwrap_or(Value::Blank);
                let key_existed = rel_info.to_index.contains_key(&key);
                // Keys on the "to" side must be unique for 1:* and 1:1 relationships.
                if rel.cardinality != Cardinality::ManyToMany && key_existed {
                    self.tables
                        .get_mut(&table_key)
                        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?
                        .pop_row();
                    return Err(DaxError::NonUniqueKey {
                        table: rel.to_table.clone(),
                        column: rel.to_column.clone(),
                        value: key,
                    });
                }
                to_index_updates.push((rel_idx, key, key_existed));
            }

            if rel_info.from_table_key == table_key {
                let key = full_row
                    .get(rel_info.from_idx)
                    .cloned()
                    .unwrap_or(Value::Blank);

                // For 1:1 relationships, the "from" side must also be unique. We treat BLANK as
                // a real key for uniqueness, matching the existing to-side semantics.
                if rel.cardinality == Cardinality::OneToOne {
                    if let Some(from_index) = rel_info.from_index.as_ref() {
                        if from_index.contains_key(&key) {
                            self.tables
                                .get_mut(&table_key)
                                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?
                                .pop_row();
                            return Err(DaxError::NonUniqueKey {
                                table: rel.from_table.clone(),
                                column: rel.from_column.clone(),
                                value: key,
                            });
                        }
                    }
                }

                if !rel.enforce_referential_integrity {
                    continue;
                }
                if key.is_blank() {
                    continue;
                }
                if !rel_info.to_index.contains_key(&key) {
                    self.tables
                        .get_mut(&table_key)
                        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?
                        .pop_row();
                    return Err(DaxError::ReferentialIntegrityViolation {
                        relationship: rel.name.clone(),
                        from_table: rel.from_table.clone(),
                        from_column: rel.from_column.clone(),
                        to_table: rel.to_table.clone(),
                        to_column: rel.to_column.clone(),
                        value: key,
                    });
                }
            }
        }

        for (rel_idx, key, key_existed) in to_index_updates {
            // Keep a copy for any downstream bookkeeping (e.g. updating cached unmatched fact
            // rows for columnar relationships).
            let key_for_updates = key.clone();
            let cardinality = self.relationships[rel_idx].rel.cardinality;
            match &mut self.relationships[rel_idx].to_index {
                ToIndex::RowSets {
                    map,
                    has_duplicates,
                } => match cardinality {
                    Cardinality::OneToMany | Cardinality::OneToOne => {
                        map.insert(key, RowSet::One(row_index));
                    }
                    Cardinality::ManyToMany => {
                        // Physical BLANK keys do not participate in relationship joins (fact-side BLANK
                        // foreign keys map to the relationship-generated virtual blank member). Skip
                        // materializing row lists for BLANK to avoid a potentially huge, never-used
                        // vector when the dimension contains many BLANK keys.
                        if key.is_blank() {
                            // no-op
                        } else {
                            match map.entry(key) {
                                Entry::Vacant(v) => {
                                    v.insert(RowSet::One(row_index));
                                }
                                Entry::Occupied(mut o) => {
                                    *has_duplicates = true;
                                    o.get_mut().push(row_index);
                                }
                            }
                        }
                    }
                },
                ToIndex::KeySet {
                    keys,
                    has_duplicates,
                } => {
                    // Columnar tables are immutable, so this should be unreachable. Keep the
                    // structure consistent in case a mutable columnar backend is introduced.
                    if key.is_blank() {
                        // BLANK keys do not participate in joins.
                    } else if !keys.insert(key) {
                        *has_duplicates = true;
                    }
                }
            }

            // If this relationship tracks unmatched fact rows (columnar from-table), inserting a
            // new dimension key can "resolve" some of those rows. Remove any fact rows whose FK now
            // matches the inserted key so they no longer belong to the virtual blank member.
            // Facts whose FK is BLANK always belong to the virtual blank member, even if a
            // physical BLANK key exists on the dimension side.
            if key_existed || key_for_updates.is_blank() {
                continue;
            }

            let Some(rel_info) = self.relationships.get_mut(rel_idx) else {
                debug_assert!(false, "relationship index from updates out of bounds");
                continue;
            };
            let Some(unmatched) = rel_info.unmatched_fact_rows.as_mut() else {
                continue;
            };

            let from_table_key = rel_info.from_table_key.clone();
            let from_idx = rel_info.from_idx;
            let from_table_ref = self
                .tables
                .get(&from_table_key)
                .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.from_table.clone()))?;
            match unmatched {
                UnmatchedFactRows::Sparse(rows) => {
                    // When the unmatched set is sparse, scanning it is cheaper than finding all
                    // matches and removing them.
                    let key = &key_for_updates;
                    rows.retain(|row| {
                        let v = from_table_ref
                            .value_by_idx(*row, from_idx)
                            .unwrap_or(Value::Blank);
                        v.is_blank() || &v != key
                    });
                }
                UnmatchedFactRows::Dense { .. } => {
                    let key = &key_for_updates;
                    // Prefer the precomputed fact-side index when available (in-memory fact
                    // tables). For columnar facts, fall back to backend filtering.
                    if let Some(from_index) = rel_info.from_index.as_ref() {
                        if let Some(rows) = from_index.get(key) {
                            for &row in rows {
                                unmatched.clear_row(row);
                            }
                        }
                    } else if let Some(rows) = from_table_ref.filter_eq(from_idx, key) {
                        for row in rows {
                            unmatched.clear_row(row);
                        }
                    } else {
                        // Fallback: scan the unmatched set and drop any rows whose FK now matches
                        // the inserted key. (FK BLANK values always belong to the blank member.)
                        unmatched.retain(|row| {
                            let v = from_table_ref
                                .value_by_idx(*row, from_idx)
                                .unwrap_or(Value::Blank);
                            v.is_blank() || &v != key
                        });
                    }
                }
            }

            if unmatched.is_empty() {
                rel_info.unmatched_fact_rows = None;
            }
        }

        for rel_info in &mut self.relationships {
            if rel_info.from_table_key == table_key {
                let key = full_row
                    .get(rel_info.from_idx)
                    .cloned()
                    .unwrap_or(Value::Blank);
                if let Some(from_index) = rel_info.from_index.as_mut() {
                    from_index.entry(key.clone()).or_default().push(row_index);
                }

                // Maintain the cached set of "virtual blank member" fact rows so `VALUES` /
                // `DISTINCTCOUNT` and relationship propagation don't need to scan relationship
                // indexes to determine blank-row existence.
                let is_unmatched = key.is_blank() || !rel_info.to_index.contains_key(&key);
                let new_len = row_index + 1;
                match rel_info.unmatched_fact_rows.as_mut() {
                    Some(unmatched) => {
                        unmatched.push_row(row_index, new_len, is_unmatched);
                        if unmatched.is_empty() {
                            rel_info.unmatched_fact_rows = None;
                        }
                    }
                    None if is_unmatched => {
                        let mut builder = UnmatchedFactRowsBuilder::new(new_len);
                        builder.push(row_index);
                        rel_info.unmatched_fact_rows = Some(builder.finish());
                    }
                    None => {}
                }
            }
        }

        Ok(())
    }

    /// Add a relationship between two tables.
    ///
    /// Relationship join columns must have compatible types.
    ///
    /// Tabular/Power Pivot relationships with mismatched join column types typically fail
    /// silently (filters don't propagate and functions like `RELATED`/`RELATEDTABLE` appear to
    /// return no matches). To avoid confusing runtime behavior, `formula-dax` performs a
    /// best-effort type compatibility check when relationships are added:
    ///
    /// - **Columnar tables**: use the declared [`formula_columnar::ColumnType`] for each join
    ///   column and compare their *join kind* (Numeric/Text/Boolean). Numeric-like columnar
    ///   types (`Number`, `DateTime`, `Currency`, `Percentage`) are considered compatible.
    /// - **In-memory tables**: scan up to 1k rows for the first non-BLANK value in each join
    ///   column and compare the [`Value`] variant. If either side is all BLANKs in the scan
    ///   window, validation is skipped.
    pub fn add_relationship(&mut self, relationship: Relationship) -> DaxResult<()> {
        let mut relationship = relationship;
        let from_table_lookup_key = normalize_ident(&relationship.from_table);
        let to_table_lookup_key = normalize_ident(&relationship.to_table);
        let from_table = self
            .tables
            .get(&from_table_lookup_key)
            .ok_or_else(|| DaxError::UnknownTable(relationship.from_table.clone()))?;
        let to_table = self
            .tables
            .get(&to_table_lookup_key)
            .ok_or_else(|| DaxError::UnknownTable(relationship.to_table.clone()))?;

        let from_col = relationship.from_column.clone();
        let to_col = relationship.to_column.clone();

        let from_idx = from_table
            .column_idx(&from_col)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: from_table.name.clone(),
                column: from_col.clone(),
            })?;
        let to_idx = to_table
            .column_idx(&to_col)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: to_table.name.clone(),
                column: to_col.clone(),
            })?;

        relationship.from_table = from_table.name.clone();
        relationship.to_table = to_table.name.clone();
        relationship.from_column = from_table
            .columns()
            .get(from_idx)
            .cloned()
            .unwrap_or(from_col.clone());
        relationship.to_column = to_table
            .columns()
            .get(to_idx)
            .cloned()
            .unwrap_or(to_col.clone());

        Self::validate_relationship_join_column_types(
            &relationship,
            from_table,
            from_idx,
            to_table,
            to_idx,
        )?;

        let from_table_key = normalize_ident(&relationship.from_table);
        let to_table_key = normalize_ident(&relationship.to_table);
        let from_column_key = normalize_ident(&relationship.from_column);
        let to_column_key = normalize_ident(&relationship.to_column);

        let to_index = match (&to_table.storage, relationship.cardinality) {
            // For many-to-many relationships where the `to_table` is columnar, storing per-key row
            // lists can be prohibitively expensive when keys are highly duplicated. Store only the
            // key set and rely on backend filter primitives to retrieve row indices on demand.
            (TableStorage::Columnar(_), Cardinality::ManyToMany) => {
                // Prefer backend distinct-value enumeration so we don't hash every row when the
                // `to_table` is highly duplicated (a common columnar fact-table pattern).
                let distinct_values = to_table
                    .distinct_values_filtered(to_idx, None)
                    .unwrap_or_else(|| {
                        let mut seen = HashSet::<Value>::new();
                        let mut out = Vec::new();
                        for row in 0..to_table.row_count() {
                            let value = to_table.value_by_idx(row, to_idx).unwrap_or(Value::Blank);
                            if seen.insert(value.clone()) {
                                out.push(value);
                            }
                        }
                        out
                    });

                let mut keys = HashSet::<Value>::new();
                let _ = keys.try_reserve(distinct_values.len());
                for v in distinct_values {
                    if v.is_blank() {
                        continue;
                    }
                    keys.insert(v);
                }

                // `has_duplicates` is used to decide whether relationship traversal/grouping needs
                // a many-to-many expansion algorithm. Physical BLANK keys never participate in
                // joins, so compute duplication only across non-blank keys.
                let non_blank_rows = to_table
                    .stats_non_blank_count(to_idx)
                    .unwrap_or_else(|| to_table.row_count());
                let has_duplicates = keys.len() < non_blank_rows;
                ToIndex::KeySet {
                    keys,
                    has_duplicates,
                }
            }
            _ => {
                let mut map = HashMap::<Value, RowSet>::new();
                let mut has_duplicates = false;
                for row in 0..to_table.row_count() {
                    let value = to_table.value_by_idx(row, to_idx).unwrap_or(Value::Blank);
                    match relationship.cardinality {
                        Cardinality::OneToMany | Cardinality::OneToOne => {
                            if map.insert(value.clone(), RowSet::One(row)).is_some() {
                                return Err(DaxError::NonUniqueKey {
                                    table: relationship.to_table.clone(),
                                    column: relationship.to_column.clone(),
                                    value: value.clone(),
                                });
                            }
                        }
                        Cardinality::ManyToMany => {
                            // Physical BLANK keys do not participate in relationship joins (fact-side BLANK
                            // foreign keys map to the relationship-generated virtual blank member). Skip
                            // materializing row lists for BLANK to avoid a potentially huge, never-used
                            // vector when the dimension contains many BLANK keys.
                            if value.is_blank() {
                                continue;
                            }

                            match map.entry(value) {
                                Entry::Vacant(v) => {
                                    v.insert(RowSet::One(row));
                                }
                                Entry::Occupied(mut o) => {
                                    has_duplicates = true;
                                    o.get_mut().push(row);
                                }
                            }
                        }
                    }
                }
                ToIndex::RowSets {
                    map,
                    has_duplicates,
                }
            }
        };

        let (from_index, unmatched_fact_rows) = match &from_table.storage {
            TableStorage::InMemory(_) => {
                let mut from_index: HashMap<Value, Vec<usize>> = HashMap::new();
                let mut unmatched = UnmatchedFactRowsBuilder::new(from_table.row_count());
                for row in 0..from_table.row_count() {
                    let value = from_table
                        .value_by_idx(row, from_idx)
                        .unwrap_or(Value::Blank);
                    let rows = from_index.entry(value.clone()).or_default();

                    // For 1:1 relationships, the "from" side must also be unique. We treat
                    // BLANK as a real key for uniqueness, matching the existing to-side semantics.
                    if relationship.cardinality == Cardinality::OneToOne && !rows.is_empty() {
                        return Err(DaxError::NonUniqueKey {
                            table: relationship.from_table.clone(),
                            column: from_col.clone(),
                            value: value.clone(),
                        });
                    }

                    rows.push(row);

                    let matched = to_index.contains_key(&value);
                    if value.is_blank() || !matched {
                        unmatched.push(row);
                    }

                    if relationship.enforce_referential_integrity && !value.is_blank() && !matched {
                        return Err(DaxError::ReferentialIntegrityViolation {
                            relationship: relationship.name.clone(),
                            from_table: relationship.from_table.clone(),
                            from_column: from_col.clone(),
                            to_table: relationship.to_table.clone(),
                            to_column: to_col.clone(),
                            value: value.clone(),
                        });
                    }
                }

                let unmatched = unmatched.finish();
                let unmatched_fact_rows = (!unmatched.is_empty()).then_some(unmatched);
                (Some(from_index), unmatched_fact_rows)
            }
            TableStorage::Columnar(_) => {
                // Avoid materializing `from_index` for columnar fact tables. Instead, precompute
                // the set of fact rows that belong to the virtual blank dimension member.
                let mut unmatched = UnmatchedFactRowsBuilder::new(from_table.row_count());
                let mut seen = (relationship.cardinality == Cardinality::OneToOne)
                    .then_some(HashSet::<Value>::new());
                for row in 0..from_table.row_count() {
                    let value = from_table
                        .value_by_idx(row, from_idx)
                        .unwrap_or(Value::Blank);

                    if let Some(seen) = seen.as_mut() {
                        if !seen.insert(value.clone()) {
                            return Err(DaxError::NonUniqueKey {
                                table: relationship.from_table.clone(),
                                column: from_col.clone(),
                                value: value.clone(),
                            });
                        }
                    }

                    let matched = to_index.contains_key(&value);

                    if value.is_blank() || !matched {
                        unmatched.push(row);
                    }

                    if relationship.enforce_referential_integrity && !value.is_blank() && !matched {
                        return Err(DaxError::ReferentialIntegrityViolation {
                            relationship: relationship.name.clone(),
                            from_table: relationship.from_table.clone(),
                            from_column: from_col.clone(),
                            to_table: relationship.to_table.clone(),
                            to_column: to_col.clone(),
                            value: value.clone(),
                        });
                    }
                }

                (None, Some(unmatched.finish()))
            }
        };

        self.relationships.push(RelationshipInfo {
            rel: relationship,
            from_table_key,
            from_column_key,
            to_table_key,
            to_column_key,
            from_idx,
            to_idx,
            to_index,
            from_index,
            unmatched_fact_rows,
        });
        Ok(())
    }

    fn validate_relationship_join_column_types(
        relationship: &Relationship,
        from_table: &Table,
        from_idx: usize,
        to_table: &Table,
        to_idx: usize,
    ) -> DaxResult<()> {
        #[derive(Clone, Copy, Debug, PartialEq, Eq)]
        enum JoinType {
            Numeric,
            Text,
            Boolean,
        }

        struct JoinTypeInfo {
            kind: JoinType,
            display: String,
        }

        const SCAN_ROWS: usize = 1_000;

        // Columnar tables have a declared type per column. Relationships with incompatible join
        // column types almost always lead to "no matches" during filter propagation, which is
        // extremely confusing to debug. Fail fast when the types are clearly incompatible.
        //
        // Compatibility rules:
        // - `Number`, `DateTime`, `Currency`, and `Percentage` are treated as "numeric-like" and
        //   considered compatible for relationship joins. The DAX engine coerces these logical
        //   types to `Value::Number` internally (see `ColumnarTableBackend::dax_from_columnar`).
        // - `String` and `Boolean` must match exactly with their respective kinds.
        fn join_type_from_columnar(column_type: formula_columnar::ColumnType) -> JoinTypeInfo {
            let kind = match column_type {
                formula_columnar::ColumnType::Number
                | formula_columnar::ColumnType::DateTime
                | formula_columnar::ColumnType::Currency { .. }
                | formula_columnar::ColumnType::Percentage { .. } => JoinType::Numeric,
                formula_columnar::ColumnType::String => JoinType::Text,
                formula_columnar::ColumnType::Boolean => JoinType::Boolean,
            };

            let display = match column_type {
                formula_columnar::ColumnType::Number => "Number".to_string(),
                formula_columnar::ColumnType::String => "String".to_string(),
                formula_columnar::ColumnType::Boolean => "Boolean".to_string(),
                formula_columnar::ColumnType::DateTime => "DateTime".to_string(),
                formula_columnar::ColumnType::Currency { scale } => {
                    format!("Currency(scale={scale})")
                }
                formula_columnar::ColumnType::Percentage { scale } => {
                    format!("Percentage(scale={scale})")
                }
            };

            JoinTypeInfo { kind, display }
        }

        fn join_type_from_in_memory_values(table: &Table, idx: usize) -> Option<JoinTypeInfo> {
            let row_count = table.row_count();
            let scan = row_count.min(SCAN_ROWS);
            for row in 0..scan {
                let value = table.value_by_idx(row, idx).unwrap_or(Value::Blank);
                if value.is_blank() {
                    continue;
                }

                let (kind, display) = match value {
                    Value::Number(_) => (JoinType::Numeric, "Number".to_string()),
                    Value::Text(_) => (JoinType::Text, "Text".to_string()),
                    Value::Boolean(_) => (JoinType::Boolean, "Boolean".to_string()),
                    Value::Blank => continue,
                };
                return Some(JoinTypeInfo { kind, display });
            }
            None
        }

        fn join_type_for_table_column(table: &Table, idx: usize) -> Option<JoinTypeInfo> {
            if let Some(col_table) = table.columnar_table() {
                let column_type = col_table.schema().get(idx)?.column_type;
                return Some(join_type_from_columnar(column_type));
            }
            join_type_from_in_memory_values(table, idx)
        }

        let from_type = join_type_for_table_column(from_table, from_idx);
        let to_type = join_type_for_table_column(to_table, to_idx);

        // If we can't infer a type for one side (e.g. all BLANKs in the scan window), skip
        // validation. This avoids false positives when loading sparse/empty in-memory tables.
        let (Some(from_type), Some(to_type)) = (from_type, to_type) else {
            return Ok(());
        };

        if from_type.kind != to_type.kind {
            return Err(DaxError::RelationshipJoinColumnTypeMismatch {
                relationship: relationship.name.clone(),
                from_table: relationship.from_table.clone(),
                from_column: relationship.from_column.clone(),
                from_type: from_type.display,
                to_table: relationship.to_table.clone(),
                to_column: relationship.to_column.clone(),
                to_type: to_type.display,
            });
        }

        Ok(())
    }

    pub fn add_measure(
        &mut self,
        name: impl Into<String>,
        expression: impl Into<String>,
    ) -> DaxResult<()> {
        let name = name.into();
        let display_name = Self::normalize_measure_name(&name).to_string();
        let key = normalize_ident(&display_name);
        if self.measures.contains_key(&key) {
            return Err(DaxError::DuplicateMeasure {
                measure: display_name,
            });
        }
        let expression = expression.into();
        let parsed = crate::parser::parse(&expression)?;
        self.measures.insert(
            key,
            Measure {
                name: display_name,
                expression,
                parsed,
            },
        );
        Ok(())
    }

    pub fn add_calculated_column(
        &mut self,
        table: impl Into<String>,
        name: impl Into<String>,
        expression: impl Into<String>,
    ) -> DaxResult<()> {
        let table = table.into();
        let table_key = normalize_ident(&table);
        let name = name.into();
        let expression = expression.into();

        let parsed = crate::parser::parse(&expression)?;

        enum NewColumn {
            InMemory(Vec<Value>),
            Columnar(formula_columnar::EncodedColumn),
        }

        let new_column = {
            let Some(table_ref) = self.tables.get(&table_key) else {
                return Err(DaxError::UnknownTable(table.clone()));
            };

            // Power Pivot stores calculated column values physically in the table. We mirror that
            // behavior by evaluating the expression eagerly for every existing row and then
            // materializing the resulting values into the table backend (including columnar
            // tables).
            //
            // Note: columnar-backed tables require a single physical column type; expressions that
            // produce mixed value types across rows will currently return a type error.
            match &table_ref.storage {
                TableStorage::InMemory(_) => {
                    let mut results = Vec::new();
                    let _ = results.try_reserve_exact(table_ref.row_count());
                    let engine = crate::engine::DaxEngine::new();
                    let filter_ctx = FilterContext::default();
                    let mut row_ctx = RowContext::default();
                    row_ctx.push(table_ref.name(), 0);
                    for row in 0..table_ref.row_count() {
                        row_ctx.set_current_row(row);
                        let value = engine.evaluate_expr(self, &parsed, &filter_ctx, &row_ctx)?;
                        results.push(value);
                    }
                    NewColumn::InMemory(results)
                }
                TableStorage::Columnar(backend) => {
                    use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder};

                    if table_ref.column_idx(&name).is_some() {
                        return Err(DaxError::DuplicateColumn {
                            table: table.clone(),
                            column: name.clone(),
                        });
                    }

                    let options = backend.table.options();
                    let row_count = table_ref.row_count();
                    let engine = crate::engine::DaxEngine::new();
                    let filter_ctx = FilterContext::default();
                    let mut row_ctx = RowContext::default();
                    row_ctx.push(table_ref.name(), 0);

                    let mut leading_nulls: usize = 0;
                    let mut inferred_type: Option<ColumnType> = None;
                    let mut builder: Option<ColumnarTableBuilder> = None;

                    for row in 0..row_count {
                        row_ctx.set_current_row(row);
                        let value = engine.evaluate_expr(self, &parsed, &filter_ctx, &row_ctx)?;

                        match (inferred_type, &value) {
                            (None, Value::Blank) => {
                                leading_nulls += 1;
                                continue;
                            }
                            (None, Value::Number(_))
                            | (None, Value::Text(_))
                            | (None, Value::Boolean(_)) => {
                                let ty = match &value {
                                    Value::Number(_) => ColumnType::Number,
                                    Value::Text(_) => ColumnType::String,
                                    Value::Boolean(_) => ColumnType::Boolean,
                                    Value::Blank => {
                                        debug_assert!(false, "blank value reached type inference");
                                        return Err(DaxError::Eval(
                                            "calculated column type inference failed".into(),
                                        ));
                                    }
                                };

                                let schema = vec![ColumnSchema {
                                    name: name.clone(),
                                    column_type: ty,
                                }];
                                let mut b = ColumnarTableBuilder::new(schema, options);
                                for _ in 0..leading_nulls {
                                    b.append_value(formula_columnar::Value::Null);
                                }

                                let encoded = match &value {
                                    Value::Number(n) => formula_columnar::Value::Number(n.0),
                                    Value::Text(s) => formula_columnar::Value::String(s.clone()),
                                    Value::Boolean(bv) => formula_columnar::Value::Boolean(*bv),
                                    Value::Blank => formula_columnar::Value::Null,
                                };
                                b.append_value(encoded);

                                inferred_type = Some(ty);
                                builder = Some(b);
                            }
                            (Some(_), Value::Blank) => {
                                if let Some(b) = builder.as_mut() {
                                    b.append_value(formula_columnar::Value::Null);
                                }
                            }
                            (Some(ty), v) => {
                                let matches = match (ty, v) {
                                    (ColumnType::Number, Value::Number(_)) => true,
                                    (ColumnType::String, Value::Text(_)) => true,
                                    (ColumnType::Boolean, Value::Boolean(_)) => true,
                                    _ => false,
                                };
                                if !matches {
                                    let expected = match ty {
                                        ColumnType::Number => "Number",
                                        ColumnType::String => "Text",
                                        ColumnType::Boolean => "Boolean",
                                        ColumnType::DateTime => "DateTime",
                                        ColumnType::Currency { .. } => "Currency",
                                        ColumnType::Percentage { .. } => "Percentage",
                                    };
                                    let actual = match v {
                                        Value::Blank => "Blank",
                                        Value::Number(_) => "Number",
                                        Value::Text(_) => "Text",
                                        Value::Boolean(_) => "Boolean",
                                    };
                                    return Err(DaxError::Type(format!(
                                        "calculated column {table}[{name}] produced {actual} after inferring {expected}"
                                    )));
                                }

                                let encoded = match v {
                                    Value::Number(n) => formula_columnar::Value::Number(n.0),
                                    Value::Text(s) => formula_columnar::Value::String(s.clone()),
                                    Value::Boolean(bv) => formula_columnar::Value::Boolean(*bv),
                                    Value::Blank => formula_columnar::Value::Null,
                                };
                                if let Some(b) = builder.as_mut() {
                                    b.append_value(encoded);
                                }
                            }
                        }
                    }

                    // If the column is entirely blank, choose a deterministic default type (Number)
                    // and encode the entire column as nulls.
                    let b = match builder {
                        Some(b) => b,
                        None => {
                            let schema = vec![ColumnSchema {
                                name: name.clone(),
                                column_type: ColumnType::Number,
                            }];
                            let mut b = ColumnarTableBuilder::new(schema, options);
                            for _ in 0..row_count {
                                b.append_value(formula_columnar::Value::Null);
                            }
                            b
                        }
                    };

                    let mut encoded_cols = b.finalize().into_encoded_columns();
                    let encoded = encoded_cols.pop().ok_or_else(|| {
                        DaxError::Eval("expected one encoded column from builder".into())
                    })?;
                    if !encoded_cols.is_empty() {
                        return Err(DaxError::Eval(
                            "expected one encoded column from builder".into(),
                        ));
                    }

                    NewColumn::Columnar(encoded)
                }
            }
        };

        let calc = CalculatedColumn {
            table: self
                .tables
                .get(&table_key)
                .map(|t| t.name.clone())
                .unwrap_or_else(|| table.clone()),
            name: name.clone(),
            expression,
            parsed: parsed.clone(),
        };

        let table_mut = self
            .tables
            .get_mut(&table_key)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

        match (&mut table_mut.storage, new_column) {
            (TableStorage::InMemory(_), NewColumn::InMemory(values)) => {
                table_mut.add_column(name.clone(), values)?;
            }
            (TableStorage::Columnar(backend), NewColumn::Columnar(encoded)) => {
                let base = backend.table.as_ref().clone();
                let appended = base
                    .with_appended_encoded_column(encoded)
                    .map_err(|e| match e {
                        formula_columnar::ColumnAppendError::DuplicateColumn { name: col } => {
                            DaxError::DuplicateColumn {
                                table: table.clone(),
                                column: col,
                            }
                        }
                        formula_columnar::ColumnAppendError::LengthMismatch {
                            expected,
                            actual,
                        } => DaxError::ColumnLengthMismatch {
                            table: table.clone(),
                            column: name.clone(),
                            expected,
                            actual,
                        },
                        other => DaxError::Eval(format!(
                            "failed to append encoded calculated column {table}[{name}]: {other}"
                        )),
                    })?;

                backend.table = Arc::new(appended);

                backend.columns = backend
                    .table
                    .schema()
                    .iter()
                    .map(|c| c.name.clone())
                    .collect();
                backend.column_index = backend
                    .columns
                    .iter()
                    .enumerate()
                    .map(|(idx, c)| (normalize_ident(c), idx))
                    .collect();
            }
            (TableStorage::InMemory(_), NewColumn::Columnar(_))
            | (TableStorage::Columnar(_), NewColumn::InMemory(_)) => {
                return Err(DaxError::Eval("calculated column backend mismatch".into()));
            }
        }

        self.calculated_columns.push(calc);
        // Recompute the evaluation order for calculated columns in this table so `insert_row`
        // can honor intra-table dependencies (Power Pivot allows calculated columns to reference
        // other calculated columns in the same table).
        if let Err(err) = self.refresh_calculated_column_order(&table_key) {
            // Roll back the definition. The physical column was already added; for in-memory
            // tables we can also remove the last column, but columnar tables do not currently
            // support removing columns.
            self.calculated_columns.pop();
            if let Some(table_mut) = self.tables.get_mut(&table_key) {
                let _ = table_mut.pop_last_column();
            }
            return Err(err);
        }
        Ok(())
    }

    /// Register a calculated column definition for a table that already contains the computed
    /// values.
    ///
    /// This is useful when loading persisted models: Power Pivot stores calculated column values
    /// physically in the table, so re-evaluating them on load is both expensive and can fail if
    /// the backend is immutable (e.g. [`Table::from_columnar`]).
    pub fn add_calculated_column_definition(
        &mut self,
        table: impl Into<String>,
        name: impl Into<String>,
        expression: impl Into<String>,
    ) -> DaxResult<()> {
        let table = table.into();
        let table_key = normalize_ident(&table);
        let name = name.into();
        let name_key = normalize_ident(&name);
        if self
            .calculated_columns
            .iter()
            .any(|c| normalize_ident(&c.table) == table_key && normalize_ident(&c.name) == name_key)
        {
            return Err(DaxError::DuplicateColumn {
                table,
                column: name,
            });
        }

        let table_ref = self
            .tables
            .get(&table_key)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        let Some(col_idx) = table_ref.column_idx(&name) else {
            return Err(DaxError::UnknownColumn {
                table,
                column: name,
            });
        };

        let expression = expression.into();
        let parsed = crate::parser::parse(&expression)?;
        self.calculated_columns.push(CalculatedColumn {
            table: table_ref.name.clone(),
            name: table_ref.columns().get(col_idx).cloned().unwrap_or(name),
            expression,
            parsed,
        });
        if let Err(err) = self.refresh_calculated_column_order(&table_key) {
            self.calculated_columns.pop();
            return Err(err);
        }
        Ok(())
    }

    pub fn evaluate_measure(&self, name: &str, filter: &FilterContext) -> DaxResult<Value> {
        let key = normalize_ident(Self::normalize_measure_name(name));
        let measure = self
            .measures
            .get(&key)
            .ok_or_else(|| DaxError::UnknownMeasure(name.to_string()))?;
        crate::engine::DaxEngine::new().evaluate_expr(
            self,
            &measure.parsed,
            filter,
            &RowContext::default(),
        )
    }

    pub(crate) fn measures(&self) -> &HashMap<String, Measure> {
        &self.measures
    }

    pub(crate) fn relationships(&self) -> &[RelationshipInfo] {
        &self.relationships
    }

    pub(crate) fn find_unique_active_relationship_path<F>(
        &self,
        from_table: &str,
        to_table: &str,
        direction: RelationshipPathDirection,
        is_relationship_active: F,
    ) -> DaxResult<Option<Vec<usize>>>
    where
        F: Fn(usize, &RelationshipInfo) -> bool,
    {
        let from_key = normalize_ident(from_table);
        let to_key = normalize_ident(to_table);
        let from_display = self
            .tables
            .get(&from_key)
            .map(|t| t.name().to_string())
            .unwrap_or_else(|| from_table.trim().to_string());
        let to_display = self
            .tables
            .get(&to_key)
            .map(|t| t.name().to_string())
            .unwrap_or_else(|| to_table.trim().to_string());

        // We intentionally do not treat `from_table == to_table` as a valid path here.
        // Callers like `RELATED`/`RELATEDTABLE` are defined in terms of relationships, and
        // previously errored when the target table was the current table.
        if from_key == to_key {
            return Ok(None);
        }

        fn dfs<F>(
            model: &DataModel,
            start_table_display: &str,
            current_table_key: &str,
            target_table_key: &str,
            target_table_display: &str,
            direction: RelationshipPathDirection,
            is_relationship_active: &F,
            visited: &mut HashSet<String>,
            path: &mut Vec<usize>,
            table_path_display: &mut Vec<String>,
            found_path: &mut Option<Vec<usize>>,
            found_table_path_display: &mut Option<Vec<String>>,
        ) -> DaxResult<()>
        where
            F: Fn(usize, &RelationshipInfo) -> bool,
        {
            if current_table_key == target_table_key {
                if found_path.is_some() {
                    let first = found_table_path_display
                        .as_ref()
                        .map(|p| p.join(" -> "))
                        .unwrap_or_else(|| "<unknown>".to_string());
                    let second = table_path_display.join(" -> ");
                    return Err(DaxError::Eval(format!(
                        "ambiguous active relationship path between {start_table_display} and {target_table_display}: {first}; {second}"
                    )));
                }
                *found_path = Some(path.clone());
                *found_table_path_display = Some(table_path_display.clone());
                return Ok(());
            }

            for (idx, rel) in model.relationships.iter().enumerate() {
                if !(is_relationship_active(idx, rel)) {
                    continue;
                }

                let next_table = match direction {
                    RelationshipPathDirection::ManyToOne => {
                        if rel.from_table_key != current_table_key {
                            continue;
                        }
                        rel.to_table_key.as_str()
                    }
                    RelationshipPathDirection::OneToMany => {
                        if rel.to_table_key != current_table_key {
                            continue;
                        }
                        rel.from_table_key.as_str()
                    }
                };

                if visited.contains(next_table) {
                    continue;
                }

                visited.insert(next_table.to_string());
                path.push(idx);
                let next_display = model
                    .tables
                    .get(next_table)
                    .map(|t| t.name().to_string())
                    .unwrap_or_else(|| next_table.to_string());
                table_path_display.push(next_display);

                dfs(
                    model,
                    start_table_display,
                    next_table,
                    target_table_key,
                    target_table_display,
                    direction,
                    is_relationship_active,
                    visited,
                    path,
                    table_path_display,
                    found_path,
                    found_table_path_display,
                )?;

                table_path_display.pop();
                path.pop();
                visited.remove(next_table);
            }

            Ok(())
        }

        let mut visited = HashSet::new();
        visited.insert(from_key.clone());
        let mut path = Vec::new();
        let mut table_path_display = vec![from_display.clone()];
        let mut found_path = None;
        let mut found_table_path_display = None;

        dfs(
            self,
            &from_display,
            &from_key,
            &to_key,
            &to_display,
            direction,
            &is_relationship_active,
            &mut visited,
            &mut path,
            &mut table_path_display,
            &mut found_path,
            &mut found_table_path_display,
        )?;

        Ok(found_path)
    }

    pub(crate) fn find_relationship_index(
        &self,
        table_a: &str,
        column_a: &str,
        table_b: &str,
        column_b: &str,
    ) -> DaxResult<Option<usize>> {
        let table_a_display = table_a.trim();
        let column_a_display = column_a.trim();
        let table_b_display = table_b.trim();
        let column_b_display = column_b.trim();

        let table_a = normalize_ident(table_a_display);
        let column_a = normalize_ident(column_a_display);
        let table_b = normalize_ident(table_b_display);
        let column_b = normalize_ident(column_b_display);

        let mut matches = Vec::new();
        for (idx, info) in self.relationships.iter().enumerate() {
            let forward = info.from_table_key == table_a
                && info.from_column_key == column_a
                && info.to_table_key == table_b
                && info.to_column_key == column_b;
            let reverse = info.from_table_key == table_b
                && info.from_column_key == column_b
                && info.to_table_key == table_a
                && info.to_column_key == column_a;
            if forward || reverse {
                matches.push(idx);
            }
        }

        match matches.len() {
            0 => Ok(None),
            1 => Ok(Some(matches[0])),
            _ => {
                let mut relationship_names: Vec<String> = matches
                    .iter()
                    .filter_map(|&idx| self.relationships.get(idx).map(|rel| rel.rel.name.clone()))
                    .collect();
                relationship_names.sort();
                relationship_names.dedup();
                Err(DaxError::Eval(format!(
                    "multiple relationships found between {table_a_display}[{column_a_display}] and {table_b_display}[{column_b_display}]: {}",
                    relationship_names.join(", ")
                )))
            }
        }
    }

    pub(crate) fn normalize_measure_name(name: &str) -> &str {
        name.strip_prefix('[')
            .and_then(|n| n.strip_suffix(']'))
            .unwrap_or(name)
            .trim()
    }

    fn refresh_calculated_column_order(&mut self, table: &str) -> DaxResult<()> {
        let order = self.build_calculated_column_order(table)?;
        self.calculated_column_order
            .insert(table.to_string(), order);
        Ok(())
    }

    fn build_calculated_column_order(&self, table: &str) -> DaxResult<Vec<usize>> {
        let table_key = normalize_ident(table);
        let table_display = self
            .tables
            .get(&table_key)
            .map(|t| t.name().to_string())
            .unwrap_or_else(|| table.trim().to_string());
        let calc_indices: Vec<usize> = self
            .calculated_columns
            .iter()
            .enumerate()
            .filter_map(|(idx, c)| (normalize_ident(&c.table) == table_key).then_some(idx))
            .collect();

        if calc_indices.is_empty() {
            return Ok(Vec::new());
        }

        // Map calculated column names to their global index in `self.calculated_columns`.
        let mut name_to_idx: HashMap<String, usize> = HashMap::new();
        for &idx in &calc_indices {
            if let Some(calc) = self.calculated_columns.get(idx) {
                name_to_idx.insert(normalize_ident(&calc.name), idx);
            }
        }

        // Build dependency edges between calculated columns within the same table.
        let mut deps_by_calc: HashMap<usize, Vec<usize>> = HashMap::new();
        for &idx in &calc_indices {
            let Some(calc) = self.calculated_columns.get(idx) else {
                continue;
            };
            let deps = self.collect_same_table_column_dependencies(&calc.parsed, &table_key);
            let mut calc_deps: Vec<usize> = deps
                .into_iter()
                .filter_map(|dep_name| name_to_idx.get(&dep_name).copied())
                .collect();
            // Keep ordering deterministic.
            calc_deps.sort_unstable();
            calc_deps.dedup();
            deps_by_calc.insert(idx, calc_deps);
        }

        #[derive(Clone, Copy, Debug, PartialEq, Eq)]
        enum VisitState {
            Visiting,
            Visited,
        }

        let mut state: HashMap<usize, VisitState> = HashMap::new();
        let mut stack: Vec<usize> = Vec::new();
        let mut out: Vec<usize> = Vec::new();
        let _ = out.try_reserve_exact(calc_indices.len());

        let visit = |start: usize,
                     state: &mut HashMap<usize, VisitState>,
                     stack: &mut Vec<usize>,
                     out: &mut Vec<usize>,
                     deps_by_calc: &HashMap<usize, Vec<usize>>,
                     this: &DataModel|
         -> DaxResult<()> {
            fn dfs(
                node: usize,
                state: &mut HashMap<usize, VisitState>,
                stack: &mut Vec<usize>,
                out: &mut Vec<usize>,
                deps_by_calc: &HashMap<usize, Vec<usize>>,
                table: &str,
                model: &DataModel,
            ) -> DaxResult<()> {
                match state.get(&node) {
                    Some(VisitState::Visited) => return Ok(()),
                    Some(VisitState::Visiting) => {
                        // Should only happen when we re-enter via an edge; handled by caller.
                        return Ok(());
                    }
                    None => {}
                }

                state.insert(node, VisitState::Visiting);
                stack.push(node);

                if let Some(deps) = deps_by_calc.get(&node) {
                    for &dep in deps {
                        if matches!(state.get(&dep), Some(VisitState::Visiting)) {
                            let start_pos = stack.iter().position(|&n| n == dep).unwrap_or(0);
                            let mut cycle_nodes: Vec<usize> =
                                stack[start_pos..].iter().copied().collect();
                            cycle_nodes.push(dep);
                            let cycle_names: Vec<String> = cycle_nodes
                                .iter()
                                .filter_map(|idx| model.calculated_columns.get(*idx))
                                .map(|c| c.name.clone())
                                .collect();
                            return Err(DaxError::Eval(format!(
                                "calculated column dependency cycle in {table}: {}",
                                cycle_names.join(" -> ")
                            )));
                        }
                        dfs(dep, state, stack, out, deps_by_calc, table, model)?;
                    }
                }

                stack.pop();
                state.insert(node, VisitState::Visited);
                out.push(node);
                Ok(())
            }

            dfs(
                start,
                state,
                stack,
                out,
                deps_by_calc,
                table_display.as_str(),
                this,
            )
        };

        // Use definition order as a stable traversal order.
        for &idx in &calc_indices {
            if matches!(state.get(&idx), Some(VisitState::Visited)) {
                continue;
            }
            visit(idx, &mut state, &mut stack, &mut out, &deps_by_calc, self)?;
        }

        Ok(out)
    }

    fn collect_same_table_column_dependencies(
        &self,
        expr: &Expr,
        current_table: &str,
    ) -> HashSet<String> {
        let mut out = HashSet::new();
        self.collect_same_table_column_dependencies_inner(expr, current_table, &mut out);
        out
    }

    #[deny(unreachable_patterns)]
    fn collect_same_table_column_dependencies_inner(
        &self,
        expr: &Expr,
        current_table: &str,
        out: &mut HashSet<String>,
    ) {
        match expr {
            Expr::Let { bindings, body } => {
                for (_, binding_expr) in bindings {
                    self.collect_same_table_column_dependencies_inner(
                        binding_expr,
                        current_table,
                        out,
                    );
                }
                self.collect_same_table_column_dependencies_inner(body, current_table, out);
            }
            Expr::Tuple(values) => {
                for value in values {
                    self.collect_same_table_column_dependencies_inner(value, current_table, out);
                }
            }
            Expr::ColumnRef { table, column } => {
                if normalize_ident(table) == current_table {
                    out.insert(normalize_ident(column));
                }
            }
            Expr::Measure(name) => {
                let normalized = Self::normalize_measure_name(name);
                // In a calculated column (row context), `[Name]` can resolve to either a measure
                // or a same-table column. Only treat it as a column dependency when no measure
                // exists and the table contains a column with that name.
                let measure_key = normalize_ident(normalized);
                if !self.measures.contains_key(&measure_key) {
                    if let Some(table_ref) = self.tables.get(current_table) {
                        if table_ref.column_idx(normalized).is_some() {
                            out.insert(normalize_ident(normalized));
                        }
                    }
                }
            }
            Expr::Call { args, .. } => {
                for arg in args {
                    self.collect_same_table_column_dependencies_inner(arg, current_table, out);
                }
            }
            Expr::UnaryOp { expr, .. } => {
                self.collect_same_table_column_dependencies_inner(expr, current_table, out);
            }
            Expr::BinaryOp { left, right, .. } => {
                self.collect_same_table_column_dependencies_inner(left, current_table, out);
                self.collect_same_table_column_dependencies_inner(right, current_table, out);
            }
            Expr::TableLiteral { rows } => {
                for row in rows {
                    for cell in row {
                        self.collect_same_table_column_dependencies_inner(cell, current_table, out);
                    }
                }
            }
            Expr::Number(_) | Expr::Text(_) | Expr::Boolean(_) | Expr::TableName(_) => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::FilterContext;
    use formula_columnar::{
        ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
    };
    use std::time::Instant;

    #[test]
    fn unmatched_fact_rows_builder_switches_to_dense_when_more_efficient() {
        // Threshold is `row_count / 64`, and we switch when `unmatched_count > threshold`.
        // For 128 rows, this becomes `> 2`.
        let mut builder = UnmatchedFactRowsBuilder::new(128);
        builder.push(0);
        builder.push(1);
        assert!(matches!(builder.rows, UnmatchedFactRows::Sparse(_)));

        builder.push(2);
        match &builder.rows {
            UnmatchedFactRows::Dense { len, count, bits } => {
                assert_eq!(*len, 128);
                assert_eq!(*count, 3);
                assert_eq!(bits.len(), 2); // 128 rows => 2 u64 words.
            }
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }

        let dense = builder.finish();
        let mut rows = Vec::new();
        dense.extend_into(&mut rows);
        rows.sort_unstable();
        assert_eq!(rows, vec![0, 1, 2]);

        let mut allowed = BitVec::with_len_all_false(128);
        allowed.set(1, true);
        assert!(dense.any_row_allowed(&allowed));
    }

    #[test]
    fn unmatched_fact_rows_builder_dense_does_not_double_count_duplicates() {
        // Force the builder into the dense representation quickly.
        //
        // For 64 rows, the threshold is `row_count / 64 == 1`, and we switch when
        // `unmatched_count > threshold`, i.e. on the 2nd push.
        let mut builder = UnmatchedFactRowsBuilder::new(64);
        builder.push(0);
        builder.push(1);

        match &builder.rows {
            UnmatchedFactRows::Dense { count, .. } => assert_eq!(*count, 2),
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }

        // Pushing a duplicate row should not change `count` once in the dense representation.
        builder.push(1);
        match &builder.rows {
            UnmatchedFactRows::Dense { count, .. } => assert_eq!(*count, 2),
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }

        let dense = builder.finish();
        let mut rows = Vec::new();
        dense.extend_into(&mut rows);
        rows.sort_unstable();
        assert_eq!(rows, vec![0, 1]);
    }

    #[test]
    fn unmatched_fact_rows_builder_sparse_to_dense_dedups_duplicates() {
        // When converting from the sparse vec to the dense bitmap representation, we should
        // compute `count` based on the number of unique bits set, not the length of the vec.
        let mut builder = UnmatchedFactRowsBuilder::new(64);
        builder.push(0);
        assert!(matches!(builder.rows, UnmatchedFactRows::Sparse(_)));

        // Duplicate push triggers conversion (threshold is 1, and we switch when len > 1).
        builder.push(0);

        match &builder.rows {
            UnmatchedFactRows::Dense { count, .. } => assert_eq!(*count, 1),
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }

        let dense = builder.finish();
        let mut rows = Vec::new();
        dense.extend_into(&mut rows);
        rows.sort_unstable();
        assert_eq!(rows, vec![0]);
    }

    #[test]
    fn unmatched_fact_rows_retain_updates_count_and_bits() {
        let mut sparse = UnmatchedFactRows::Sparse(vec![0, 1, 2, 3, 4]);
        sparse.retain(|row| *row % 2 == 0);
        match sparse {
            UnmatchedFactRows::Sparse(rows) => assert_eq!(rows, vec![0, 2, 4]),
            UnmatchedFactRows::Dense { .. } => panic!("expected sparse representation"),
        }

        // Use a length that is not a multiple of 64 so the last word has unused bits. We
        // intentionally set one such out-of-range bit to ensure `retain` clears it and keeps
        // `count` accurate.
        let len = 130;
        let mut bits = vec![0u64; (len + 63) / 64];
        // In-range rows.
        bits[0] |= 1u64 << 0; // row 0
        bits[0] |= 1u64 << 1; // row 1
        bits[1] |= 1u64 << 0; // row 64
        bits[1] |= 1u64 << 1; // row 65
        bits[2] |= 1u64 << 1; // row 129
                              // Out-of-range row (191 >= 130) in the last word.
        bits[2] |= 1u64 << 63;

        let mut dense = UnmatchedFactRows::Dense {
            bits,
            len,
            count: 6,
        };
        dense.retain(|row| *row % 2 == 1);

        let mut rows = Vec::new();
        dense.extend_into(&mut rows);
        rows.sort_unstable();
        assert_eq!(rows, vec![1, 65, 129]);

        match dense {
            UnmatchedFactRows::Dense { bits, len, count } => {
                assert_eq!(len, 130);
                assert_eq!(count, 3);
                assert_eq!(bits[2] & (1u64 << 63), 0);
            }
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }
    }

    #[test]
    fn unmatched_fact_rows_retain_dense_does_not_read_past_len() {
        // Ensure the Dense retain implementation never evaluates the predicate for rows >= len,
        // even if unused bits in the last word are (incorrectly) set.
        let len = 70usize;
        let word_len = (len + 63) / 64;
        assert_eq!(word_len, 2);

        let mut bits = vec![0u64; word_len];
        bits[0] |= 1u64 << 0; // row 0 (in-range)
        bits[1] |= 1u64 << 5; // row 69 (in-range)
        bits[1] |= 1u64 << 6; // row 70 (out-of-range for len=70)
        assert_ne!(bits[1] & (1u64 << 6), 0);

        // Count tracks bits within len only.
        let mut rows = UnmatchedFactRows::Dense {
            bits,
            len,
            count: 2,
        };

        rows.retain(|row| {
            assert!(
                *row < len,
                "retain predicate must not be evaluated for rows >= len"
            );
            true
        });

        match &rows {
            UnmatchedFactRows::Dense { bits, len, count } => {
                assert_eq!(*len, 70);
                assert_eq!(*count, 2);
                assert_eq!(bits.len(), 2);
                assert_eq!(bits[0], 1u64);
                // Only row 69 remains set in the second word (bit 5); out-of-range bit 6 cleared.
                assert_eq!(bits[1], 1u64 << 5);
            }
            UnmatchedFactRows::Sparse(_) => panic!("expected dense representation"),
        }

        let mut out = Vec::new();
        rows.extend_into(&mut out);
        out.sort_unstable();
        assert_eq!(out, vec![0, 69]);
    }

    #[test]
    fn relationship_large_columnar_does_not_explode_memory() {
        if std::env::var_os("FORMULA_DAX_REL_BENCH").is_none() {
            return;
        }

        let mut model = DataModel::new();

        // Small in-memory dimension table.
        let mut dim = Table::new("Dim", vec!["Id"]);
        for i in 0..100i64 {
            dim.push_row(vec![i.into()]).unwrap();
        }
        model.add_table(dim).unwrap();

        // Large columnar fact table with high-cardinality foreign keys. Building a `from_index`
        // (FK -> Vec<row>) for this would be prohibitively expensive.
        let rows = 1_000_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Id".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 65_536,
            cache: PageCacheConfig { max_entries: 8 },
        };
        let mut fact = ColumnarTableBuilder::new(schema, options);
        for i in 0..rows {
            fact.append_row(&[
                // Make most keys unmatched to stress "virtual blank member" handling.
                formula_columnar::Value::Number((1_000_000usize + i) as f64),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }
        model
            .add_table(Table::from_columnar("Fact", fact.finalize()))
            .unwrap();

        let start = Instant::now();
        model
            .add_relationship(Relationship {
                name: "Fact_Dim".into(),
                from_table: "Fact".into(),
                from_column: "Id".into(),
                to_table: "Dim".into(),
                to_column: "Id".into(),
                cardinality: Cardinality::OneToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();
        let elapsed = start.elapsed();

        assert_eq!(model.relationships.len(), 1);
        let rel = &model.relationships[0];
        assert!(
            rel.from_index.is_none(),
            "columnar fact tables should not build RelationshipInfo::from_index"
        );
        assert!(
            rel.unmatched_fact_rows
                .as_ref()
                .map(|rows| !rows.is_empty())
                .unwrap_or(false),
            "expected some unmatched fact rows"
        );

        println!(
            "relationship_large_columnar_does_not_explode_memory: built relationship over {rows} fact rows in {:?}",
            elapsed
        );
    }

    #[test]
    fn relationship_large_columnar_many_to_many_does_not_explode_memory() {
        if std::env::var_os("FORMULA_DAX_REL_BENCH").is_none() {
            return;
        }

        let mut model = DataModel::new();

        // Small in-memory dimension table with duplicate keys (valid for many-to-many).
        let mut dim = Table::new("Dim", vec!["Id"]);
        for i in 0..100i64 {
            dim.push_row(vec![i.into()]).unwrap();
            dim.push_row(vec![i.into()]).unwrap();
        }
        model.add_table(dim).unwrap();

        // Large columnar fact table. Most keys are unmatched to stress virtual blank member
        // handling without building a `from_index`.
        let rows = 1_000_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Id".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 65_536,
            cache: PageCacheConfig { max_entries: 8 },
        };
        let mut fact = ColumnarTableBuilder::new(schema, options);
        for i in 0..rows {
            fact.append_row(&[
                // Make all keys unmatched.
                formula_columnar::Value::Number((1_000_000usize + i) as f64),
                formula_columnar::Value::Number(1.0),
            ]);
        }
        model
            .add_table(Table::from_columnar("Fact", fact.finalize()))
            .unwrap();

        let start = Instant::now();
        model
            .add_relationship(Relationship {
                name: "Fact_Dim".into(),
                from_table: "Fact".into(),
                from_column: "Id".into(),
                to_table: "Dim".into(),
                to_column: "Id".into(),
                cardinality: Cardinality::ManyToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();
        let rel_elapsed = start.elapsed();

        assert_eq!(model.relationships.len(), 1);
        let rel = &model.relationships[0];
        assert!(
            rel.from_index.is_none(),
            "columnar fact tables should not build RelationshipInfo::from_index for many-to-many relationships"
        );

        model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

        // Selecting BLANK on the dimension side should return all unmatched fact rows.
        let filter = FilterContext::empty().with_column_equals("Dim", "Id", Value::Blank);
        let start = Instant::now();
        let value = model.evaluate_measure("Total", &filter).unwrap();
        let eval_elapsed = start.elapsed();
        assert_eq!(value, Value::from(rows as f64));

        println!(
            "relationship_large_columnar_many_to_many_does_not_explode_memory: built relationship over {rows} fact rows in {:?}, evaluated in {:?}",
            rel_elapsed, eval_elapsed
        );
    }

    #[test]
    fn in_memory_relationships_compute_unmatched_fact_rows() {
        let mut model = DataModel::new();

        let mut dim = Table::new("Dim", vec!["Id"]);
        dim.push_row(vec![1.into()]).unwrap();
        model.add_table(dim).unwrap();

        let mut fact = Table::new("Fact", vec!["Id"]);
        fact.push_row(vec![1.into()]).unwrap();
        fact.push_row(vec![999.into()]).unwrap(); // unmatched
        model.add_table(fact).unwrap();

        model
            .add_relationship(Relationship {
                name: "Fact_Dim".into(),
                from_table: "Fact".into(),
                from_column: "Id".into(),
                to_table: "Dim".into(),
                to_column: "Id".into(),
                cardinality: Cardinality::OneToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();

        let rel = model.relationships.first().expect("relationship exists");
        assert!(rel.from_index.is_some());
        assert!(
            matches!(rel.unmatched_fact_rows.as_ref(), Some(rows) if !rows.is_empty()),
            "expected unmatched fact rows cache to be populated for in-memory relationships"
        );
    }

    #[test]
    fn columnar_relationships_do_not_build_from_row_list_index() {
        let mut model = DataModel::new();

        let mut dim = Table::new("Dim", vec!["Id"]);
        dim.push_row(vec![1.into()]).unwrap();
        model.add_table(dim).unwrap();

        let schema = vec![ColumnSchema {
            name: "Id".to_string(),
            column_type: ColumnType::Number,
        }];
        let options = TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 4 },
        };
        let mut fact = ColumnarTableBuilder::new(schema, options);
        fact.append_row(&[formula_columnar::Value::Number(1.0)]);
        fact.append_row(&[formula_columnar::Value::Number(2.0)]);
        model
            .add_table(Table::from_columnar("Fact", fact.finalize()))
            .unwrap();

        model
            .add_relationship(Relationship {
                name: "Fact_Dim".into(),
                from_table: "Fact".into(),
                from_column: "Id".into(),
                to_table: "Dim".into(),
                to_column: "Id".into(),
                cardinality: Cardinality::OneToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();

        let rel = model.relationships.first().expect("relationship exists");
        assert!(rel.from_index.is_none());
    }

    #[test]
    fn columnar_table_add_column_appends_into_storage() {
        let schema = vec![ColumnSchema {
            name: "X".to_string(),
            column_type: ColumnType::Number,
        }];
        let options = TableOptions {
            page_size_rows: 16,
            cache: PageCacheConfig { max_entries: 2 },
        };

        let mut builder = ColumnarTableBuilder::new(schema, options);
        builder.append_row(&[formula_columnar::Value::Number(10.0)]);
        builder.append_row(&[formula_columnar::Value::Number(5.0)]);

        let mut table = Table::from_columnar("T", builder.finalize());
        assert_eq!(table.columns(), &["X".to_string()]);

        table
            .add_column("Y", vec![20.0.into(), 10.0.into()])
            .unwrap();
        assert_eq!(table.columns(), &["X".to_string(), "Y".to_string()]);
        assert_eq!(table.value(0, "Y"), Some(20.0.into()));
        assert_eq!(table.value(1, "Y"), Some(10.0.into()));

        let col_table = table.columnar_table().unwrap();
        let y_schema = col_table.schema().iter().find(|c| c.name == "Y").unwrap();
        assert_eq!(y_schema.column_type, ColumnType::Number);

        // Hold a reference to the underlying `Arc<ColumnarTable>` to force the "shared Arc"
        // clone-fallback path when appending another column.
        let shared_before_b = table.columnar_table().unwrap().clone();
        assert_eq!(shared_before_b.column_count(), 2);
        let x_chunks_ptr = shared_before_b
            .encoded_chunks(0)
            .expect("X chunks")
            .as_ptr();
        let y_chunks_ptr = shared_before_b
            .encoded_chunks(1)
            .expect("Y chunks")
            .as_ptr();

        table
            .add_column("B", vec![Value::Boolean(true), Value::Blank])
            .unwrap();
        let col_table = table.columnar_table().unwrap();
        assert_eq!(col_table.column_count(), 3);
        assert_eq!(shared_before_b.column_count(), 2);
        assert!(shared_before_b.schema().iter().all(|c| c.name != "B"));
        // Appending a column when the underlying `Arc<ColumnarTable>` is shared should not require
        // deep-cloning the existing encoded chunks.
        assert_eq!(
            col_table.encoded_chunks(0).expect("X chunks").as_ptr(),
            x_chunks_ptr
        );
        assert_eq!(
            col_table.encoded_chunks(1).expect("Y chunks").as_ptr(),
            y_chunks_ptr
        );

        let b_schema = col_table.schema().iter().find(|c| c.name == "B").unwrap();
        assert_eq!(b_schema.column_type, ColumnType::Boolean);
        assert_eq!(table.value(0, "B"), Some(true.into()));
        assert_eq!(table.value(1, "B"), Some(Value::Blank));

        table
            .add_column("AllBlank", vec![Value::Blank, Value::Blank])
            .unwrap();
        let col_table = table.columnar_table().unwrap();
        let blank_schema = col_table
            .schema()
            .iter()
            .find(|c| c.name == "AllBlank")
            .unwrap();
        assert_eq!(blank_schema.column_type, ColumnType::Number);
        assert_eq!(table.value(0, "AllBlank"), Some(Value::Blank));
        assert_eq!(table.value(1, "AllBlank"), Some(Value::Blank));

        let err = table
            .add_column("Y", vec![Value::Blank, Value::Blank])
            .unwrap_err();
        assert!(matches!(
            err,
            DaxError::DuplicateColumn { table, column } if table == "T" && column == "Y"
        ));

        let err = table
            .add_column("TooShort", vec![Value::Blank])
            .unwrap_err();
        assert!(matches!(
            err,
            DaxError::ColumnLengthMismatch { table, column, expected, actual }
                if table == "T" && column == "TooShort" && expected == 2 && actual == 1
        ));

        let err = table
            .add_column("Mixed", vec![1.0.into(), Value::from("x")])
            .unwrap_err();
        assert!(matches!(err, DaxError::Type(_)));
    }
}
