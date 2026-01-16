use crate::backend::{AggregationKind, AggregationSpec, TableBackend};
use crate::engine::{DaxError, DaxResult, FilterContext, RowContext};
use crate::model::{normalize_ident, Cardinality, RelationshipPathDirection, RowSet, ToIndex};
use crate::parser::{BinaryOp, Expr, UnaryOp};
use crate::{DataModel, DaxEngine, Value};
use formula_columnar::BitVec;
#[cfg(feature = "pivot-model")]
use formula_model::pivots::PivotFieldRef;
#[cfg(feature = "pivot-model")]
use formula_model::pivots::PivotValue;
use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicU8, Ordering as AtomicOrdering};
use std::sync::OnceLock;

/// A group-by column used by the pivot engine.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct GroupByColumn {
    pub table: String,
    pub column: String,
}

impl GroupByColumn {
    pub fn new(table: impl Into<String>, column: impl Into<String>) -> Self {
        Self {
            table: table.into(),
            column: column.into(),
        }
    }
}

/// A measure expression to evaluate for each pivot group.
#[derive(Clone, Debug)]
pub struct PivotMeasure {
    pub name: String,
    pub expression: String,
    pub(crate) parsed: Expr,
}

impl PivotMeasure {
    pub fn new(name: impl Into<String>, expression: impl Into<String>) -> DaxResult<Self> {
        let name = name.into();
        let expression = expression.into();
        let parsed = crate::parser::parse(&expression)?;
        Ok(Self {
            name,
            expression,
            parsed,
        })
    }
}

/// Excel-style aggregation modes for pivot value fields.
///
/// These map onto a subset of DAX scalar aggregations.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ValueFieldAggregation {
    Sum,
    Average,
    Min,
    Max,
    /// Count numeric values (Excel "Count Numbers").
    CountNumbers,
    /// Count non-blank values (Excel "Count").
    Count,
    /// Count distinct values (if the higher layer exposes it).
    DistinctCount,
    // Unsupported (for now).
    Product,
    StdDev,
    StdDevP,
    Var,
    VarP,
}

/// Higher-level pivot configuration for a single value field ("Values" area).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ValueFieldSpec {
    pub source_field: String,
    pub name: String,
    pub aggregation: ValueFieldAggregation,
}

#[cfg(feature = "pivot-model")]
impl From<formula_model::pivots::AggregationType> for ValueFieldAggregation {
    fn from(value: formula_model::pivots::AggregationType) -> Self {
        use formula_model::pivots::AggregationType as Agg;
        match value {
            Agg::Sum => ValueFieldAggregation::Sum,
            Agg::Average => ValueFieldAggregation::Average,
            Agg::Min => ValueFieldAggregation::Min,
            Agg::Max => ValueFieldAggregation::Max,
            Agg::CountNumbers => ValueFieldAggregation::CountNumbers,
            Agg::Count => ValueFieldAggregation::Count,
            Agg::Product => ValueFieldAggregation::Product,
            Agg::StdDev => ValueFieldAggregation::StdDev,
            Agg::StdDevP => ValueFieldAggregation::StdDevP,
            Agg::Var => ValueFieldAggregation::Var,
            Agg::VarP => ValueFieldAggregation::VarP,
        }
    }
}

fn looks_like_measure_ref(source_field: &str) -> bool {
    let source_field = source_field.trim();
    source_field.starts_with('[') && source_field.ends_with(']')
}

fn normalize_measure_ref(source_field: &str) -> String {
    let source_field = source_field.trim();
    if crate::parser::parse(source_field).is_ok() {
        return source_field.to_string();
    }

    let inner = source_field
        .strip_prefix('[')
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(source_field);
    let escaped = crate::ident::escape_dax_bracket_identifier(inner);
    format!("[{escaped}]")
}

fn escape_unparsed_column_ref(source_field: &str) -> String {
    let source_field = source_field.trim();
    let Some(open_bracket) = source_field.rfind('[') else {
        return source_field.to_string();
    };
    if !source_field.ends_with(']') || open_bracket + 1 >= source_field.len() {
        return source_field.to_string();
    }

    let prefix = &source_field[..=open_bracket];
    let inner = &source_field[open_bracket + 1..source_field.len() - 1];
    let escaped = crate::ident::escape_dax_bracket_identifier(inner);
    format!("{prefix}{escaped}]")
}

fn normalize_column_ref(base_table: &str, source_field: &str) -> String {
    let source_field = source_field.trim();
    if source_field.contains('[') && source_field.ends_with(']') && !source_field.starts_with('[') {
        if crate::parser::parse(source_field).is_ok() {
            source_field.to_string()
        } else {
            escape_unparsed_column_ref(source_field)
        }
    } else {
        let escaped = crate::ident::escape_dax_bracket_identifier(source_field);
        format!("{base_table}[{escaped}]")
    }
}

/// Build DAX pivot measures for a list of Excel-style value field specs.
///
/// Mapping notes:
/// - `ValueFieldAggregation::Count` maps to DAX `COUNTA(...)` (counts non-blank values).
/// - `ValueFieldAggregation::CountNumbers` maps to DAX `COUNT(...)` (counts numeric values).
pub fn measures_from_value_fields(
    base_table: &str,
    value_fields: &[ValueFieldSpec],
) -> DaxResult<Vec<PivotMeasure>> {
    value_fields
        .iter()
        .map(|vf| measure_from_value_field(base_table, &vf.source_field, &vf.name, vf.aggregation))
        .collect()
}

#[cfg(feature = "pivot-model")]
pub fn measures_from_pivot_model_value_fields(
    base_table: &str,
    value_fields: &[formula_model::pivots::ValueField],
) -> DaxResult<Vec<PivotMeasure>> {
    value_fields
        .iter()
        .map(|vf| {
            let source_field = pivot_model_field_ref_to_dax_source(base_table, &vf.source_field);
            measure_from_value_field(
                base_table,
                &source_field,
                &vf.name,
                ValueFieldAggregation::from(vf.aggregation),
            )
        })
        .collect()
}

#[cfg(feature = "pivot-model")]
fn pivot_model_field_ref_to_dax_source(_base_table: &str, field: &PivotFieldRef) -> String {
    use crate::ident::{format_dax_column_ref, format_dax_measure_ref};

    match field {
        PivotFieldRef::CacheFieldName(name) => name.clone(),
        PivotFieldRef::DataModelMeasure(name) => format_dax_measure_ref(name),
        PivotFieldRef::DataModelColumn { table, column } => {
            format_dax_column_ref(table, column)
        }
    }
}

fn measure_from_value_field(
    base_table: &str,
    source_field: &str,
    name: &str,
    aggregation: ValueFieldAggregation,
) -> DaxResult<PivotMeasure> {
    let source = source_field.trim();
    if looks_like_measure_ref(source) {
        let normalized = normalize_measure_ref(source);
        return PivotMeasure::new(name.to_string(), normalized);
    }

    let column_ref = normalize_column_ref(base_table, source);
    let expr = match aggregation {
        ValueFieldAggregation::Sum => format!("SUM({column_ref})"),
        ValueFieldAggregation::Average => format!("AVERAGE({column_ref})"),
        ValueFieldAggregation::Min => format!("MIN({column_ref})"),
        ValueFieldAggregation::Max => format!("MAX({column_ref})"),
        ValueFieldAggregation::CountNumbers => format!("COUNT({column_ref})"),
        ValueFieldAggregation::Count => format!("COUNTA({column_ref})"),
        ValueFieldAggregation::DistinctCount => format!("DISTINCTCOUNT({column_ref})"),
        other => {
            return Err(DaxError::Eval(format!(
                "pivot value field aggregation {other:?} is not supported yet"
            )))
        }
    };
    PivotMeasure::new(name.to_string(), expr)
}

/// The result of a pivot/group-by query.
#[derive(Clone, Debug, PartialEq)]
pub struct PivotResult {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<Value>>,
}

/// The result of a pivot query shaped into a 2D grid (crosstab) suitable for rendering like an
/// Excel worksheet pivot table.
///
/// The first row in [`PivotResultGrid::data`] is a header row. Subsequent rows contain the
/// row-axis keys followed by aggregated measure values for each column-axis key.
#[derive(Clone, Debug, PartialEq)]
pub struct PivotResultGrid {
    pub data: Vec<Vec<Value>>,
}

fn canonicalize_table_name(model: &DataModel, table: &str) -> DaxResult<String> {
    model
        .table(table)
        .map(|t| t.name().to_string())
        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))
}

fn canonicalize_group_by_columns(
    model: &DataModel,
    columns: &[GroupByColumn],
) -> DaxResult<Vec<GroupByColumn>> {
    let mut out = Vec::with_capacity(columns.len());
    for col in columns {
        let table_ref = model
            .table(&col.table)
            .ok_or_else(|| DaxError::UnknownTable(col.table.clone()))?;
        let table = table_ref.name().to_string();
        let idx = table_ref
            .column_idx(&col.column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.clone(),
                column: col.column.clone(),
            })?;
        let column = table_ref
            .columns()
            .get(idx)
            .cloned()
            .unwrap_or_else(|| col.column.clone());
        out.push(GroupByColumn { table, column });
    }
    Ok(out)
}

impl PivotResultGrid {
    /// Convert this grid into the workbook model's pivot scalar values.
    ///
    /// This is provided as an optional convenience for consumers that ultimately need to render
    /// the grid via `formula-model` pivot structures.
    #[cfg(feature = "pivot-model")]
    pub fn to_pivot_scalars(&self) -> Vec<Vec<formula_model::pivots::ScalarValue>> {
        self.data
            .iter()
            .map(|row| row.iter().map(|v| v.to_pivot_scalar()).collect())
            .collect()
    }
}

/// Options controlling how [`pivot_crosstab`] shapes the output grid.
///
/// This is intentionally small for MVP rendering. Future flags (grand totals, subtotals, compact
/// layout, etc.) can be added without breaking the call signature by introducing
/// `pivot_crosstab_with_options`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotCrosstabOptions {
    /// Separator used to join multiple column-field values into a single column header label.
    ///
    /// This matches how the worksheet pivot engine flattens multi-level column keys in its header
    /// row (e.g. `2024 / Q1`).
    pub column_field_separator: &'static str,
    /// Separator used between a column-key label and a measure name when rendering multiple
    /// measures (e.g. `A - Total`).
    pub column_measure_separator: &'static str,
    /// If true, include the measure name in the column headers even when there is only a single
    /// measure (e.g. `A - Total` instead of just `A`).
    pub include_measure_name_when_single: bool,
}

impl Default for PivotCrosstabOptions {
    fn default() -> Self {
        Self {
            column_field_separator: " / ",
            column_measure_separator: " - ",
            include_measure_name_when_single: false,
        }
    }
}

/// A [`PivotResult`] where every scalar is represented using the canonical pivot value type.
///
/// This is intended for higher layers (IPC/JS rendering) that need the same scalar representation
/// as the worksheet pivot engine.
#[cfg(feature = "pivot-model")]
#[derive(Clone, Debug, PartialEq)]
pub struct PivotResultPivotValues {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<PivotValue>>,
}

#[cfg(feature = "pivot-model")]
impl PivotResult {
    /// Converts this DAX pivot result into the canonical pivot scalar representation.
    pub fn into_pivot_values(self) -> PivotResultPivotValues {
        PivotResultPivotValues {
            columns: self.columns,
            rows: self
                .rows
                .into_iter()
                .map(|row| row.iter().map(crate::dax_value_to_pivot_value).collect())
                .collect(),
        }
    }
}

/// A [`PivotResultGrid`] where every scalar is represented using the canonical pivot value type.
#[cfg(feature = "pivot-model")]
#[derive(Clone, Debug, PartialEq)]
pub struct PivotResultGridPivotValues {
    pub data: Vec<Vec<PivotValue>>,
}

#[cfg(feature = "pivot-model")]
impl PivotResultGrid {
    /// Converts this DAX pivot grid result into the canonical pivot scalar representation.
    pub fn into_pivot_values(self) -> PivotResultGridPivotValues {
        PivotResultGridPivotValues {
            data: self
                .data
                .into_iter()
                .map(|row| row.iter().map(crate::dax_value_to_pivot_value).collect())
                .collect(),
        }
    }
}

fn header_value_display_string(value: &Value) -> String {
    match value {
        Value::Blank => "(blank)".to_string(),
        Value::Boolean(b) => b.to_string(),
        Value::Number(n) => {
            let n = n.0;
            if n.fract() == 0.0 {
                format!("{}", n as i64)
            } else {
                format!("{n}")
            }
        }
        Value::Text(s) => s.to_string(),
    }
}

fn key_display_string(key: &[Value], separator: &str) -> String {
    if key.is_empty() {
        return String::new();
    }
    key.iter()
        .map(header_value_display_string)
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>()
        .join(separator)
}
fn cmp_value(a: &Value, b: &Value) -> Ordering {
    fn sort_rank(v: &Value) -> u8 {
        // Excel-like pivot ordering:
        //   Number/Date < Text < Boolean < Blank
        match v {
            Value::Number(_) => 0,
            Value::Text(_) => 1,
            Value::Boolean(_) => 2,
            Value::Blank => 3,
        }
    }

    match (a, b) {
        (Value::Number(a), Value::Number(b)) => a.cmp(b),
        (Value::Text(a), Value::Text(b)) => {
            let a = a.as_ref();
            let b = b.as_ref();

            // Pivot tables in Excel sort text case-insensitively by default. Use a casefolded
            // comparison as the primary key, with a deterministic case-sensitive tiebreak so the
            // overall ordering remains total (important because group keys are collected from hash
            // maps/sets).
            let ord = cmp_text_case_insensitive(a, b);
            if ord != Ordering::Equal {
                ord
            } else {
                a.cmp(b)
            }
        }
        (Value::Boolean(a), Value::Boolean(b)) => a.cmp(b),
        (Value::Blank, Value::Blank) => Ordering::Equal,
        _ => sort_rank(a).cmp(&sort_rank(b)),
    }
}

fn cmp_text_case_insensitive(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    // Compare using Unicode-aware uppercasing so semantics match Excel-like case-insensitive
    // ordering for non-ASCII text (e.g. ß -> SS).
    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn cmp_ascii_case_insensitive(a: &str, b: &str) -> Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn cmp_key(a: &[Value], b: &[Value]) -> Ordering {
    for (a, b) in a.iter().zip(b.iter()) {
        let ord = cmp_value(a, b);
        if ord != Ordering::Equal {
            return ord;
        }
    }
    a.len().cmp(&b.len())
}

#[derive(Clone, Copy)]
struct RelatedHop<'a> {
    relationship_idx: usize,
    from_idx: usize,
    to_table_key: &'a str,
    to_idx: usize,
    to_index: &'a ToIndex,
    to_table: &'a crate::model::Table,
}

enum GroupKeyAccessor<'a> {
    Base {
        idx: usize,
    },
    RelatedPath {
        hops: Vec<RelatedHop<'a>>,
        to_column_idx: usize,
    },
}

fn build_group_key_accessors<'a>(
    model: &'a DataModel,
    base_table: &'a str,
    group_by: &'a [GroupByColumn],
    filter: &FilterContext,
) -> DaxResult<(&'a crate::model::Table, Vec<GroupKeyAccessor<'a>>)> {
    let base_table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    let base_table_key = normalize_ident(base_table);

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in filter.relationship_overrides() {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
        }
    }
    let is_relationship_active = |idx: usize, rel: &crate::model::RelationshipInfo| {
        let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
        let is_active = if override_pairs.contains(&pair) {
            filter.relationship_overrides().contains(&idx)
        } else {
            rel.rel.is_active
        };

        is_active && !filter.is_relationship_disabled(idx)
    };

    let mut accessors = Vec::with_capacity(group_by.len());
    for col in group_by {
        let col_table_key = normalize_ident(&col.table);
        if col_table_key == base_table_key {
            let idx =
                base_table_ref
                    .column_idx(&col.column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: base_table.to_string(),
                        column: col.column.clone(),
                    })?;
            accessors.push(GroupKeyAccessor::Base { idx });
            continue;
        }

        let Some(path) = model.find_unique_active_relationship_path(
            base_table,
            &col.table,
            RelationshipPathDirection::ManyToOne,
            |idx, rel| is_relationship_active(idx, rel),
        )?
        else {
            return Err(DaxError::Eval(format!(
                "no active relationship from {base_table} to {} for RELATED",
                col.table
            )));
        };

        let mut hops = Vec::with_capacity(path.len());
        for rel_idx in path {
            let rel_info = model
                .relationships()
                .get(rel_idx)
                .expect("relationship index from path");

            let from_idx = rel_info.from_idx;

            let to_table_ref = model
                .table(&rel_info.rel.to_table)
                .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.to_table.clone()))?;

            hops.push(RelatedHop {
                relationship_idx: rel_idx,
                from_idx,
                to_table_key: rel_info.to_table_key.as_str(),
                to_idx: rel_info.to_idx,
                to_index: &rel_info.to_index,
                to_table: to_table_ref,
            });
        }

        let to_table_ref = model
            .table(&col.table)
            .ok_or_else(|| DaxError::UnknownTable(col.table.clone()))?;
        let to_column_idx =
            to_table_ref
                .column_idx(&col.column)
                .ok_or_else(|| DaxError::UnknownColumn {
                    table: col.table.clone(),
                    column: col.column.clone(),
                })?;

        accessors.push(GroupKeyAccessor::RelatedPath {
            hops,
            to_column_idx,
        });
    }

    Ok((base_table_ref, accessors))
}

fn fill_group_key(
    accessors: &[GroupKeyAccessor<'_>],
    base_table: &crate::model::Table,
    row: usize,
    out: &mut Vec<Value>,
) -> DaxResult<()> {
    out.clear();
    for accessor in accessors {
        let value = match accessor {
            GroupKeyAccessor::Base { idx } => {
                base_table.value_by_idx(row, *idx).unwrap_or(Value::Blank)
            }
            GroupKeyAccessor::RelatedPath {
                hops,
                to_column_idx,
            } => {
                let mut current_table = base_table;
                let mut current_row = row;
                let mut ok = true;
                for hop in hops {
                    let key = current_table
                        .value_by_idx(current_row, hop.from_idx)
                        .unwrap_or(Value::Blank);
                    if key.is_blank() {
                        ok = false;
                        break;
                    }
                    let to_row = match hop.to_index {
                        ToIndex::RowSets { map, .. } => {
                            let Some(to_row_set) = map.get(&key) else {
                                ok = false;
                                break;
                            };
                            match to_row_set {
                                RowSet::One(row) => *row,
                                RowSet::Many(rows) => {
                                    if rows.len() == 1 {
                                        rows[0]
                                    } else {
                                        return Err(DaxError::Eval(format!(
                                            "pivot related group key is ambiguous: key {key} matches multiple rows in {}",
                                            hop.to_table.name()
                                        )));
                                    }
                                }
                            }
                        }
                        ToIndex::KeySet { keys, .. } => {
                            if !keys.contains(&key) {
                                ok = false;
                                break;
                            }
                            let rows =
                                hop.to_table.filter_eq(hop.to_idx, &key).unwrap_or_else(|| {
                                    let mut out = Vec::new();
                                    for row in 0..hop.to_table.row_count() {
                                        let v = hop
                                            .to_table
                                            .value_by_idx(row, hop.to_idx)
                                            .unwrap_or(Value::Blank);
                                        if v == key {
                                            out.push(row);
                                        }
                                    }
                                    out
                                });
                            match rows.as_slice() {
                                [] => {
                                    ok = false;
                                    break;
                                }
                                [row] => *row,
                                _ => {
                                    return Err(DaxError::Eval(format!(
                                        "pivot related group key is ambiguous: key {key} matches multiple rows in {}",
                                        hop.to_table.name()
                                    )));
                                }
                            }
                        }
                    };
                    current_table = hop.to_table;
                    current_row = to_row;
                }
                if ok {
                    current_table
                        .value_by_idx(current_row, *to_column_idx)
                        .unwrap_or(Value::Blank)
                } else {
                    Value::Blank
                }
            }
        };
        out.push(value);
    }
    Ok(())
}

fn requires_many_to_many_grouping(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    filter: &FilterContext,
) -> DaxResult<bool> {
    // Fast path: no relationship traversal needed.
    let base_table_key = normalize_ident(base_table);
    if group_by
        .iter()
        .all(|col| normalize_ident(&col.table) == base_table_key)
    {
        return Ok(false);
    }

    let (_, accessors) = build_group_key_accessors(model, base_table, group_by, filter)?;
    for accessor in &accessors {
        let GroupKeyAccessor::RelatedPath { hops, .. } = accessor else {
            continue;
        };
        for hop in hops {
            if hop.to_index.has_duplicates() {
                return Ok(true);
            }
        }
    }

    Ok(false)
}

#[derive(Clone, Debug)]
enum PlannedExpr {
    Const(Value),
    AggRef(usize),
    Negate(Box<PlannedExpr>),
    IsBlank(Box<PlannedExpr>),
    Binary {
        op: BinaryOp,
        left: Box<PlannedExpr>,
        right: Box<PlannedExpr>,
    },
    Not(Box<PlannedExpr>),
    If {
        cond: Box<PlannedExpr>,
        then_expr: Box<PlannedExpr>,
        else_expr: Option<Box<PlannedExpr>>,
    },
    Divide {
        numerator: Box<PlannedExpr>,
        denominator: Box<PlannedExpr>,
        alternate: Option<Box<PlannedExpr>>,
    },
    Coalesce(Vec<PlannedExpr>),
}

fn coerce_number_planned(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => Some(n.0),
        Value::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Some(0.0),
        Value::Text(_) => None,
    }
}

fn coerce_text_planned(value: &Value) -> Cow<'_, str> {
    match value {
        Value::Text(s) => Cow::Borrowed(s.as_ref()),
        // DAX has nuanced formatting semantics. For now we use Rust's default formatting.
        Value::Number(n) => Cow::Owned(n.0.to_string()),
        // In DAX, BLANK coerces to the empty string for text operations like concatenation.
        Value::Blank => Cow::Borrowed(""),
        // DAX displays boolean values as TRUE/FALSE.
        Value::Boolean(b) => Cow::Borrowed(if *b { "TRUE" } else { "FALSE" }),
    }
}

fn compare_values_planned(op: &BinaryOp, left: &Value, right: &Value) -> Option<bool> {
    let cmp = match (left, right) {
        // Text comparisons (BLANK coerces to empty string).
        (Value::Text(l), Value::Text(r)) => Some(l.as_ref().cmp(r.as_ref())),
        (Value::Text(l), Value::Blank) => Some(l.as_ref().cmp("")),
        (Value::Blank, Value::Text(r)) => Some("".cmp(r.as_ref())),
        (Value::Text(_), _) | (_, Value::Text(_)) => return None,
        // Numeric comparisons (BLANK coerces to 0, TRUE/FALSE to 1/0).
        _ => {
            let l = coerce_number_planned(left)?;
            let r = coerce_number_planned(right)?;
            Some(l.partial_cmp(&r)?)
        }
    }?;

    Some(match op {
        BinaryOp::Equals => cmp == Ordering::Equal,
        BinaryOp::NotEquals => cmp != Ordering::Equal,
        BinaryOp::Less => cmp == Ordering::Less,
        BinaryOp::LessEquals => cmp != Ordering::Greater,
        BinaryOp::Greater => cmp == Ordering::Greater,
        BinaryOp::GreaterEquals => cmp != Ordering::Less,
        _ => return None,
    })
}

fn eval_planned(expr: &PlannedExpr, agg_values: &[Value]) -> Value {
    match expr {
        PlannedExpr::Const(v) => v.clone(),
        PlannedExpr::AggRef(idx) => agg_values.get(*idx).cloned().unwrap_or(Value::Blank),
        PlannedExpr::Negate(inner) => {
            let v = eval_planned(inner, agg_values);
            let Some(n) = coerce_number_planned(&v) else {
                return Value::Blank;
            };
            Value::from(-n)
        }
        PlannedExpr::IsBlank(inner) => {
            let v = eval_planned(inner, agg_values);
            Value::from(v.is_blank())
        }
        PlannedExpr::Binary { op, left, right } => {
            let l = eval_planned(left, agg_values);
            let r = eval_planned(right, agg_values);
            match op {
                BinaryOp::Add | BinaryOp::Subtract | BinaryOp::Multiply | BinaryOp::Divide => {
                    let Some(l) = coerce_number_planned(&l) else {
                        return Value::Blank;
                    };
                    let Some(r) = coerce_number_planned(&r) else {
                        return Value::Blank;
                    };
                    let out = match op {
                        BinaryOp::Add => l + r,
                        BinaryOp::Subtract => l - r,
                        BinaryOp::Multiply => l * r,
                        BinaryOp::Divide => l / r,
                        _ => unreachable!(),
                    };
                    Value::from(out)
                }
                BinaryOp::Concat => {
                    let l = coerce_text_planned(&l);
                    let r = coerce_text_planned(&r);
                    let mut out = String::with_capacity(l.len() + r.len());
                    out.push_str(&l);
                    out.push_str(&r);
                    Value::from(out)
                }
                BinaryOp::Equals
                | BinaryOp::NotEquals
                | BinaryOp::Less
                | BinaryOp::LessEquals
                | BinaryOp::Greater
                | BinaryOp::GreaterEquals => compare_values_planned(op, &l, &r)
                    .map(Value::from)
                    .unwrap_or(Value::Blank),
                BinaryOp::And | BinaryOp::Or => {
                    let Ok(l) = l.truthy() else {
                        return Value::Blank;
                    };
                    let Ok(r) = r.truthy() else {
                        return Value::Blank;
                    };
                    Value::from(match op {
                        BinaryOp::And => l && r,
                        BinaryOp::Or => l || r,
                        _ => unreachable!(),
                    })
                }
                BinaryOp::In => Value::Blank,
            }
        }
        PlannedExpr::Not(inner) => {
            let value = eval_planned(inner, agg_values);
            let Ok(b) = value.truthy() else {
                return Value::Blank;
            };
            Value::from(!b)
        }
        PlannedExpr::If {
            cond,
            then_expr,
            else_expr,
        } => {
            let cond = eval_planned(cond, agg_values);
            let Ok(cond) = cond.truthy() else {
                return Value::Blank;
            };
            if cond {
                eval_planned(then_expr, agg_values)
            } else {
                else_expr
                    .as_ref()
                    .map(|expr| eval_planned(expr, agg_values))
                    .unwrap_or(Value::Blank)
            }
        }
        PlannedExpr::Divide {
            numerator,
            denominator,
            alternate,
        } => {
            let num = eval_planned(numerator, agg_values);
            let denom = eval_planned(denominator, agg_values);
            let Some(denom) = coerce_number_planned(&denom) else {
                return Value::Blank;
            };
            if denom == 0.0 {
                alternate
                    .as_ref()
                    .map(|alt| eval_planned(alt, agg_values))
                    .unwrap_or(Value::Blank)
            } else {
                let Some(num) = coerce_number_planned(&num) else {
                    return Value::Blank;
                };
                Value::from(num / denom)
            }
        }
        PlannedExpr::Coalesce(args) => {
            for arg in args {
                let value = eval_planned(arg, agg_values);
                if !value.is_blank() {
                    return value;
                }
            }
            Value::Blank
        }
    }
}

fn ensure_agg(
    kind: AggregationKind,
    column_idx: Option<usize>,
    agg_specs: &mut Vec<AggregationSpec>,
    agg_map: &mut HashMap<(AggregationKind, Option<usize>), usize>,
) -> usize {
    let key = (kind, column_idx);
    if let Some(&idx) = agg_map.get(&key) {
        return idx;
    }
    let idx = agg_specs.len();
    agg_specs.push(AggregationSpec { kind, column_idx });
    agg_map.insert(key, idx);
    idx
}

fn plan_pivot_expr(
    model: &DataModel,
    base_table: &crate::model::Table,
    base_table_name: &str,
    expr: &Expr,
    depth: usize,
    agg_specs: &mut Vec<AggregationSpec>,
    agg_map: &mut HashMap<(AggregationKind, Option<usize>), usize>,
) -> DaxResult<Option<PlannedExpr>> {
    if depth > 32 {
        return Ok(None);
    }

    match expr {
        Expr::Number(n) => Ok(Some(PlannedExpr::Const(Value::from(*n)))),
        Expr::Text(s) => Ok(Some(PlannedExpr::Const(Value::from(s.clone())))),
        Expr::Boolean(b) => Ok(Some(PlannedExpr::Const(Value::from(*b)))),
        Expr::Measure(name) => {
            let normalized = DataModel::normalize_measure_name(name);
            let key = normalize_ident(normalized);
            let measure = model
                .measures()
                .get(&key)
                .ok_or_else(|| DaxError::UnknownMeasure(name.clone()))?;
            plan_pivot_expr(
                model,
                base_table,
                base_table_name,
                &measure.parsed,
                depth + 1,
                agg_specs,
                agg_map,
            )
        }
        Expr::UnaryOp { op, expr } => match op {
            UnaryOp::Negate => {
                let planned = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    expr,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?;
                Ok(planned.map(|inner| PlannedExpr::Negate(Box::new(inner))))
            }
        },
        Expr::BinaryOp { op, left, right } => match op {
            BinaryOp::Concat
            | BinaryOp::Add
            | BinaryOp::Subtract
            | BinaryOp::Multiply
            | BinaryOp::Divide
            | BinaryOp::Equals
            | BinaryOp::NotEquals
            | BinaryOp::Less
            | BinaryOp::LessEquals
            | BinaryOp::Greater
            | BinaryOp::GreaterEquals
            | BinaryOp::And
            | BinaryOp::Or => {
                let Some(left) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    left,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(right) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    right,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                Ok(Some(PlannedExpr::Binary {
                    op: *op,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            BinaryOp::In => Ok(None),
        },
        Expr::Call { name, args } => crate::with_ascii_uppercase(name, |name| match name {
            "BLANK" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::Blank))),
            "TRUE" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::from(true)))),
            "FALSE" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::from(false)))),
            "ISBLANK" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Some(inner) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    arg,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                Ok(Some(PlannedExpr::IsBlank(Box::new(inner))))
            }
            "IF" => {
                if args.len() < 2 || args.len() > 3 {
                    return Ok(None);
                }
                let Some(cond) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[0],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(then_expr) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[1],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let else_expr = if args.len() == 3 {
                    let Some(expr) = plan_pivot_expr(
                        model,
                        base_table,
                        base_table_name,
                        &args[2],
                        depth + 1,
                        agg_specs,
                        agg_map,
                    )?
                    else {
                        return Ok(None);
                    };
                    Some(Box::new(expr))
                } else {
                    None
                };
                Ok(Some(PlannedExpr::If {
                    cond: Box::new(cond),
                    then_expr: Box::new(then_expr),
                    else_expr,
                }))
            }
            "COALESCE" => {
                if args.is_empty() {
                    return Ok(None);
                }
                let mut planned_args = Vec::with_capacity(args.len());
                for arg in args {
                    let Some(planned) = plan_pivot_expr(
                        model,
                        base_table,
                        base_table_name,
                        arg,
                        depth + 1,
                        agg_specs,
                        agg_map,
                    )?
                    else {
                        return Ok(None);
                    };
                    planned_args.push(planned);
                }
                Ok(Some(PlannedExpr::Coalesce(planned_args)))
            }
            "NOT" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Some(inner) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    arg,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                Ok(Some(PlannedExpr::Not(Box::new(inner))))
            }
            name @ ("AND" | "OR") => {
                let [left, right] = args.as_slice() else {
                    return Ok(None);
                };
                let Some(left) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    left,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(right) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    right,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let op = match name {
                    "AND" => BinaryOp::And,
                    "OR" => BinaryOp::Or,
                    _ => unreachable!(),
                };
                Ok(Some(PlannedExpr::Binary {
                    op,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            "DIVIDE" => {
                if args.len() < 2 || args.len() > 3 {
                    return Ok(None);
                }
                let Some(numerator) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[0],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(denominator) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[1],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let alternate = if args.len() == 3 {
                    let Some(alt) = plan_pivot_expr(
                        model,
                        base_table,
                        base_table_name,
                        &args[2],
                        depth + 1,
                        agg_specs,
                        agg_map,
                    )?
                    else {
                        return Ok(None);
                    };
                    Some(Box::new(alt))
                } else {
                    None
                };
                Ok(Some(PlannedExpr::Divide {
                    numerator: Box::new(numerator),
                    denominator: Box::new(denominator),
                    alternate,
                }))
            }
            // Express AVERAGE as SUM / COUNT so downstream rollups (e.g. star-schema pivots) can
            // aggregate SUM and COUNT independently and compute the final average at the end.
            //
            // This matches the engine's AVERAGE semantics (average of numeric values, ignoring
            // blanks/non-numeric values).
            "AVERAGE" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Ok(None);
                };
                if table != base_table_name {
                    return Ok(None);
                }
                let idx = base_table
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                let sum_idx = ensure_agg(AggregationKind::Sum, Some(idx), agg_specs, agg_map);
                let count_idx =
                    ensure_agg(AggregationKind::CountNumbers, Some(idx), agg_specs, agg_map);
                Ok(Some(PlannedExpr::Divide {
                    numerator: Box::new(PlannedExpr::AggRef(sum_idx)),
                    denominator: Box::new(PlannedExpr::AggRef(count_idx)),
                    alternate: None,
                }))
            }
            name @ ("SUM" | "MIN" | "MAX" | "DISTINCTCOUNT" | "COUNT" | "COUNTA") => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Ok(None);
                };
                if table != base_table_name {
                    return Ok(None);
                }
                let idx = base_table
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                let kind = match name {
                    "SUM" => AggregationKind::Sum,
                    "MIN" => AggregationKind::Min,
                    "MAX" => AggregationKind::Max,
                    "COUNT" => AggregationKind::CountNumbers,
                    "COUNTA" => AggregationKind::CountNonBlank,
                    "DISTINCTCOUNT" => AggregationKind::DistinctCount,
                    _ => unreachable!(),
                };
                let agg_idx = ensure_agg(kind, Some(idx), agg_specs, agg_map);
                Ok(Some(PlannedExpr::AggRef(agg_idx)))
            }
            "COUNTBLANK" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Ok(None);
                };
                if normalize_ident(table) != normalize_ident(base_table_name) {
                    return Ok(None);
                }
                let idx = base_table
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                let count_rows = ensure_agg(AggregationKind::CountRows, None, agg_specs, agg_map);
                let count_non_blank = ensure_agg(
                    AggregationKind::CountNonBlank,
                    Some(idx),
                    agg_specs,
                    agg_map,
                );
                Ok(Some(PlannedExpr::Binary {
                    op: BinaryOp::Subtract,
                    left: Box::new(PlannedExpr::AggRef(count_rows)),
                    right: Box::new(PlannedExpr::AggRef(count_non_blank)),
                }))
            }
            "COUNTROWS" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::TableName(table) = arg else {
                    return Ok(None);
                };
                if normalize_ident(table) != normalize_ident(base_table_name) {
                    return Ok(None);
                }
                let agg_idx = ensure_agg(AggregationKind::CountRows, None, agg_specs, agg_map);
                Ok(Some(PlannedExpr::AggRef(agg_idx)))
            }
            _ => Ok(None),
        }),
        _ => Ok(None),
    }
}

fn pivot_columnar_group_by(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<Option<PivotResult>> {
    let base_table_key = normalize_ident(base_table);
    if group_by
        .iter()
        .any(|c| normalize_ident(&c.table) != base_table_key)
    {
        return Ok(None);
    }

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    if table_ref.columnar_table().is_none() {
        return Ok(None);
    }

    let mut group_idxs = Vec::with_capacity(group_by.len());
    for col in group_by {
        let idx = table_ref
            .column_idx(&col.column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: base_table.to_string(),
                column: col.column.clone(),
            })?;
        group_idxs.push(idx);
    }

    let mut agg_specs: Vec<AggregationSpec> = Vec::new();
    let mut agg_map: HashMap<(AggregationKind, Option<usize>), usize> = HashMap::new();
    let mut plans: Vec<PlannedExpr> = Vec::with_capacity(measures.len());
    for measure in measures {
        let Some(plan) = plan_pivot_expr(
            model,
            table_ref,
            base_table,
            &measure.parsed,
            0,
            &mut agg_specs,
            &mut agg_map,
        )?
        else {
            return Ok(None);
        };
        plans.push(plan);
    }

    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;
    let allowed = if let Some(sets) = row_sets.as_ref() {
        Some(
            sets.get(&base_table_key)
                .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?,
        )
    } else {
        None
    };

    let Some(grouped_rows) = table_ref.group_by_aggregations_mask(&group_idxs, &agg_specs, allowed)
    else {
        return Ok(None);
    };

    let key_len = group_idxs.len();
    let mut rows_out: Vec<Vec<Value>> = Vec::with_capacity(grouped_rows.len());
    for mut row in grouped_rows {
        let agg_values = row.get(key_len..).unwrap_or(&[]);
        let mut measure_values = Vec::with_capacity(plans.len());
        for plan in &plans {
            measure_values.push(eval_planned(plan, agg_values));
        }
        row.truncate(key_len);
        row.extend(measure_values);
        rows_out.push(row);
    }

    rows_out.sort_by(|a, b| cmp_key(&a[..key_len], &b[..key_len]));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    Ok(Some(PivotResult {
        columns,
        rows: rows_out,
    }))
}

fn pivot_columnar_groups_with_measure_eval(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<Option<PivotResult>> {
    let base_table_key = normalize_ident(base_table);
    if group_by.is_empty()
        || group_by
            .iter()
            .any(|c| normalize_ident(&c.table) != base_table_key)
    {
        return Ok(None);
    }

    let engine = DaxEngine::new();

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    if table_ref.columnar_table().is_none() {
        return Ok(None);
    }

    let mut group_idxs = Vec::with_capacity(group_by.len());
    for col in group_by {
        let idx = table_ref
            .column_idx(&col.column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: base_table.to_string(),
                column: col.column.clone(),
            })?;
        group_idxs.push(idx);
    }

    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;
    let allowed = if let Some(sets) = row_sets.as_ref() {
        Some(
            sets.get(&base_table_key)
                .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?,
        )
    } else {
        None
    };

    let Some(mut groups) = table_ref.group_by_aggregations_mask(&group_idxs, &[], allowed) else {
        return Ok(None);
    };

    groups.sort_by(|a, b| cmp_key(a, b));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    let mut rows_out = Vec::with_capacity(groups.len());
    let mut group_filter = filter.clone();
    group_filter.in_scope_columns = group_by
        .iter()
        .map(|c| (normalize_ident(&c.table), normalize_ident(&c.column)))
        .collect();
    for mut key in groups {
        for (col, value) in group_by.iter().zip(key.iter()) {
            group_filter.set_column_equals(&col.table, &col.column, value.clone());
        }

        for measure in measures {
            let value = engine.evaluate_expr(
                model,
                &measure.parsed,
                &group_filter,
                &RowContext::default(),
            )?;
            key.push(value);
        }
        rows_out.push(key);
    }

    Ok(Some(PivotResult {
        columns,
        rows: rows_out,
    }))
}

enum StarSchemaGroupKeyAccessor<'a> {
    Base {
        key_pos: usize,
    },
    Related {
        fk_key_pos: usize,
        to_idx: usize,
        to_index: &'a ToIndex,
        to_table: &'a crate::model::Table,
        to_column_idx: usize,
    },
}

fn pivot_columnar_star_schema_group_by(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<Option<PivotResult>> {
    if std::env::var_os("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA").is_some() {
        return Ok(None);
    }

    if group_by.is_empty() {
        // Preserve the existing grand-total behavior by falling back to the legacy code paths.
        return Ok(None);
    }

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    if table_ref.columnar_table().is_none() {
        return Ok(None);
    }
    let base_table_key = normalize_ident(base_table);

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in filter.relationship_overrides() {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
        }
    }
    let is_relationship_active = |idx: usize, rel: &crate::model::RelationshipInfo| {
        let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
        let is_active = if override_pairs.contains(&pair) {
            filter.relationship_overrides().contains(&idx)
        } else {
            rel.rel.is_active
        };

        is_active && !filter.is_relationship_disabled(idx)
    };

    let mut group_idxs: Vec<usize> = Vec::new();
    let mut idx_to_pos: HashMap<usize, usize> = HashMap::new();
    let mut accessors: Vec<StarSchemaGroupKeyAccessor<'_>> = Vec::with_capacity(group_by.len());
    let mut base_group_idxs: HashSet<usize> = HashSet::new();

    for col in group_by {
        if normalize_ident(&col.table) == base_table_key {
            let idx = table_ref
                .column_idx(&col.column)
                .ok_or_else(|| DaxError::UnknownColumn {
                    table: base_table.to_string(),
                    column: col.column.clone(),
                })?;
            base_group_idxs.insert(idx);
            let pos = *idx_to_pos.entry(idx).or_insert_with(|| {
                let pos = group_idxs.len();
                group_idxs.push(idx);
                pos
            });
            accessors.push(StarSchemaGroupKeyAccessor::Base { key_pos: pos });
            continue;
        }

        let Some(path) = model.find_unique_active_relationship_path(
            base_table,
            &col.table,
            RelationshipPathDirection::ManyToOne,
            |idx, rel| is_relationship_active(idx, rel),
        )?
        else {
            // `pivot()` should ultimately error for unsupported RELATED columns (the planned
            // row-group-by path will do so). Keep this fast path conservative.
            return Ok(None);
        };

        if path.len() != 1 {
            // Multi-hop traversal isn't supported by this fast path (yet).
            return Ok(None);
        }

        let rel_info = model
            .relationships()
            .get(path[0])
            .expect("relationship index from path");
        // This fast path assumes each foreign-key value maps to at most one row on the related
        // table (i.e. "one" side). Many-to-many relationships can map to multiple rows and require
        // full row-wise evaluation.
        if rel_info.rel.cardinality == Cardinality::ManyToMany {
            return Ok(None);
        }

        let from_idx = rel_info.from_idx;

        let pos = *idx_to_pos.entry(from_idx).or_insert_with(|| {
            let pos = group_idxs.len();
            group_idxs.push(from_idx);
            pos
        });

        let to_table_ref = model
            .table(&col.table)
            .ok_or_else(|| DaxError::UnknownTable(col.table.clone()))?;
        let to_column_idx =
            to_table_ref
                .column_idx(&col.column)
                .ok_or_else(|| DaxError::UnknownColumn {
                    table: col.table.clone(),
                    column: col.column.clone(),
                })?;

        accessors.push(StarSchemaGroupKeyAccessor::Related {
            fk_key_pos: pos,
            to_idx: rel_info.to_idx,
            to_index: &rel_info.to_index,
            to_table: to_table_ref,
            to_column_idx,
        });
    }

    let mut agg_specs: Vec<AggregationSpec> = Vec::new();
    let mut agg_map: HashMap<(AggregationKind, Option<usize>), usize> = HashMap::new();
    let mut plans: Vec<PlannedExpr> = Vec::with_capacity(measures.len());
    for measure in measures {
        let Some(plan) = plan_pivot_expr(
            model,
            table_ref,
            base_table,
            &measure.parsed,
            0,
            &mut agg_specs,
            &mut agg_map,
        )?
        else {
            return Ok(None);
        };
        plans.push(plan);
    }

    // The rollup step requires that aggregations be composable across groups. Avoid AVERAGE (needs
    // sum+count rollup) and DISTINCTCOUNT (non-additive) for this first iteration.
    for spec in &agg_specs {
        match spec.kind {
            AggregationKind::Average => return Ok(None),
            AggregationKind::DistinctCount => {
                let Some(idx) = spec.column_idx else {
                    return Ok(None);
                };
                // Allow DISTINCTCOUNT only when it is constant within the *final* group. This is
                // the case when the column itself is part of the user-specified group-by keys.
                if !base_group_idxs.contains(&idx) {
                    return Ok(None);
                }
            }
            _ => {}
        }
    }

    let base_table_key = normalize_ident(base_table);
    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;
    let allowed = if let Some(sets) = row_sets.as_ref() {
        Some(
            sets.get(&base_table_key)
                .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?,
        )
    } else {
        None
    };

    let Some(grouped_rows) = table_ref.group_by_aggregations_mask(&group_idxs, &agg_specs, allowed)
    else {
        return Ok(None);
    };

    #[derive(Clone)]
    enum RollupAggState {
        Sum { sum: f64, count: usize },
        Min { best: Option<f64> },
        Max { best: Option<f64> },
        CountRows { count: i64 },
        CountNonBlank { count: i64 },
        CountNumbers { count: i64 },
        DistinctCountConst { any: bool },
    }

    impl RollupAggState {
        fn new(spec: &AggregationSpec) -> Option<Self> {
            Some(match spec.kind {
                AggregationKind::Sum => RollupAggState::Sum { sum: 0.0, count: 0 },
                AggregationKind::Min => RollupAggState::Min { best: None },
                AggregationKind::Max => RollupAggState::Max { best: None },
                AggregationKind::CountRows => RollupAggState::CountRows { count: 0 },
                AggregationKind::CountNonBlank => RollupAggState::CountNonBlank { count: 0 },
                AggregationKind::CountNumbers => RollupAggState::CountNumbers { count: 0 },
                AggregationKind::DistinctCount => RollupAggState::DistinctCountConst { any: false },
                AggregationKind::Average => return None,
            })
        }

        fn update(&mut self, value: &Value) {
            match self {
                RollupAggState::Sum { sum, count } => {
                    if let Value::Number(n) = value {
                        *sum += n.0;
                        *count += 1;
                    }
                }
                RollupAggState::Min { best } => {
                    if let Value::Number(n) = value {
                        *best = Some(best.map_or(n.0, |current| current.min(n.0)));
                    }
                }
                RollupAggState::Max { best } => {
                    if let Value::Number(n) = value {
                        *best = Some(best.map_or(n.0, |current| current.max(n.0)));
                    }
                }
                RollupAggState::CountRows { count } => {
                    if let Value::Number(n) = value {
                        *count += n.0 as i64;
                    }
                }
                RollupAggState::CountNonBlank { count } => {
                    if let Value::Number(n) = value {
                        *count += n.0 as i64;
                    }
                }
                RollupAggState::CountNumbers { count } => {
                    if let Value::Number(n) = value {
                        *count += n.0 as i64;
                    }
                }
                RollupAggState::DistinctCountConst { any } => {
                    if !value.is_blank() {
                        *any = true;
                    }
                }
            }
        }

        fn finalize(self) -> Value {
            match self {
                RollupAggState::Sum { sum, count } => {
                    if count == 0 {
                        Value::Blank
                    } else {
                        Value::from(sum)
                    }
                }
                RollupAggState::Min { best } => best.map(Value::from).unwrap_or(Value::Blank),
                RollupAggState::Max { best } => best.map(Value::from).unwrap_or(Value::Blank),
                RollupAggState::CountRows { count } => Value::from(count),
                RollupAggState::CountNonBlank { count } => Value::from(count),
                RollupAggState::CountNumbers { count } => Value::from(count),
                RollupAggState::DistinctCountConst { any } => {
                    if any {
                        Value::from(1)
                    } else {
                        Value::Blank
                    }
                }
            }
        }
    }

    let mut state_template: Vec<RollupAggState> = Vec::with_capacity(agg_specs.len());
    for spec in &agg_specs {
        let Some(state) = RollupAggState::new(spec) else {
            return Ok(None);
        };
        state_template.push(state);
    }

    let key_len = group_idxs.len();
    let mut groups: HashMap<Vec<Value>, Vec<RollupAggState>> = HashMap::new();
    let mut key_buf: Vec<Value> = Vec::with_capacity(group_by.len());

    for row in grouped_rows {
        let keys = row.get(..key_len).unwrap_or(&[]);
        let values = row.get(key_len..).unwrap_or(&[]);

        key_buf.clear();
        for accessor in &accessors {
            match accessor {
                StarSchemaGroupKeyAccessor::Base { key_pos } => {
                    key_buf.push(keys.get(*key_pos).cloned().unwrap_or(Value::Blank));
                }
                StarSchemaGroupKeyAccessor::Related {
                    fk_key_pos,
                    to_idx,
                    to_index,
                    to_table,
                    to_column_idx,
                } => {
                    let fk = keys.get(*fk_key_pos).cloned().unwrap_or(Value::Blank);
                    if fk.is_blank() {
                        key_buf.push(Value::Blank);
                    } else {
                        let to_row = match to_index {
                            ToIndex::RowSets { map, .. } => match map.get(&fk) {
                                Some(RowSet::One(row)) => Some(*row),
                                Some(RowSet::Many(_)) => {
                                    // Many-to-many expansion isn't supported by this star-schema
                                    // fast path; fall back to the generic (slower) implementation.
                                    return Ok(None);
                                }
                                None => None,
                            },
                            ToIndex::KeySet { keys, .. } => {
                                if !keys.contains(&fk) {
                                    None
                                } else {
                                    let rows =
                                        to_table.filter_eq(*to_idx, &fk).unwrap_or_else(|| {
                                            let mut out = Vec::new();
                                            for row in 0..to_table.row_count() {
                                                let v = to_table
                                                    .value_by_idx(row, *to_idx)
                                                    .unwrap_or(Value::Blank);
                                                if v == fk {
                                                    out.push(row);
                                                }
                                            }
                                            out
                                        });
                                    match rows.as_slice() {
                                        [] => None,
                                        [row] => Some(*row),
                                        _ => return Ok(None),
                                    }
                                }
                            }
                        };

                        if let Some(to_row) = to_row {
                            key_buf.push(
                                to_table
                                    .value_by_idx(to_row, *to_column_idx)
                                    .unwrap_or(Value::Blank),
                            );
                        } else {
                            key_buf.push(Value::Blank);
                        }
                    }
                }
            }
        }

        if let Some(states) = groups.get_mut(key_buf.as_slice()) {
            for (state, value) in states.iter_mut().zip(values) {
                state.update(value);
            }
            continue;
        }

        let mut states = state_template.clone();
        for (state, value) in states.iter_mut().zip(values) {
            state.update(value);
        }
        groups.insert(key_buf.clone(), states);
    }

    let final_key_len = group_by.len();
    let mut rows_out: Vec<Vec<Value>> = Vec::with_capacity(groups.len());
    for (key, states) in groups {
        let agg_values: Vec<Value> = states.into_iter().map(RollupAggState::finalize).collect();
        let mut row = key;
        for plan in &plans {
            row.push(eval_planned(plan, &agg_values));
        }
        rows_out.push(row);
    }

    rows_out.sort_by(|a, b| cmp_key(&a[..final_key_len], &b[..final_key_len]));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    Ok(Some(PivotResult {
        columns,
        rows: rows_out,
    }))
}

fn pivot_planned_row_group_by(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<Option<PivotResult>> {
    if group_by.is_empty() {
        // When there are no group keys, pivot becomes a single "grand total" row. Preserve existing
        // behavior (including any backend-specific stat fast paths) by falling back to the legacy
        // per-group evaluation.
        return Ok(None);
    }

    let (table_ref, group_key_accessors) =
        build_group_key_accessors(model, base_table, group_by, filter)?;
    let mut agg_specs: Vec<AggregationSpec> = Vec::new();
    let mut agg_map: HashMap<(AggregationKind, Option<usize>), usize> = HashMap::new();
    let mut plans: Vec<PlannedExpr> = Vec::with_capacity(measures.len());
    for measure in measures {
        let Some(plan) = plan_pivot_expr(
            model,
            table_ref,
            base_table,
            &measure.parsed,
            0,
            &mut agg_specs,
            &mut agg_map,
        )?
        else {
            return Ok(None);
        };
        plans.push(plan);
    }

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
        fn new(spec: &AggregationSpec) -> Self {
            match spec.kind {
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
            }
        }

        fn update(&mut self, spec: &AggregationSpec, table: &crate::model::Table, row: usize) {
            match (self, spec.kind) {
                (AggState::CountRows { count }, AggregationKind::CountRows) => {
                    *count += 1;
                }
                (AggState::CountNonBlank { count }, AggregationKind::CountNonBlank) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if !table
                        .value_by_idx(row, idx)
                        .unwrap_or(Value::Blank)
                        .is_blank()
                    {
                        *count += 1;
                    }
                }
                (AggState::CountNumbers { count }, AggregationKind::CountNumbers) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if matches!(table.value_by_idx(row, idx), Some(Value::Number(_))) {
                        *count += 1;
                    }
                }
                (AggState::Sum { sum, count }, AggregationKind::Sum) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if let Some(Value::Number(n)) = table.value_by_idx(row, idx) {
                        *sum += n.0;
                        *count += 1;
                    }
                }
                (AggState::Avg { sum, count }, AggregationKind::Average) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if let Some(Value::Number(n)) = table.value_by_idx(row, idx) {
                        *sum += n.0;
                        *count += 1;
                    }
                }
                (AggState::Min { best }, AggregationKind::Min) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if let Some(Value::Number(n)) = table.value_by_idx(row, idx) {
                        *best = Some(best.map_or(n.0, |current| current.min(n.0)));
                    }
                }
                (AggState::Max { best }, AggregationKind::Max) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    if let Some(Value::Number(n)) = table.value_by_idx(row, idx) {
                        *best = Some(best.map_or(n.0, |current| current.max(n.0)));
                    }
                }
                (AggState::DistinctCount { set }, AggregationKind::DistinctCount) => {
                    let Some(idx) = spec.column_idx else {
                        return;
                    };
                    let value = table.value_by_idx(row, idx).unwrap_or(Value::Blank);
                    set.insert(value);
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

    let state_template: Vec<AggState> = agg_specs.iter().map(AggState::new).collect();
    let base_table_key = normalize_ident(base_table);
    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;

    let mut groups: HashMap<Vec<Value>, Vec<AggState>> = HashMap::new();
    let mut key_buf: Vec<Value> = Vec::with_capacity(group_by.len());

    let mut process_row = |row: usize| -> DaxResult<()> {
        fill_group_key(&group_key_accessors, table_ref, row, &mut key_buf)?;

        if let Some(states) = groups.get_mut(key_buf.as_slice()) {
            for (state, spec) in states.iter_mut().zip(&agg_specs) {
                state.update(spec, table_ref, row);
            }
            return Ok(());
        }

        let mut states = state_template.clone();
        for (state, spec) in states.iter_mut().zip(&agg_specs) {
            state.update(spec, table_ref, row);
        }
        groups.insert(key_buf.clone(), states);
        Ok(())
    };

    if let Some(sets) = row_sets.as_ref() {
        let allowed = sets
            .get(base_table_key.as_str())
            .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
        for row in allowed.iter_ones() {
            process_row(row)?;
        }
    } else {
        for row in 0..table_ref.row_count() {
            process_row(row)?;
        }
    }

    let key_len = group_by.len();
    let mut rows_out: Vec<Vec<Value>> = Vec::with_capacity(groups.len());
    for (key, states) in groups {
        let agg_values: Vec<Value> = states.into_iter().map(AggState::finalize).collect();
        let mut row = key;
        for plan in &plans {
            row.push(eval_planned(plan, &agg_values));
        }
        rows_out.push(row);
    }

    rows_out.sort_by(|a, b| cmp_key(&a[..key_len], &b[..key_len]));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    Ok(Some(PivotResult {
        columns,
        rows: rows_out,
    }))
}

fn pivot_row_scan(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    let engine = DaxEngine::new();

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    let base_table_key = normalize_ident(base_table);
    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;
    let mut seen: HashSet<Vec<Value>> = HashSet::new();
    let (_, group_key_accessors) = build_group_key_accessors(model, base_table, group_by, filter)?;
    let mut key_buf: Vec<Value> = Vec::with_capacity(group_by.len());

    // Build the set of groups by scanning the base table rows. This ensures we only create
    // groups that actually exist in the fact table under the current filter context.
    let mut process_row = |row: usize| -> DaxResult<()> {
        fill_group_key(&group_key_accessors, table_ref, row, &mut key_buf)?;
        let _ = seen.insert(key_buf.clone());
        Ok(())
    };

    if let Some(sets) = row_sets.as_ref() {
        let allowed = sets
            .get(base_table_key.as_str())
            .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
        for row in allowed.iter_ones() {
            process_row(row)?;
        }
    } else {
        for row in 0..table_ref.row_count() {
            process_row(row)?;
        }
    }

    let mut groups: Vec<Vec<Value>> = seen.into_iter().collect();
    groups.sort_by(|a, b| cmp_key(a, b));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    let mut rows_out = Vec::with_capacity(groups.len());
    let mut group_filter = filter.clone();
    group_filter.in_scope_columns = group_by
        .iter()
        .map(|c| (normalize_ident(&c.table), normalize_ident(&c.column)))
        .collect();
    for key in groups {
        for (col, value) in group_by.iter().zip(key.iter()) {
            group_filter.set_column_equals(&col.table, &col.column, value.clone());
        }

        let mut row = key;
        for measure in measures {
            let value = engine.evaluate_expr(
                model,
                &measure.parsed,
                &group_filter,
                &RowContext::default(),
            )?;
            row.push(value);
        }
        rows_out.push(row);
    }

    Ok(PivotResult {
        columns,
        rows: rows_out,
    })
}

fn pivot_row_scan_many_to_many(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    let engine = DaxEngine::new();

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    let base_table_key = normalize_ident(base_table);

    let row_sets = (!filter.is_empty())
        .then(|| crate::engine::resolve_row_sets(model, filter))
        .transpose()?;

    let (_, group_key_accessors) = build_group_key_accessors(model, base_table, group_by, filter)?;

    let mut seen: HashSet<Vec<Value>> = HashSet::new();
    #[derive(Default)]
    struct PathNode<'a> {
        columns: Vec<(usize, usize)>,
        children: Vec<(RelatedHop<'a>, PathNode<'a>)>,
    }

    fn child_node_mut<'a, 'm>(
        node: &'m mut PathNode<'a>,
        hop: &RelatedHop<'a>,
    ) -> &'m mut PathNode<'a> {
        if let Some(idx) = node
            .children
            .iter()
            .position(|(h, _)| h.relationship_idx == hop.relationship_idx)
        {
            return &mut node.children[idx].1;
        }

        node.children.push((*hop, PathNode::default()));
        let idx = node.children.len().saturating_sub(1);
        &mut node.children[idx].1
    }

    fn next_rows_for_hop(
        table: &crate::model::Table,
        row: usize,
        hop: &RelatedHop<'_>,
        row_sets: Option<&HashMap<String, BitVec>>,
    ) -> Vec<usize> {
        let key = table
            .value_by_idx(row, hop.from_idx)
            .unwrap_or(Value::Blank);
        if key.is_blank() {
            return Vec::new();
        }
        let allowed = row_sets.and_then(|sets| sets.get(hop.to_table_key));
        let mut rows: Vec<usize> = Vec::new();
        match hop.to_index {
            ToIndex::RowSets { map, .. } => {
                let Some(to_row_set) = map.get(&key) else {
                    return Vec::new();
                };
                to_row_set.for_each_row(|to_row| {
                    if allowed
                        .map(|set| to_row < set.len() && set.get(to_row))
                        .unwrap_or(true)
                    {
                        rows.push(to_row);
                    }
                });
            }
            ToIndex::KeySet { keys, .. } => {
                if !keys.contains(&key) {
                    return Vec::new();
                }

                if let Some(matches) = hop.to_table.filter_eq(hop.to_idx, &key) {
                    for to_row in matches {
                        if allowed
                            .map(|set| to_row < set.len() && set.get(to_row))
                            .unwrap_or(true)
                        {
                            rows.push(to_row);
                        }
                    }
                } else if let Some(set) = allowed {
                    // Fallback: scan only the allowed rows and compare values.
                    for to_row in set.iter_ones() {
                        let v = hop
                            .to_table
                            .value_by_idx(to_row, hop.to_idx)
                            .unwrap_or(Value::Blank);
                        if v == key {
                            rows.push(to_row);
                        }
                    }
                } else {
                    // Fallback: scan the full related table and compare values.
                    for to_row in 0..hop.to_table.row_count() {
                        let v = hop
                            .to_table
                            .value_by_idx(to_row, hop.to_idx)
                            .unwrap_or(Value::Blank);
                        if v == key {
                            rows.push(to_row);
                        }
                    }
                }
            }
        }
        rows.sort_unstable();
        rows.dedup();
        rows
    }

    fn collect_keys_for_node<'a>(
        node: &PathNode<'a>,
        table: &crate::model::Table,
        rows: &[usize],
        key_template: Vec<Value>,
        row_sets: Option<&HashMap<String, BitVec>>,
    ) -> DaxResult<Vec<Vec<Value>>> {
        let mut results: Vec<Vec<Value>> = Vec::new();
        let row_opts: Vec<Option<usize>> = if rows.is_empty() {
            vec![None]
        } else {
            rows.iter().copied().map(Some).collect()
        };

        for row_opt in row_opts {
            let mut key = key_template.clone();
            for (pos, col_idx) in &node.columns {
                key[*pos] = match row_opt {
                    Some(row) => table.value_by_idx(row, *col_idx).unwrap_or(Value::Blank),
                    None => Value::Blank,
                };
            }

            let mut partials = vec![key];
            for (hop, child) in &node.children {
                let child_rows = row_opt
                    .map(|row| next_rows_for_hop(table, row, hop, row_sets))
                    .unwrap_or_default();

                let mut next: Vec<Vec<Value>> = Vec::new();
                for partial in partials {
                    next.extend(collect_keys_for_node(
                        child,
                        hop.to_table,
                        &child_rows,
                        partial,
                        row_sets,
                    )?);
                }
                partials = next;
                if partials.is_empty() {
                    break;
                }
            }

            results.extend(partials);
        }

        Ok(results)
    }

    // Build a trie of relationship paths so group keys stay correlated across snowflake hops.
    let mut root: PathNode<'_> = PathNode::default();
    for (pos, accessor) in group_key_accessors.iter().enumerate() {
        match accessor {
            GroupKeyAccessor::Base { idx } => root.columns.push((pos, *idx)),
            GroupKeyAccessor::RelatedPath {
                hops,
                to_column_idx,
            } => {
                let mut node = &mut root;
                for hop in hops {
                    node = child_node_mut(node, hop);
                }
                node.columns.push((pos, *to_column_idx));
            }
        }
    }

    let blank_key = vec![Value::Blank; group_by.len()];
    let mut process_row = |row: usize| -> DaxResult<()> {
        for key in collect_keys_for_node(
            &root,
            table_ref,
            std::slice::from_ref(&row),
            blank_key.clone(),
            row_sets.as_ref(),
        )? {
            seen.insert(key);
        }
        Ok(())
    };

    if let Some(sets) = row_sets.as_ref() {
        let allowed = sets
            .get(base_table_key.as_str())
            .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
        for row in allowed.iter_ones() {
            process_row(row)?;
        }
    } else {
        for row in 0..table_ref.row_count() {
            process_row(row)?;
        }
    }

    let mut groups: Vec<Vec<Value>> = seen.into_iter().collect();
    groups.sort_by(|a, b| cmp_key(a, b));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    let mut rows_out = Vec::with_capacity(groups.len());
    let mut group_filter = filter.clone();
    group_filter.in_scope_columns = group_by
        .iter()
        .map(|c| (normalize_ident(&c.table), normalize_ident(&c.column)))
        .collect();
    for key in groups {
        for (col, value) in group_by.iter().zip(key.iter()) {
            group_filter.set_column_equals(&col.table, &col.column, value.clone());
        }

        let mut row = key;
        for measure in measures {
            let value = engine.evaluate_expr(
                model,
                &measure.parsed,
                &group_filter,
                &RowContext::default(),
            )?;
            row.push(value);
        }
        rows_out.push(row);
    }

    Ok(PivotResult {
        columns,
        rows: rows_out,
    })
}

/// Compute a grouped table suitable for rendering a pivot table.
///
/// This API is intentionally small: it takes a base table (typically the fact table),
/// a set of group-by columns, and a list of measure expressions.
pub fn pivot(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    let base_table = canonicalize_table_name(model, base_table)?;
    let group_by = canonicalize_group_by_columns(model, group_by)?;
    pivot_impl(model, &base_table, &group_by, measures, filter)
}

fn pivot_impl(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    if let Some(result) = pivot_columnar_group_by(model, base_table, group_by, measures, filter)? {
        maybe_trace_pivot_path(PivotPath::ColumnarGroupBy);
        return Ok(result);
    }

    if let Some(result) =
        pivot_columnar_groups_with_measure_eval(model, base_table, group_by, measures, filter)?
    {
        maybe_trace_pivot_path(PivotPath::ColumnarGroupsWithMeasureEval);
        return Ok(result);
    }

    if requires_many_to_many_grouping(model, base_table, group_by, filter)? {
        maybe_trace_pivot_path(PivotPath::RowScan);
        return pivot_row_scan_many_to_many(model, base_table, group_by, measures, filter);
    }

    if let Some(result) =
        pivot_columnar_star_schema_group_by(model, base_table, group_by, measures, filter)?
    {
        maybe_trace_pivot_path(PivotPath::ColumnarStarSchemaGroupBy);
        return Ok(result);
    }

    if let Some(result) = pivot_planned_row_group_by(model, base_table, group_by, measures, filter)?
    {
        maybe_trace_pivot_path(PivotPath::PlannedRowGroupBy);
        return Ok(result);
    }

    maybe_trace_pivot_path(PivotPath::RowScan);
    pivot_row_scan(model, base_table, group_by, measures, filter)
}

/// Compute a pivot table shaped as a crosstab / 2D grid.
///
/// This builds a grouped table by calling [`pivot`] with `row_fields + column_fields`, then
/// reshapes the grouped result into a grid:
/// - The header row contains the row field captions, followed by one column per unique column-key
///   (and per-measure when more than one measure is requested).
/// - Each body row corresponds to a unique row-key.
/// - Missing (row_key, col_key) combinations are filled with [`Value::Blank`].
pub fn pivot_crosstab(
    model: &DataModel,
    base_table: &str,
    row_fields: &[GroupByColumn],
    column_fields: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResultGrid> {
    pivot_crosstab_with_options(
        model,
        base_table,
        row_fields,
        column_fields,
        measures,
        filter,
        &PivotCrosstabOptions::default(),
    )
}

/// Identical to [`pivot_crosstab`], but allows callers to control header formatting through
/// [`PivotCrosstabOptions`].
pub fn pivot_crosstab_with_options(
    model: &DataModel,
    base_table: &str,
    row_fields: &[GroupByColumn],
    column_fields: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
    options: &PivotCrosstabOptions,
) -> DaxResult<PivotResultGrid> {
    if measures.is_empty() {
        return Err(DaxError::Eval(
            "pivot_crosstab requires at least one measure".to_string(),
        ));
    }

    let base_table = canonicalize_table_name(model, base_table)?;
    let row_fields = canonicalize_group_by_columns(model, row_fields)?;
    let column_fields = canonicalize_group_by_columns(model, column_fields)?;

    let mut group_by = Vec::with_capacity(row_fields.len() + column_fields.len());
    group_by.extend_from_slice(&row_fields);
    group_by.extend_from_slice(&column_fields);

    let grouped = pivot_impl(model, &base_table, &group_by, measures, filter)?;

    let row_key_len = row_fields.len();
    let col_key_len = column_fields.len();
    let value_start = row_key_len + col_key_len;

    let mut row_keys: HashSet<Vec<Value>> = HashSet::new();
    let mut col_keys: HashSet<Vec<Value>> = HashSet::new();
    let mut cells: HashMap<Vec<Value>, HashMap<Vec<Value>, Vec<Value>>> = HashMap::new();

    for row in &grouped.rows {
        if row.len() < value_start {
            continue;
        }
        let row_key = row[..row_key_len].to_vec();
        let col_key = row[row_key_len..value_start].to_vec();
        let values = row[value_start..].to_vec();

        row_keys.insert(row_key.clone());
        col_keys.insert(col_key.clone());
        cells.entry(row_key).or_default().insert(col_key, values);
    }

    let mut row_keys: Vec<Vec<Value>> = row_keys.into_iter().collect();
    row_keys.sort_by(|a, b| cmp_key(a, b));

    let mut col_keys: Vec<Vec<Value>> = col_keys.into_iter().collect();
    col_keys.sort_by(|a, b| cmp_key(a, b));

    // Ensure a deterministic (and non-empty) column axis even when there are no column fields.
    if col_keys.is_empty() {
        col_keys.push(Vec::new());
    }

    let mut data: Vec<Vec<Value>> = Vec::new();

    // Header row.
    let mut header: Vec<Value> = Vec::new();
    for col in &row_fields {
        header.push(Value::from(format!("{}[{}]", col.table, col.column)));
    }
    if col_key_len == 0 {
        // No column fields: behave like a normal grouped table (row keys + measures).
        header.extend(measures.iter().map(|m| Value::from(m.name.clone())));
    } else {
        for col_key in &col_keys {
            let key_label = key_display_string(col_key, options.column_field_separator);
            let include_measure_suffix =
                measures.len() > 1 || options.include_measure_name_when_single;
            if key_label.is_empty() {
                for measure in measures {
                    header.push(Value::from(measure.name.clone()));
                }
            } else if !include_measure_suffix {
                header.push(Value::from(key_label));
            } else {
                for measure in measures {
                    header.push(Value::from(format!(
                        "{key_label}{}{}",
                        options.column_measure_separator, measure.name
                    )));
                }
            }
        }
    }
    data.push(header);

    // Body rows.
    for row_key in &row_keys {
        let mut out_row = row_key.clone();
        let row_map = cells.get(row_key.as_slice());
        for col_key in &col_keys {
            let values = row_map
                .and_then(|m| m.get(col_key.as_slice()))
                .cloned()
                .unwrap_or_else(|| vec![Value::Blank; measures.len()]);
            let include_measure_suffix =
                measures.len() > 1 || options.include_measure_name_when_single;
            if !include_measure_suffix && col_key_len > 0 {
                out_row.push(values.into_iter().next().unwrap_or(Value::Blank));
            } else {
                out_row.extend(values);
            }
        }
        data.push(out_row);
    }

    Ok(PivotResultGrid { data })
}

#[derive(Clone, Copy)]
enum PivotPath {
    ColumnarGroupBy = 1 << 0,
    ColumnarGroupsWithMeasureEval = 1 << 1,
    ColumnarStarSchemaGroupBy = 1 << 2,
    PlannedRowGroupBy = 1 << 3,
    RowScan = 1 << 4,
}

impl PivotPath {
    fn label(self) -> &'static str {
        match self {
            PivotPath::ColumnarGroupBy => "columnar_group_by",
            PivotPath::ColumnarGroupsWithMeasureEval => "columnar_groups_with_measure_eval",
            PivotPath::ColumnarStarSchemaGroupBy => "columnar_star_schema_group_by",
            PivotPath::PlannedRowGroupBy => "planned_row_group_by",
            PivotPath::RowScan => "row_scan",
        }
    }
}

fn pivot_trace_enabled() -> bool {
    static ENABLED: OnceLock<bool> = OnceLock::new();
    *ENABLED.get_or_init(|| std::env::var_os("FORMULA_DAX_PIVOT_TRACE").is_some())
}

fn maybe_trace_pivot_path(path: PivotPath) {
    if !pivot_trace_enabled() {
        return;
    }

    static EMITTED: AtomicU8 = AtomicU8::new(0);
    let bit = path as u8;
    let prev = EMITTED.fetch_or(bit, AtomicOrdering::Relaxed);
    if prev & bit == 0 {
        eprintln!("formula-dax pivot path: {}", path.label());
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_columnar::{
        ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
    };
    use std::sync::Arc;
    use std::time::Instant;

    #[test]
    fn measures_from_value_fields_generates_measures_and_column_aggs() {
        let value_fields = vec![
            ValueFieldSpec {
                source_field: "[Total Sales]".into(),
                name: "Total Sales".into(),
                aggregation: ValueFieldAggregation::Sum,
            },
            ValueFieldSpec {
                source_field: "Amount".into(),
                name: "Sum of Amount".into(),
                aggregation: ValueFieldAggregation::Sum,
            },
            ValueFieldSpec {
                source_field: "Sales[Amount]".into(),
                name: "Average Amount".into(),
                aggregation: ValueFieldAggregation::Average,
            },
            ValueFieldSpec {
                source_field: " Amount ".into(),
                name: "Min Amount".into(),
                aggregation: ValueFieldAggregation::Min,
            },
            ValueFieldSpec {
                source_field: "Amount".into(),
                name: "Max Amount".into(),
                aggregation: ValueFieldAggregation::Max,
            },
            ValueFieldSpec {
                source_field: "Amount".into(),
                name: "CountNums Amount".into(),
                aggregation: ValueFieldAggregation::CountNumbers,
            },
            ValueFieldSpec {
                source_field: "Amount".into(),
                name: "Count Amount".into(),
                aggregation: ValueFieldAggregation::Count,
            },
            ValueFieldSpec {
                source_field: "Amount".into(),
                name: "Distinct Amount".into(),
                aggregation: ValueFieldAggregation::DistinctCount,
            },
        ];

        let measures = measures_from_value_fields("Fact", &value_fields).unwrap();
        let expressions: Vec<&str> = measures.iter().map(|m| m.expression.as_str()).collect();
        assert_eq!(
            expressions,
            vec![
                "[Total Sales]",
                "SUM(Fact[Amount])",
                "AVERAGE(Sales[Amount])",
                "MIN(Fact[Amount])",
                "MAX(Fact[Amount])",
                "COUNT(Fact[Amount])",
                "COUNTA(Fact[Amount])",
                "DISTINCTCOUNT(Fact[Amount])",
            ]
        );
    }

    #[test]
    fn measures_from_value_fields_escapes_brackets_in_column_identifiers() {
        let value_fields = vec![ValueFieldSpec {
            source_field: "Amount]USD".into(),
            name: "Sum of Amount]USD".into(),
            aggregation: ValueFieldAggregation::Sum,
        }];

        let measures = measures_from_value_fields("Fact", &value_fields).unwrap();
        assert_eq!(measures[0].expression, "SUM(Fact[Amount]]USD])");
    }

    #[test]
    fn measures_from_value_fields_escapes_brackets_in_qualified_column_refs() {
        let value_fields = vec![ValueFieldSpec {
            source_field: "'Orders'[Amount]USD]".into(),
            name: "Sum of Orders Amount]USD".into(),
            aggregation: ValueFieldAggregation::Sum,
        }];

        let measures = measures_from_value_fields("Fact", &value_fields).unwrap();
        assert_eq!(measures[0].expression, "SUM('Orders'[Amount]]USD])");
    }

    #[test]
    fn measures_from_value_fields_escapes_brackets_in_measure_refs() {
        let value_fields = vec![ValueFieldSpec {
            source_field: "[Total]USD]".into(),
            name: "Total]USD".into(),
            aggregation: ValueFieldAggregation::Sum,
        }];

        let measures = measures_from_value_fields("Fact", &value_fields).unwrap();
        assert_eq!(measures[0].expression, "[Total]]USD]");
    }

    #[test]
    fn measures_from_value_fields_errors_on_unsupported_aggs() {
        let value_fields = vec![ValueFieldSpec {
            source_field: "Amount".into(),
            name: "Product Amount".into(),
            aggregation: ValueFieldAggregation::Product,
        }];

        let err = measures_from_value_fields("Fact", &value_fields).unwrap_err();
        match err {
            DaxError::Eval(message) => {
                assert!(message.contains("Product"));
                assert!(message.contains("not supported"));
            }
            other => panic!("expected eval error, got {other:?}"),
        }
    }

    #[cfg(feature = "pivot-model")]
    #[test]
    fn measures_from_pivot_model_value_fields_maps_aggregation_types() {
        use formula_model::pivots::{AggregationType, PivotFieldRef, ValueField};

        let value_fields = vec![ValueField {
            source_field: PivotFieldRef::CacheFieldName("Amount".to_string()),
            name: "Count Amount".into(),
            aggregation: AggregationType::Count,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }];

        let measures = measures_from_pivot_model_value_fields("Fact", &value_fields).unwrap();
        assert_eq!(measures[0].expression, "COUNTA(Fact[Amount])");
    }

    #[cfg(feature = "pivot-model")]
    #[test]
    fn measures_from_pivot_model_value_fields_escapes_brackets_in_identifiers() {
        use formula_model::pivots::{AggregationType, PivotFieldRef, ValueField};

        let value_fields = vec![
            ValueField {
                source_field: PivotFieldRef::DataModelMeasure("Total]USD".to_string()),
                name: "Total Measure".into(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
            ValueField {
                source_field: PivotFieldRef::DataModelColumn {
                    table: "Orders".to_string(),
                    column: "Amount]USD".to_string(),
                },
                name: "Sum Amount".into(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            },
        ];

        let measures = measures_from_pivot_model_value_fields("Orders", &value_fields).unwrap();
        assert_eq!(measures[0].expression, "[Total]]USD]");
        assert_eq!(measures[1].expression, "SUM('Orders'[Amount]]USD])");
    }

    #[test]
    fn cmp_value_text_sort_is_case_insensitive_with_case_sensitive_tiebreak() {
        // Primary comparison should be case-insensitive (Excel-like), so "a" sorts before "B".
        assert_eq!(
            cmp_value(&Value::from("a"), &Value::from("B")),
            Ordering::Less
        );

        // Tiebreak should be case-sensitive to keep ordering total/deterministic.
        assert_eq!(
            cmp_value(&Value::from("B"), &Value::from("b")),
            Ordering::Less
        );

        let mut values = vec![Value::from("b"), Value::from("a"), Value::from("B")];
        values.sort_by(cmp_value);
        assert_eq!(
            values,
            vec![Value::from("a"), Value::from("B"), Value::from("b")]
        );
    }

    #[test]
    fn pivot_group_key_sorting_matches_case_insensitive_text_order() {
        let mut model = DataModel::new();
        let mut fact = crate::Table::new("Fact", vec!["Group"]);
        // Insert in an order that would be incorrect under a case-sensitive ASCII compare ("B"
        // would come before "a").
        fact.push_row(vec![Value::from("b")]).unwrap();
        fact.push_row(vec![Value::from("a")]).unwrap();
        fact.push_row(vec![Value::from("B")]).unwrap();
        model.add_table(fact).unwrap();

        let result = pivot(
            &model,
            "Fact",
            &[GroupByColumn::new("Fact", "Group")],
            &[],
            &FilterContext::empty(),
        )
        .unwrap();

        let got: Vec<&str> = result
            .rows
            .iter()
            .map(|row| match &row[0] {
                Value::Text(s) => s.as_ref(),
                other => panic!("expected text group key, got {other:?}"),
            })
            .collect();

        assert_eq!(got, vec!["a", "B", "b"]);
    }

    #[test]
    fn cmp_value_sorts_mixed_types_like_excel_pivots() {
        let mut values = vec![
            Value::Blank,
            true.into(),
            false.into(),
            2.0.into(),
            1.0.into(),
            "b".into(),
            "a".into(),
        ];

        values.sort_by(cmp_value);

        assert_eq!(
            values,
            vec![
                1.0.into(),
                2.0.into(),
                "a".into(),
                "b".into(),
                false.into(),
                true.into(),
                Value::Blank,
            ]
        );

        assert_eq!(cmp_value(&1.0.into(), &"x".into()), Ordering::Less);
        assert_eq!(cmp_value(&"x".into(), &false.into()), Ordering::Less);
        assert_eq!(cmp_value(&false.into(), &Value::Blank), Ordering::Less);
    }

    #[test]
    fn pivot_sorting_is_deterministic_for_mixed_type_group_keys() {
        fn build_model(rows: Vec<(Value, f64)>) -> DataModel {
            let mut model = DataModel::new();
            let mut fact = crate::Table::new("Fact", vec!["Key", "Amount"]);
            for (key, amount) in rows {
                fact.push_row(vec![key, amount.into()]).unwrap();
            }
            model.add_table(fact).unwrap();
            model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
            model.add_measure("Rows", "COUNTROWS(Fact)").unwrap();
            model
        }

        let rows_a = vec![
            (2.0.into(), 10.0),
            ("A".into(), 1.0),
            (true.into(), 3.0),
            (false.into(), 4.0),
            (Value::Blank, 5.0),
            (1.0.into(), 6.0),
            ("B".into(), 7.0),
            (Value::Blank, 8.0),
            (2.0.into(), 9.0),
        ];

        // Same rows, different insertion order (should not affect pivot output ordering).
        let rows_b = vec![
            (Value::Blank, 8.0),
            ("B".into(), 7.0),
            (1.0.into(), 6.0),
            (Value::Blank, 5.0),
            (false.into(), 4.0),
            (true.into(), 3.0),
            ("A".into(), 1.0),
            (2.0.into(), 9.0),
            (2.0.into(), 10.0),
        ];

        let model_a = build_model(rows_a);
        let model_b = build_model(rows_b);

        let measures = vec![
            PivotMeasure::new("Rows", "[Rows]").unwrap(),
            PivotMeasure::new("Total", "[Total]").unwrap(),
        ];
        let group_by = vec![GroupByColumn::new("Fact", "Key")];

        let result_a = pivot(
            &model_a,
            "Fact",
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .unwrap();
        let result_b = pivot(
            &model_b,
            "Fact",
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .unwrap();

        assert_eq!(result_a, result_b);
        assert_eq!(
            result_a.rows,
            vec![
                vec![1.0.into(), 1.into(), 6.0.into()],
                vec![2.0.into(), 2.into(), 19.0.into()],
                vec!["A".into(), 1.into(), 1.0.into()],
                vec!["B".into(), 1.into(), 7.0.into()],
                vec![false.into(), 1.into(), 4.0.into()],
                vec![true.into(), 1.into(), 3.0.into()],
                vec![Value::Blank, 2.into(), 13.0.into()],
            ]
        );
    }

    #[test]
    fn pivot_benchmark_old_vs_new_columnar() {
        if std::env::var_os("FORMULA_DAX_PIVOT_BENCH").is_none() {
            return;
        }

        let rows = 1_000_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Group".to_string(),
                column_type: ColumnType::String,
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
        let mut builder = ColumnarTableBuilder::new(schema, options);
        let groups = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
        for i in 0..rows {
            builder.append_row(&[
                formula_columnar::Value::String(Arc::<str>::from(groups[i % groups.len()])),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar("Fact", builder.finalize()))
            .unwrap();
        model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

        let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];
        let group_by = vec![GroupByColumn::new("Fact", "Group")];
        let filter = FilterContext::empty();

        let start = Instant::now();
        let scan = pivot_row_scan(&model, "Fact", &group_by, &measures, &filter).unwrap();
        let scan_elapsed = start.elapsed();

        let start = Instant::now();
        let fast = pivot(&model, "Fact", &group_by, &measures, &filter).unwrap();
        let fast_elapsed = start.elapsed();

        assert_eq!(scan, fast);

        println!(
            "pivot row-scan: {:?}, columnar group-by: {:?} ({:.2}x speedup)",
            scan_elapsed,
            fast_elapsed,
            scan_elapsed.as_secs_f64() / fast_elapsed.as_secs_f64()
        );

        // Star schema benchmark: group by a related dimension attribute instead of a fact column.
        // This is the common case where the previous implementation fell back to per-row decoding.
        let fact_rows = 1_000_000usize;
        let dim_rows = 10_000usize;
        let regions = ["East", "West", "North", "South", "Central"];

        let customers_schema = vec![
            ColumnSchema {
                name: "CustomerId".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Region".to_string(),
                column_type: ColumnType::String,
            },
        ];
        let mut customers = ColumnarTableBuilder::new(customers_schema, options);
        for id in 1..=dim_rows {
            customers.append_row(&[
                formula_columnar::Value::Number(id as f64),
                formula_columnar::Value::String(Arc::<str>::from(regions[id % regions.len()])),
            ]);
        }

        let sales_schema = vec![
            ColumnSchema {
                name: "CustomerId".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let mut sales = ColumnarTableBuilder::new(sales_schema, options);
        for i in 0..fact_rows {
            let customer_id = (i % dim_rows + 1) as f64;
            sales.append_row(&[
                formula_columnar::Value::Number(customer_id),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }

        let mut star_model = DataModel::new();
        star_model
            .add_table(crate::Table::from_columnar(
                "Customers",
                customers.finalize(),
            ))
            .unwrap();
        star_model
            .add_table(crate::Table::from_columnar("Sales", sales.finalize()))
            .unwrap();
        star_model
            .add_relationship(crate::Relationship {
                name: "Sales_Customers".into(),
                from_table: "Sales".into(),
                from_column: "CustomerId".into(),
                to_table: "Customers".into(),
                to_column: "CustomerId".into(),
                cardinality: crate::Cardinality::OneToMany,
                cross_filter_direction: crate::CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: true,
            })
            .unwrap();
        star_model
            .add_measure("Total", "SUM(Sales[Amount])")
            .unwrap();

        let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];
        let group_by = vec![GroupByColumn::new("Customers", "Region")];
        let filter = FilterContext::empty();

        let start = Instant::now();
        let planned_scan =
            pivot_planned_row_group_by(&star_model, "Sales", &group_by, &measures, &filter)
                .unwrap()
                .unwrap();
        let planned_elapsed = start.elapsed();

        let start = Instant::now();
        let fast = pivot(&star_model, "Sales", &group_by, &measures, &filter).unwrap();
        let fast_elapsed = start.elapsed();

        assert_eq!(planned_scan, fast);

        println!(
            "pivot planned row-scan (RELATED): {:?}, columnar star-schema group-by: {:?} ({:.2}x speedup)",
            planned_elapsed,
            fast_elapsed,
            planned_elapsed.as_secs_f64() / fast_elapsed.as_secs_f64()
        );
    }

    #[test]
    fn pivot_columnar_star_schema_group_by_fast_path_returns_result() {
        let options = TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 4 },
        };

        let customers_schema = vec![
            ColumnSchema {
                name: "CustomerId".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Region".to_string(),
                column_type: ColumnType::String,
            },
        ];
        let mut customers = ColumnarTableBuilder::new(customers_schema, options);
        customers.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::String(Arc::<str>::from("East")),
        ]);
        customers.append_row(&[
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::String(Arc::<str>::from("East")),
        ]);
        customers.append_row(&[
            formula_columnar::Value::Number(3.0),
            formula_columnar::Value::String(Arc::<str>::from("West")),
        ]);

        let sales_schema = vec![
            ColumnSchema {
                name: "CustomerId".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let mut sales = ColumnarTableBuilder::new(sales_schema, options);
        sales.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(10.0),
        ]);
        sales.append_row(&[
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::Null,
        ]);
        sales.append_row(&[
            formula_columnar::Value::Number(3.0),
            formula_columnar::Value::Number(7.0),
        ]);

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar(
                "Customers",
                customers.finalize(),
            ))
            .unwrap();
        model
            .add_table(crate::Table::from_columnar("Sales", sales.finalize()))
            .unwrap();
        model
            .add_relationship(crate::Relationship {
                name: "Sales_Customers".into(),
                from_table: "Sales".into(),
                from_column: "CustomerId".into(),
                to_table: "Customers".into(),
                to_column: "CustomerId".into(),
                cardinality: crate::Cardinality::OneToMany,
                cross_filter_direction: crate::CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: true,
            })
            .unwrap();
        model
            .add_measure("Total Sales", "SUM(Sales[Amount])")
            .unwrap();
        model
            .add_measure("Avg Amount", "AVERAGE(Sales[Amount])")
            .unwrap();
        model.add_measure("Rows", "COUNTROWS(Sales)").unwrap();
        model
            .add_measure("Count Numbers", "COUNT(Sales[Amount])")
            .unwrap();
        model
            .add_measure("Count NonBlank", "COUNTA(Sales[Amount])")
            .unwrap();
        model
            .add_measure("Blank Amounts", "COUNTBLANK(Sales[Amount])")
            .unwrap();

        let measures = vec![
            PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
            PivotMeasure::new("Avg Amount", "[Avg Amount]").unwrap(),
            PivotMeasure::new("Rows", "[Rows]").unwrap(),
            PivotMeasure::new("Count Numbers", "[Count Numbers]").unwrap(),
            PivotMeasure::new("Count NonBlank", "[Count NonBlank]").unwrap(),
            PivotMeasure::new("Blank Amounts", "[Blank Amounts]").unwrap(),
        ];
        let group_by = vec![GroupByColumn::new("Customers", "Region")];

        let result = pivot_columnar_star_schema_group_by(
            &model,
            "Sales",
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .unwrap();

        let Some(result) = result else {
            panic!("expected star schema fast path to return Some");
        };

        assert_eq!(
            result.rows,
            vec![
                vec![
                    Value::from("East"),
                    10.0.into(),
                    10.0.into(),
                    2.0.into(),
                    1.0.into(),
                    1.0.into(),
                    1.0.into(),
                ],
                vec![
                    Value::from("West"),
                    7.0.into(),
                    7.0.into(),
                    1.0.into(),
                    1.0.into(),
                    1.0.into(),
                    0.0.into(),
                ],
            ]
        );
    }

    #[test]
    fn pivot_columnar_star_schema_group_by_respects_userelationship_overrides() {
        let options = TableOptions {
            page_size_rows: 64,
            cache: PageCacheConfig { max_entries: 4 },
        };
        let customers_schema = vec![
            ColumnSchema {
                name: "CustomerId".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Region".to_string(),
                column_type: ColumnType::String,
            },
        ];
        let mut customers = ColumnarTableBuilder::new(customers_schema, options);
        customers.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::String(Arc::<str>::from("East")),
        ]);
        customers.append_row(&[
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::String(Arc::<str>::from("West")),
        ]);

        let sales_schema = vec![
            ColumnSchema {
                name: "CustomerId1".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "CustomerId2".to_string(),
                column_type: ColumnType::Number,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let mut sales = ColumnarTableBuilder::new(sales_schema, options);
        sales.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::Number(10.0),
        ]);
        sales.append_row(&[
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::Number(5.0),
        ]);
        sales.append_row(&[
            formula_columnar::Value::Number(2.0),
            formula_columnar::Value::Number(1.0),
            formula_columnar::Value::Number(7.0),
        ]);

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar(
                "Customers",
                customers.finalize(),
            ))
            .unwrap();
        model
            .add_table(crate::Table::from_columnar("Sales", sales.finalize()))
            .unwrap();
        model
            .add_relationship(crate::Relationship {
                name: "Sales_Customers_1".into(),
                from_table: "Sales".into(),
                from_column: "CustomerId1".into(),
                to_table: "Customers".into(),
                to_column: "CustomerId".into(),
                cardinality: crate::Cardinality::OneToMany,
                cross_filter_direction: crate::CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: true,
            })
            .unwrap();
        model
            .add_relationship(crate::Relationship {
                name: "Sales_Customers_2".into(),
                from_table: "Sales".into(),
                from_column: "CustomerId2".into(),
                to_table: "Customers".into(),
                to_column: "CustomerId".into(),
                cardinality: crate::Cardinality::OneToMany,
                cross_filter_direction: crate::CrossFilterDirection::Single,
                is_active: false,
                enforce_referential_integrity: true,
            })
            .unwrap();
        model
            .add_measure("Total Sales", "SUM(Sales[Amount])")
            .unwrap();

        let measures = vec![PivotMeasure::new("Total Sales", "[Total Sales]").unwrap()];
        let group_by = vec![GroupByColumn::new("Customers", "Region")];

        let Some(result) = pivot_columnar_star_schema_group_by(
            &model,
            "Sales",
            &group_by,
            &measures,
            &FilterContext::empty(),
        )
        .unwrap() else {
            panic!("expected star-schema fast path to run");
        };
        assert_eq!(
            result.rows,
            vec![
                vec![Value::from("East"), 15.0.into()],
                vec![Value::from("West"), 7.0.into()],
            ]
        );

        let filter = DaxEngine::new()
            .apply_calculate_filters(
                &model,
                &FilterContext::empty(),
                &["USERELATIONSHIP(Sales[CustomerId2], Customers[CustomerId])"],
            )
            .unwrap();
        let Some(result) =
            pivot_columnar_star_schema_group_by(&model, "Sales", &group_by, &measures, &filter)
                .unwrap()
        else {
            panic!("expected star-schema fast path to run");
        };
        assert_eq!(
            result.rows,
            vec![
                vec![Value::from("East"), 7.0.into()],
                vec![Value::from("West"), 15.0.into()],
            ]
        );
    }

    #[test]
    fn pivot_columnar_group_by_plans_if_and_comparisons() {
        let rows = 10_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Group".to_string(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 1024,
            cache: PageCacheConfig { max_entries: 4 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);
        let groups = ["A", "B", "C", "D"];
        for i in 0..rows {
            builder.append_row(&[
                formula_columnar::Value::String(Arc::<str>::from(groups[i % groups.len()])),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar("Fact", builder.finalize()))
            .unwrap();
        model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
        // Force both true and false branches to occur, and include logical ops in the condition.
        model
            .add_measure(
                "Big Total",
                "IF([Total] > 123000 && [Total] < 126000, [Total], BLANK())",
            )
            .unwrap();

        let measures = vec![PivotMeasure::new("Big Total", "[Big Total]").unwrap()];
        let group_by = vec![GroupByColumn::new("Fact", "Group")];
        let filter = FilterContext::empty();

        let fast = pivot_columnar_group_by(&model, "Fact", &group_by, &measures, &filter)
            .unwrap()
            .expect("expected planned columnar pivot to be available");
        let scan = pivot_row_scan(&model, "Fact", &group_by, &measures, &filter).unwrap();
        assert_eq!(fast, scan);
    }

    #[test]
    fn pivot_columnar_group_by_plans_isblank() {
        let rows = 10_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Group".to_string(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 1024,
            cache: PageCacheConfig { max_entries: 4 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);
        let groups = ["A", "B", "C", "D"];
        for i in 0..rows {
            builder.append_row(&[
                formula_columnar::Value::String(Arc::<str>::from(groups[i % groups.len()])),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar("Fact", builder.finalize()))
            .unwrap();
        model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
        model
            .add_measure("Total (0 if blank)", "IF(ISBLANK([Total]), 0, [Total])")
            .unwrap();

        let measures =
            vec![PivotMeasure::new("Total (0 if blank)", "[Total (0 if blank)]").unwrap()];
        let group_by = vec![GroupByColumn::new("Fact", "Group")];
        let filter = FilterContext::empty();

        let fast = pivot_columnar_group_by(&model, "Fact", &group_by, &measures, &filter)
            .unwrap()
            .expect("expected planned columnar pivot to be available");
        let scan = pivot_row_scan(&model, "Fact", &group_by, &measures, &filter).unwrap();
        assert_eq!(fast, scan);
    }
}
