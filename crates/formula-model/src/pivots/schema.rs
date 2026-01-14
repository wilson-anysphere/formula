use serde::{Deserialize, Serialize};
use std::borrow::Cow;
use std::collections::HashSet;
use std::fmt;

use super::{PivotField, PivotKeyPart, PivotSource, ValueField};
/// Canonical reference to a field used by a pivot configuration.
///
/// Pivot tables can be sourced from either:
/// - A worksheet range / pivot cache (field names come from header text)
/// - The workbook Data Model / Power Pivot (fields are `{table,column}` and measures)
///
/// This enum makes the distinction explicit so Data Model pivots no longer depend on
/// ambiguous free-form strings.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum PivotFieldRef {
    /// A worksheet / pivot-cache field identified by the header text.
    CacheFieldName(String),
    /// A Data Model column identified by `{table,column}`.
    DataModelColumn { table: String, column: String },
    /// A Data Model measure identified by its name (without brackets).
    DataModelMeasure(String),
}

impl PivotFieldRef {
    /// Returns a canonical string representation for this field reference.
    ///
    /// - Cache fields use the raw header text.
    /// - Data Model fields use a DAX-like display form (`Table[Column]` / `[Measure]`).
    ///
    /// This is intended for UI labels and for matching against pivot cache column names.
    pub fn canonical_name(&self) -> Cow<'_, str> {
        match self {
            PivotFieldRef::CacheFieldName(name) => Cow::Borrowed(name),
            PivotFieldRef::DataModelColumn { table, column } => {
                Cow::Owned(format!("{table}[{column}]"))
            }
            PivotFieldRef::DataModelMeasure(measure) => Cow::Owned(format!("[{measure}]")),
        }
    }

    /// Returns the underlying worksheet/pivot-cache field name if this reference is backed by a
    /// cache field header.
    pub fn as_cache_field_name(&self) -> Option<&str> {
        match self {
            PivotFieldRef::CacheFieldName(name) => Some(name.as_str()),
            _ => None,
        }
    }

    /// Best-effort, human-friendly string representation of this ref.
    ///
    /// This is intended for diagnostics and UI; it is not a stable serialization format.
    pub fn display_string(&self) -> String {
        match self {
            PivotFieldRef::CacheFieldName(name) => name.clone(),
            PivotFieldRef::DataModelColumn { table, column } => format!("{table}[{column}]"),
            PivotFieldRef::DataModelMeasure(name) => format!("[{name}]"),
        }
    }

    /// Best-effort parse of an unstructured field identifier.
    ///
    /// This mirrors the behavior of the `Deserialize` implementation when parsing a string:
    /// - `[Measure]` => `DataModelMeasure("Measure")`
    /// - `Table[Column]` (or `'Table Name'[Column]`) => `DataModelColumn { table, column }`
    /// - otherwise => `CacheFieldName(raw)`
    pub fn from_unstructured(raw: &str) -> Self {
        if let Some(measure) = parse_dax_measure_ref(raw) {
            return PivotFieldRef::DataModelMeasure(measure);
        }
        if let Some((table, column)) = parse_dax_column_ref(raw) {
            return PivotFieldRef::DataModelColumn { table, column };
        }
        PivotFieldRef::CacheFieldName(raw.to_string())
    }
}

impl fmt::Display for PivotFieldRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            // Keep cache field refs backward compatible: display the field name itself.
            PivotFieldRef::CacheFieldName(name) => f.write_str(name),
            PivotFieldRef::DataModelColumn { table, column } => {
                // Prefer a DAX-like display shape to make debugging/logging consistent with Excel.
                // Always quote the table name to avoid ambiguity (spaces, special chars, etc).
                let escaped_table = table.replace('\'', "''");
                write!(f, "'{escaped_table}'[{column}]")
            }
            PivotFieldRef::DataModelMeasure(name) => write!(f, "[{name}]"),
        }
    }
}
impl From<String> for PivotFieldRef {
    fn from(value: String) -> Self {
        PivotFieldRef::CacheFieldName(value)
    }
}

impl From<&str> for PivotFieldRef {
    fn from(value: &str) -> Self {
        PivotFieldRef::CacheFieldName(value.to_string())
    }
}
impl PartialEq<&str> for PivotFieldRef {
    fn eq(&self, other: &&str) -> bool {
        matches!(self, PivotFieldRef::CacheFieldName(name) if name == other)
    }
}
impl Serialize for PivotFieldRef {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        match self {
            // Keep worksheet/cache refs backward compatible: serialize as a plain string.
            PivotFieldRef::CacheFieldName(name) => serializer.serialize_str(name),
            // Data Model refs are always structured objects to eliminate ambiguity.
            PivotFieldRef::DataModelColumn { table, column } => {
                #[derive(Serialize)]
                struct Column<'a> {
                    table: &'a str,
                    column: &'a str,
                }
                Column { table, column }.serialize(serializer)
            }
            PivotFieldRef::DataModelMeasure(name) => {
                #[derive(Serialize)]
                struct Measure<'a> {
                    measure: &'a str,
                }
                Measure { measure: name }.serialize(serializer)
            }
        }
    }
}

impl<'de> Deserialize<'de> for PivotFieldRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        #[serde(untagged)]
        enum Helper {
            Str(String),
            Column { table: String, column: String },
            Measure { measure: String },
            // Allow `{ name: "…" }` as an alternate structured measure shape.
            MeasureName { name: String },
        }

        match Helper::deserialize(deserializer)? {
            Helper::Str(raw) => {
                if let Some(measure) = parse_dax_measure_ref(&raw) {
                    return Ok(PivotFieldRef::DataModelMeasure(measure));
                }
                if let Some((table, column)) = parse_dax_column_ref(&raw) {
                    return Ok(PivotFieldRef::DataModelColumn { table, column });
                }
                Ok(PivotFieldRef::CacheFieldName(raw))
            }
            Helper::Column { table, column } => {
                Ok(PivotFieldRef::DataModelColumn { table, column })
            }
            Helper::Measure { measure } => Ok(PivotFieldRef::DataModelMeasure(measure)),
            Helper::MeasureName { name } => Ok(PivotFieldRef::DataModelMeasure(name)),
        }
    }
}

/// Parse a DAX column reference of the form `Table[Column]` or `'Table Name'[Column]`.
///
/// Parsing is best-effort:
/// - Trims whitespace around identifiers
/// - Supports single-quoted table names (with `''` escape)
pub fn parse_dax_column_ref(input: &str) -> Option<(String, String)> {
    let s = input.trim();
    if s.is_empty() {
        return None;
    }

    // Quick reject: must contain `[` and end with `]`.
    let open = s.find('[')?;
    if !s.ends_with(']') {
        return None;
    }

    let (raw_table, rest) = s.split_at(open);
    let raw_table = raw_table.trim();
    let rest = rest.trim_start();
    if !rest.starts_with('[') {
        return None;
    }

    let column_body = rest.strip_prefix('[')?;
    let close = column_body.rfind(']')?;
    // Ensure the closing bracket is at the end (ignoring trailing whitespace).
    if column_body[close + 1..].trim() != "" {
        return None;
    }
    let column = column_body[..close].trim();
    if column.is_empty() {
        return None;
    }

    let table = if raw_table.starts_with('\'') {
        parse_dax_quoted_identifier(raw_table)?
    } else {
        let t = raw_table.trim();
        if t.is_empty() {
            return None;
        }
        t.to_string()
    };

    Some((table, column.to_string()))
}

/// Parse a DAX measure reference of the form `[Measure]`.
///
/// Returns the measure name *without* brackets.
pub fn parse_dax_measure_ref(input: &str) -> Option<String> {
    let s = input.trim();
    let s = s.strip_prefix('[')?;
    let s = s.strip_suffix(']')?;
    let inner = s.trim();
    if inner.is_empty() {
        return None;
    }
    // Measures are bracket-only; reject anything that looks like a column ref.
    if inner.contains('[') || inner.contains(']') {
        return None;
    }
    Some(inner.to_string())
}

fn parse_dax_quoted_identifier(raw: &str) -> Option<String> {
    let raw = raw.trim();
    let mut chars = raw.chars();
    if chars.next()? != '\'' {
        return None;
    }

    let mut out = String::new();
    let mut i = 1usize;
    let bytes = raw.as_bytes();
    while i < bytes.len() {
        let b = bytes[i];
        if b == b'\'' {
            // `''` is an escaped quote; a single `'` closes the identifier.
            if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                out.push('\'');
                i += 2;
                continue;
            }

            // Closing quote; ensure the rest is only whitespace.
            if raw[i + 1..].trim().is_empty() {
                return Some(out);
            }
            return None;
        }

        // ASCII fast-path; fall back to char iteration for non-ASCII (rare in table names).
        if b.is_ascii() {
            out.push(b as char);
            i += 1;
            continue;
        }

        // Non-ASCII: find the next char boundary.
        let rest = &raw[i..];
        let ch = rest.chars().next()?;
        out.push(ch);
        i += ch.len_utf8();
    }

    None
}
/// An Excel-style PivotTable *calculated field*.
///
/// In Excel, a calculated field is a named formula that behaves like an extra source column:
/// it is evaluated for each record in the pivot cache and can then be used as a field in the
/// pivot configuration (most commonly in the "Values" area).
///
/// The Formula pivot engine persists the raw formula text and treats the calculated field as
/// an additional cache field named [`CalculatedField::name`] when building or refreshing the
/// pivot cache.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedField {
    pub name: String,
    pub formula: String,
}

/// An Excel-style PivotTable *calculated item*.
///
/// In Excel, a calculated item creates a synthetic member (an "item") inside a specific pivot
/// field. The item is defined by a name and a formula that typically references other items
/// within the same field (for example, creating `Q1` from `Jan + Feb + Mar` inside a `Month`
/// field).
///
/// The Formula pivot engine interprets a calculated item as a post-aggregation transform:
/// after the pivot is grouped by [`CalculatedItem::field`], the engine evaluates
/// [`CalculatedItem::formula`] against the existing items for that field and inserts a new item
/// named [`CalculatedItem::name`] into the pivot results.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CalculatedItem {
    pub field: String,
    pub name: String,
    pub formula: String,
}

/// Filter configuration for a pivot field.
///
/// When `allowed` is `None`, the field is unfiltered (all values allowed). When
/// set, it contains the allowed values (represented using the canonical
/// [`PivotKeyPart`] typed key).
#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterField {
    pub source_field: PivotFieldRef,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub allowed: Option<HashSet<PivotKeyPart>>,
}

/// PivotTable report layout mode (Excel: Compact/Outline/Tabular).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Layout {
    Compact,
    Outline,
    Tabular,
}

impl Default for Layout {
    fn default() -> Self {
        // Keep the default stable for serialization/back-compat.
        Self::Tabular
    }
}

/// Where subtotals should be rendered for row field groupings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalPosition {
    /// Excel default behavior.
    Automatic,
    None,
    Top,
    Bottom,
}

impl Default for SubtotalPosition {
    fn default() -> Self {
        Self::Automatic
    }
}
/// Whether to render grand totals for rows and/or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", default)]
pub struct GrandTotals {
    pub rows: bool,
    pub columns: bool,
}

impl Default for GrandTotals {
    fn default() -> Self {
        Self {
            rows: true,
            columns: true,
        }
    }
}

/// Canonical pivot configuration stored in a [`super::PivotTableModel`] and used
/// by the pivot engine.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PivotConfig {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub row_fields: Vec<PivotField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub column_fields: Vec<PivotField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub value_fields: Vec<ValueField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filter_fields: Vec<FilterField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_fields: Vec<CalculatedField>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub calculated_items: Vec<CalculatedItem>,
    #[serde(default)]
    pub layout: Layout,
    #[serde(default)]
    pub subtotals: SubtotalPosition,
    #[serde(default)]
    pub grand_totals: GrandTotals,
}
impl PivotConfig {
    /// Validate that this config is compatible with the given pivot source.
    ///
    /// For Data Model pivots, row/column/filter fields must reference Data Model columns
    /// explicitly, and value fields must reference either measures or columns with supported
    /// aggregations.
    pub fn validate_for_source(&self, source: &PivotSource) -> Result<(), String> {
        match source {
            PivotSource::DataModel { .. } => {
                let check_col = |label: &str, r: &PivotFieldRef| {
                    if !matches!(r, PivotFieldRef::DataModelColumn { .. }) {
                        return Err(format!(
                            "{label} must reference a data model column {{table,column}}"
                        ));
                    }
                    Ok(())
                };

                for f in &self.row_fields {
                    check_col("row field", &f.source_field)?;
                }
                for f in &self.column_fields {
                    check_col("column field", &f.source_field)?;
                }
                for f in &self.filter_fields {
                    check_col("filter field", &f.source_field)?;
                }

                for f in &self.value_fields {
                    match &f.source_field {
                        PivotFieldRef::DataModelMeasure(_) => {}
                        PivotFieldRef::DataModelColumn { .. } => {
                            if !f.aggregation.is_supported_for_data_model() {
                                return Err(format!(
                                    "aggregation {:?} is not supported for data model value fields",
                                    f.aggregation
                                ));
                            }
                        }
                        PivotFieldRef::CacheFieldName(_) => {
                            return Err(
                                "value field must reference a data model measure or column"
                                    .to_string(),
                            );
                        }
                    }
                }

                Ok(())
            }
            PivotSource::Range { .. }
            | PivotSource::RangeName { .. }
            | PivotSource::NamedRange { .. }
            | PivotSource::Table { .. } => {
                let check_cache = |label: &str, r: &PivotFieldRef| {
                    if !matches!(r, PivotFieldRef::CacheFieldName(_)) {
                        return Err(format!(
                            "{label} must reference a worksheet/cache field for worksheet pivots"
                        ));
                    }
                    Ok(())
                };

                for f in &self.row_fields {
                    check_cache("row field", &f.source_field)?;
                }
                for f in &self.column_fields {
                    check_cache("column field", &f.source_field)?;
                }
                for f in &self.filter_fields {
                    check_cache("filter field", &f.source_field)?;
                }
                for f in &self.value_fields {
                    check_cache("value field", &f.source_field)?;
                }

                Ok(())
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::super::PivotField;
    use super::*;

    use pretty_assertions::assert_eq;

    #[test]
    fn parses_dax_column_refs() {
        assert_eq!(
            parse_dax_column_ref("Table[Column]"),
            Some(("Table".to_string(), "Column".to_string()))
        );
        assert_eq!(
            parse_dax_column_ref("  Table [ Column ]  "),
            Some(("Table".to_string(), "Column".to_string()))
        );
        assert_eq!(
            parse_dax_column_ref("'Dim Product'[Category]"),
            Some(("Dim Product".to_string(), "Category".to_string()))
        );
        assert_eq!(
            parse_dax_column_ref("'O''Reilly'[Name]"),
            Some(("O'Reilly".to_string(), "Name".to_string()))
        );
        assert_eq!(parse_dax_column_ref("[Measure]"), None);
        assert_eq!(parse_dax_column_ref("Table[]"), None);
    }

    #[test]
    fn parses_dax_measure_refs() {
        assert_eq!(
            parse_dax_measure_ref("[Total Sales]"),
            Some("Total Sales".to_string())
        );
        assert_eq!(
            parse_dax_measure_ref(" [ Total Sales ] "),
            Some("Total Sales".to_string())
        );
        assert_eq!(parse_dax_measure_ref("[]"), None);
        assert_eq!(parse_dax_measure_ref("Table[Column]"), None);
    }

    #[test]
    fn pivot_field_ref_serde_back_compat() {
        // Plain strings should decode as cache field names.
        let raw = serde_json::json!("Region");
        let decoded: PivotFieldRef = serde_json::from_value(raw).unwrap();
        assert_eq!(decoded, PivotFieldRef::CacheFieldName("Region".to_string()));

        // DAX-like legacy strings should decode into structured Data Model refs.
        let raw = serde_json::json!("Sales[Amount]");
        let decoded: PivotFieldRef = serde_json::from_value(raw).unwrap();
        assert_eq!(
            decoded,
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string()
            }
        );

        let raw = serde_json::json!("[Total Sales]");
        let decoded: PivotFieldRef = serde_json::from_value(raw).unwrap();
        assert_eq!(
            decoded,
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );

        // Structured objects should decode as-is.
        let raw = serde_json::json!({ "table": "Sales", "column": "Amount" });
        let decoded: PivotFieldRef = serde_json::from_value(raw).unwrap();
        assert_eq!(
            decoded,
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string()
            }
        );
        let raw = serde_json::json!({ "measure": "Total Sales" });
        let decoded: PivotFieldRef = serde_json::from_value(raw).unwrap();
        assert_eq!(
            decoded,
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );

        // Serialization keeps cache refs as plain strings and data model refs as objects.
        assert_eq!(
            serde_json::to_value(&PivotFieldRef::CacheFieldName("Region".to_string())).unwrap(),
            serde_json::json!("Region")
        );
        assert_eq!(
            serde_json::to_value(&PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string()
            })
            .unwrap(),
            serde_json::json!({ "table": "Sales", "column": "Amount" })
        );
        assert_eq!(
            serde_json::to_value(&PivotFieldRef::DataModelMeasure("Total Sales".to_string()))
                .unwrap(),
            serde_json::json!({ "measure": "Total Sales" })
        );
    }

    #[test]
    fn pivot_field_struct_accepts_legacy_string_source_field() {
        let raw = serde_json::json!({
            "sourceField": "Region",
            "sortOrder": "ascending"
        });
        let decoded: PivotField = serde_json::from_value(raw).unwrap();
        assert_eq!(
            decoded.source_field,
            PivotFieldRef::CacheFieldName("Region".to_string())
        );
    }

    #[test]
    fn pivot_field_ref_helpers_and_display_formats() {
        let cache = PivotFieldRef::CacheFieldName("Region".to_string());
        assert_eq!(cache.as_cache_field_name(), Some("Region"));
        assert_eq!(cache.display_string(), "Region");
        assert_eq!(cache.to_string(), "Region");

        let col = PivotFieldRef::DataModelColumn {
            table: "Dim Product".to_string(),
            column: "Category".to_string(),
        };
        assert_eq!(col.as_cache_field_name(), None);
        // `display_string` is intended for UI and uses a minimal DAX-like shape.
        assert_eq!(col.display_string(), "Dim Product[Category]");
        // `Display` always quotes the table name to match `formula_dax` / Excel semantics.
        assert_eq!(col.to_string(), "'Dim Product'[Category]");

        let col_with_quote = PivotFieldRef::DataModelColumn {
            table: "O'Reilly".to_string(),
            column: "Name".to_string(),
        };
        assert_eq!(col_with_quote.to_string(), "'O''Reilly'[Name]");

        let measure = PivotFieldRef::DataModelMeasure("Total Sales".to_string());
        assert_eq!(measure.as_cache_field_name(), None);
        assert_eq!(measure.display_string(), "[Total Sales]");
        assert_eq!(measure.to_string(), "[Total Sales]");
    }

    #[test]
    fn grand_totals_defaults_missing_fields_to_true() {
        let decoded: GrandTotals = serde_json::from_value(serde_json::json!({})).unwrap();
        assert_eq!(
            decoded,
            GrandTotals {
                rows: true,
                columns: true
            }
        );

        let decoded: GrandTotals =
            serde_json::from_value(serde_json::json!({ "rows": false })).unwrap();
        assert_eq!(decoded.rows, false);
        assert_eq!(decoded.columns, true);

        let decoded: GrandTotals =
            serde_json::from_value(serde_json::json!({ "columns": false })).unwrap();
        assert_eq!(decoded.rows, true);
        assert_eq!(decoded.columns, false);

        // Ensure nested defaults work when decoding a pivot config.
        let decoded: PivotConfig =
            serde_json::from_value(serde_json::json!({ "grandTotals": { "rows": false } }))
                .unwrap();
        assert_eq!(decoded.grand_totals.rows, false);
        assert_eq!(decoded.grand_totals.columns, true);
    }

    #[test]
    fn pivot_field_ref_from_and_str_eq_helpers() {
        let from_str: PivotFieldRef = "Region".into();
        assert_eq!(
            from_str,
            PivotFieldRef::CacheFieldName("Region".to_string())
        );
        assert!(from_str == "Region");

        let from_string: PivotFieldRef = "Sales".to_string().into();
        assert_eq!(
            from_string,
            PivotFieldRef::CacheFieldName("Sales".to_string())
        );
        assert!(from_string == "Sales");

        // Non-cache refs should never compare equal to a plain field name.
        assert_ne!(
            PivotFieldRef::DataModelMeasure("Total Sales".to_string()),
            "Total Sales"
        );
    }
}
