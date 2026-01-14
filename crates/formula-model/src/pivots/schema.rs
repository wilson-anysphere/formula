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
    /// - Data Model fields use a minimal DAX-like display form (`Table[Column]` / `[Measure]`).
    ///   Note: we intentionally avoid quoting table names here (even when they contain spaces)
    ///   because pivot caches and UI layers often store/display unquoted table captions.
    ///
    /// This is intended for UI labels and for matching against pivot cache column names.
    pub fn canonical_name(&self) -> Cow<'_, str> {
        match self {
            PivotFieldRef::CacheFieldName(name) => Cow::Borrowed(name),
            // For Data Model refs, use an unquoted DAX-like form (`Table[Column]` / `[Measure]`).
            // Quoting rules are handled by the `Display` impl; this helper is used primarily for
            // UI labels and matching cache column names, which often omit quotes.
            PivotFieldRef::DataModelColumn { table, column } => {
                let column = escape_dax_bracket_identifier(column);
                Cow::Owned(format!("{table}[{column}]"))
            }
            PivotFieldRef::DataModelMeasure(measure) => {
                let measure = escape_dax_bracket_identifier(measure);
                Cow::Owned(format!("[{measure}]"))
            }
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

    /// Backward-compatible alias for [`Self::as_cache_field_name`].
    pub fn cache_field_name(&self) -> Option<&str> {
        self.as_cache_field_name()
    }
    /// Best-effort, human-friendly string representation of this ref.
    ///
    /// This is intended for diagnostics and UI; it is not a stable serialization format.
    pub fn display_string(&self) -> String {
        match self {
            PivotFieldRef::CacheFieldName(name) => name.clone(),
            // For display, use an unquoted `Table[Column]` form even when the table name contains
            // spaces/punctuation. This is friendlier for UI labels while still preserving the
            // `{table,column}` structure.
            PivotFieldRef::DataModelColumn { table, column } => {
                let column = escape_dax_bracket_identifier(column);
                format!("{table}[{column}]")
            }
            PivotFieldRef::DataModelMeasure(name) => {
                let name = escape_dax_bracket_identifier(name);
                format!("[{name}]")
            }
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

    /// Best-effort parse of an untyped field identifier.
    ///
    /// Alias for [`PivotFieldRef::from_unstructured`]; kept for backward compatibility with
    /// earlier pivot APIs that accepted free-form strings.
    pub fn from_untyped(raw: &str) -> Self {
        Self::from_unstructured(raw)
    }

    /// Owned-string variant of [`Self::from_unstructured`].
    ///
    /// This is useful when the caller already has an owned `String` and wants to avoid
    /// an extra allocation in the common `CacheFieldName` case.
    pub fn from_unstructured_owned(raw: String) -> Self {
        if let Some(measure) = parse_dax_measure_ref(&raw) {
            return PivotFieldRef::DataModelMeasure(measure);
        }
        if let Some((table, column)) = parse_dax_column_ref(&raw) {
            return PivotFieldRef::DataModelColumn { table, column };
        }
        PivotFieldRef::CacheFieldName(raw)
    }
}

impl From<String> for PivotFieldRef {
    fn from(value: String) -> Self {
        // Mirror `from_unstructured`, but avoid allocating when the input is a plain cache field.
        if let Some(measure) = parse_dax_measure_ref(&value) {
            return PivotFieldRef::DataModelMeasure(measure);
        }
        if let Some((table, column)) = parse_dax_column_ref(&value) {
            return PivotFieldRef::DataModelColumn { table, column };
        }
        PivotFieldRef::CacheFieldName(value)
    }
}

impl From<&str> for PivotFieldRef {
    fn from(value: &str) -> Self {
        Self::from_unstructured(value)
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

impl fmt::Display for PivotFieldRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            PivotFieldRef::CacheFieldName(name) => f.write_str(name),
            PivotFieldRef::DataModelColumn { table, column } => {
                let table = format_dax_table_identifier(table);
                let column = escape_dax_bracket_identifier(column);
                write!(f, "{table}[{column}]")
            }
            PivotFieldRef::DataModelMeasure(name) => {
                let name = escape_dax_bracket_identifier(name);
                write!(f, "[{name}]")
            }
        }
    }
}

fn dax_identifier_requires_quotes(raw: &str) -> bool {
    // DAX identifiers (when unquoted) follow a limited "identifier" grammar. Everything else
    // (spaces, punctuation, leading digits, etc.) must be wrapped in single quotes.
    //
    // Keep this conservative—quoting is always safe and keeps `Display` stable across table names
    // that contain punctuation or whitespace. Also quote identifiers that collide with keywords
    // like `VAR`/`RETURN`/`IN`.
    let mut chars = raw.chars();
    let Some(first) = chars.next() else {
        return true;
    };

    // DAX allows unquoted identifiers in a conservative "C identifier" form. If the identifier
    // contains anything other than ASCII alphanumerics/underscore, or starts with a non-letter /
    // underscore, quote it.
    if !first.is_ascii_alphabetic() && first != '_' {
        return true;
    }

    let is_keyword = raw.eq_ignore_ascii_case("VAR")
        || raw.eq_ignore_ascii_case("RETURN")
        || raw.eq_ignore_ascii_case("IN");

    chars.any(|c| !(c.is_ascii_alphanumeric() || c == '_')) || is_keyword
}

fn quote_dax_identifier(raw: &str) -> String {
    // DAX quotes identifiers using single quotes; embedded quotes are escaped by doubling: `''`.
    let mut out = String::with_capacity(raw.len() + 2);
    out.push('\'');
    for ch in raw.chars() {
        if ch == '\'' {
            out.push_str("''");
        } else {
            out.push(ch);
        }
    }
    out.push('\'');
    out
}

fn format_dax_table_identifier(raw: &str) -> Cow<'_, str> {
    let raw = raw.trim();
    if raw.is_empty() {
        return Cow::Borrowed("''");
    }
    if dax_identifier_requires_quotes(raw) {
        Cow::Owned(quote_dax_identifier(raw))
    } else {
        Cow::Borrowed(raw)
    }
}

fn escape_dax_bracket_identifier(raw: &str) -> String {
    // In DAX, `]` is escaped as `]]` within `[...]`.
    raw.replace(']', "]]")
}
fn unescape_dax_bracket_identifier(raw: &str) -> String {
    // Best-effort: convert DAX-style `]]` escapes back to `]`.
    //
    // This is intentionally forgiving: if the identifier contains unescaped `]`, we treat it as a
    // literal character. This matches the "best-effort" nature of `parse_dax_column_ref` and helps
    // with legacy/unstructured inputs that may not follow strict DAX escaping rules.
    if !raw.contains(']') {
        return raw.to_string();
    }

    let mut out = String::with_capacity(raw.len());
    let mut chars = raw.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == ']' {
            if chars.peek() == Some(&']') {
                chars.next();
            }
            out.push(']');
        } else {
            out.push(ch);
        }
    }
    out
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
    let column = unescape_dax_bracket_identifier(column);

    let table = if raw_table.starts_with('\'') {
        let table = parse_dax_quoted_identifier(raw_table)?;
        if table.is_empty() {
            return None;
        }
        table
    } else {
        let t = raw_table.trim();
        if t.is_empty() {
            return None;
        }
        t.to_string()
    };

    Some((table, column))
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
    if inner.contains('[') {
        return None;
    }
    Some(unescape_dax_bracket_identifier(inner))
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
        Self::Tabular
    }
}

/// Where subtotals should be rendered for row field groupings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum SubtotalPosition {
    None,
    Top,
    Bottom,
}

impl Default for SubtotalPosition {
    fn default() -> Self {
        Self::None
    }
}

/// Whether to render grand totals for rows and/or columns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
#[serde(default)]
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
        assert_eq!(parse_dax_column_ref("''[Column]"), None);
        assert_eq!(parse_dax_column_ref("'O'Reilly'[Name]"), None);
        assert_eq!(parse_dax_column_ref("'Table'X[Column]"), None);
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
    fn pivot_field_ref_display_formats_dax_identifiers() {
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string()
            }
            .to_string(),
            "Sales[Amount]"
        );
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "A-B".to_string(),
                column: "C".to_string()
            }
            .to_string(),
            "'A-B'[C]"
        );
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "1Sales".to_string(),
                column: "Amount".to_string()
            }
            .to_string(),
            "'1Sales'[Amount]"
        );
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "Dim Product".to_string(),
                column: "Category".to_string()
            }
            .to_string(),
            "'Dim Product'[Category]"
        );
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "O'Reilly".to_string(),
                column: "Name".to_string()
            }
            .to_string(),
            "'O''Reilly'[Name]"
        );
        assert_eq!(
            PivotFieldRef::DataModelColumn {
                table: "T".to_string(),
                column: "Col]Name".to_string()
            }
            .to_string(),
            "T[Col]]Name]"
        );
        assert_eq!(
            PivotFieldRef::DataModelMeasure("Total Sales".to_string()).to_string(),
            "[Total Sales]"
        );
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
    fn pivot_field_ref_from_unstructured_matches_serde_string_rules() {
        for (raw, expected) in [
            (
                "Region",
                PivotFieldRef::CacheFieldName("Region".to_string()),
            ),
            (
                "Sales[Amount]",
                PivotFieldRef::DataModelColumn {
                    table: "Sales".to_string(),
                    column: "Amount".to_string(),
                },
            ),
            (
                "[Total Sales]",
                PivotFieldRef::DataModelMeasure("Total Sales".to_string()),
            ),
        ] {
            assert_eq!(PivotFieldRef::from_unstructured(raw), expected);
            assert_eq!(
                PivotFieldRef::from_unstructured_owned(raw.to_string()),
                expected
            );

            // Ensure the helper stays in sync with the serde string behavior.
            let decoded: PivotFieldRef = serde_json::from_value(serde_json::json!(raw)).unwrap();
            assert_eq!(decoded, expected);
        }
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
    fn pivot_field_new_accepts_pivot_field_refs() {
        let col = PivotField::new(PivotFieldRef::DataModelColumn {
            table: "Sales".to_string(),
            column: "Amount".to_string(),
        });
        assert_eq!(
            col.source_field,
            PivotFieldRef::DataModelColumn {
                table: "Sales".to_string(),
                column: "Amount".to_string(),
            }
        );

        let measure = PivotField::new(PivotFieldRef::DataModelMeasure("Total Sales".to_string()));
        assert_eq!(
            measure.source_field,
            PivotFieldRef::DataModelMeasure("Total Sales".to_string())
        );
    }

    #[test]
    fn pivot_field_ref_helpers_and_display_formats() {
        let cache = PivotFieldRef::CacheFieldName("Region".to_string());
        assert_eq!(cache.as_cache_field_name(), Some("Region"));
        assert_eq!(cache.display_string(), "Region");
        assert_eq!(cache.to_string(), "Region");
        assert_eq!(PivotFieldRef::from("Region"), cache);
        assert_eq!(PivotFieldRef::from("Region".to_string()), cache);
        assert!(cache == "Region");
        assert!(cache != "Other");

        let col = PivotFieldRef::DataModelColumn {
            table: "Dim Product".to_string(),
            column: "Category".to_string(),
        };
        assert_eq!(col.as_cache_field_name(), None);
        // `display_string` is intended for UI and uses a minimal DAX-like shape.
        assert_eq!(col.display_string(), "Dim Product[Category]");
        // `Display` renders a DAX-like identifier form (quoting table names when required).
        assert_eq!(col.to_string(), "'Dim Product'[Category]");
        assert!(col != "Dim Product");

        let col_with_quote = PivotFieldRef::DataModelColumn {
            table: "O'Reilly".to_string(),
            column: "Name".to_string(),
        };
        assert_eq!(col_with_quote.to_string(), "'O''Reilly'[Name]");

        // `]` is escaped as `]]` inside bracketed identifiers.
        let col_with_bracket = PivotFieldRef::DataModelColumn {
            table: "T".to_string(),
            column: "A]B".to_string(),
        };
        assert_eq!(col_with_bracket.to_string(), "T[A]]B]");

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

        // And if `grandTotals` is entirely missing, it should default to true/true.
        let decoded: PivotConfig = serde_json::from_value(serde_json::json!({})).unwrap();
        assert_eq!(decoded.grand_totals, GrandTotals::default());
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
