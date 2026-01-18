use crate::locale::ValueLocaleConfig;
use crate::sort_filter::parse::{parse_text_datetime, parse_text_number};
use crate::sort_filter::sort::datetime_to_excel_serial_1900;
use crate::sort_filter::types::{CellValue, RangeData, RangeRef};
use chrono::{Local, NaiveDate, NaiveDateTime};
use formula_format::{DateSystem, FormatOptions, Value as FormatValue};
use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use thiserror::Error;

use formula_model::autofilter as model_af;

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct FilterViewId(pub String);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterJoin {
    /// Any criterion may match (logical OR).
    Any,
    /// All criteria must match (logical AND).
    All,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FilterValue {
    Text(String),
    Number(f64),
    DateTime(NaiveDateTime),
    Bool(bool),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextMatchKind {
    Contains,
    BeginsWith,
    EndsWith,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextMatch {
    pub kind: TextMatchKind,
    pub pattern: String,
    pub case_sensitive: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum NumberComparison {
    GreaterThan(f64),
    GreaterThanOrEqual(f64),
    LessThan(f64),
    LessThanOrEqual(f64),
    Between { min: f64, max: f64 },
    NotEqual(f64),
}

#[derive(Debug, Clone, PartialEq)]
pub enum DateComparison {
    After(NaiveDateTime),
    Before(NaiveDateTime),
    Between {
        start: NaiveDateTime,
        end: NaiveDateTime,
    },
    OnDate(NaiveDate),
    Today,
    Yesterday,
    Tomorrow,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FilterCriterion {
    Equals(FilterValue),
    TextMatch(TextMatch),
    /// Negated text match (e.g. "doesNotContain").
    ///
    /// This is used for interpreting certain AutoFilter `customFilter/@operator` values that have
    /// explicit negated semantics.
    NotTextMatch(TextMatch),
    Number(NumberComparison),
    Date(DateComparison),
    Blanks,
    NonBlanks,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ColumnFilter {
    pub join: FilterJoin,
    pub criteria: Vec<FilterCriterion>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilter {
    pub range: RangeRef,
    /// Column index relative to `range.start_col`.
    pub columns: BTreeMap<usize, ColumnFilter>,
}

#[derive(Debug, Error)]
pub enum ModelAutoFilterError {
    #[error("autofilter range exceeds engine bounds")]
    RangeOverflow,
    #[error("autofilter column id exceeds engine bounds")]
    ColumnOverflow,
    #[error("allocation failed ({0})")]
    AllocationFailure(&'static str),
}

impl TryFrom<&model_af::SheetAutoFilter> for AutoFilter {
    type Error = ModelAutoFilterError;

    fn try_from(value: &model_af::SheetAutoFilter) -> Result<Self, Self::Error> {
        Self::try_from_model_with_value_locale(value, ValueLocaleConfig::en_us())
    }
}

impl AutoFilter {
    /// Convert a `formula-model` AutoFilter definition into the engine's filter representation,
    /// interpreting legacy `values` payloads using the provided value locale.
    ///
    /// Note: modern schemas should prefer `filter_columns[*].criteria` which already carries typed
    /// numeric/date values; locale-aware parsing is primarily required for the legacy `values` list
    /// which stores everything as strings.
    pub fn try_from_model_with_value_locale(
        value: &model_af::SheetAutoFilter,
        value_locale: ValueLocaleConfig,
    ) -> Result<Self, ModelAutoFilterError> {
        let range = RangeRef {
            start_row: usize::try_from(value.range.start.row)
                .map_err(|_| ModelAutoFilterError::RangeOverflow)?,
            start_col: usize::try_from(value.range.start.col)
                .map_err(|_| ModelAutoFilterError::RangeOverflow)?,
            end_row: usize::try_from(value.range.end.row)
                .map_err(|_| ModelAutoFilterError::RangeOverflow)?,
            end_col: usize::try_from(value.range.end.col)
                .map_err(|_| ModelAutoFilterError::RangeOverflow)?,
        };

        let mut columns: BTreeMap<usize, ColumnFilter> = BTreeMap::new();

        for col in &value.filter_columns {
            let col_id =
                usize::try_from(col.col_id).map_err(|_| ModelAutoFilterError::ColumnOverflow)?;

            let mut criteria: Vec<FilterCriterion> = Vec::new();
            if !col.criteria.is_empty() {
                if criteria.try_reserve_exact(col.criteria.len()).is_err() {
                    debug_assert!(false, "allocation failed (autofilter criteria)");
                    return Err(ModelAutoFilterError::AllocationFailure("autofilter criteria"));
                }
                for c in &col.criteria {
                    if let Some(c) = model_criterion_to_engine(c) {
                        criteria.push(c);
                    }
                }
            } else {
                // Backwards compatibility: older schema used the `values` list only.
                //
                // These are stored as strings, and Excel interprets them using the workbook locale.
                if criteria.try_reserve_exact(col.values.len()).is_err() {
                    debug_assert!(false, "allocation failed (autofilter values)");
                    return Err(ModelAutoFilterError::AllocationFailure("autofilter values"));
                }
                for raw in &col.values {
                    if let Some(n) = parse_text_number(raw, value_locale) {
                        criteria.push(FilterCriterion::Equals(FilterValue::Number(n)));
                    } else if let Some(dt) = parse_text_datetime(raw, value_locale) {
                        criteria.push(FilterCriterion::Equals(FilterValue::DateTime(dt)));
                    } else {
                        criteria.push(FilterCriterion::Equals(FilterValue::Text(raw.clone())));
                    }
                }
            }

            columns.insert(
                col_id,
                ColumnFilter {
                    join: match col.join {
                        model_af::FilterJoin::Any => FilterJoin::Any,
                        model_af::FilterJoin::All => FilterJoin::All,
                    },
                    criteria,
                },
            );
        }

        Ok(Self { range, columns })
    }
}

fn model_criterion_to_engine(c: &model_af::FilterCriterion) -> Option<FilterCriterion> {
    match c {
        model_af::FilterCriterion::Blanks => Some(FilterCriterion::Blanks),
        model_af::FilterCriterion::NonBlanks => Some(FilterCriterion::NonBlanks),
        model_af::FilterCriterion::Equals(v) => {
            Some(FilterCriterion::Equals(model_value_to_engine(v)))
        }
        model_af::FilterCriterion::TextMatch(m) => Some(FilterCriterion::TextMatch(TextMatch {
            kind: match m.kind {
                model_af::TextMatchKind::Contains => TextMatchKind::Contains,
                model_af::TextMatchKind::BeginsWith => TextMatchKind::BeginsWith,
                model_af::TextMatchKind::EndsWith => TextMatchKind::EndsWith,
            },
            pattern: m.pattern.clone(),
            case_sensitive: m.case_sensitive,
        })),
        model_af::FilterCriterion::Number(cmp) => Some(FilterCriterion::Number(match cmp {
            model_af::NumberComparison::GreaterThan(v) => NumberComparison::GreaterThan(*v),
            model_af::NumberComparison::GreaterThanOrEqual(v) => {
                NumberComparison::GreaterThanOrEqual(*v)
            }
            model_af::NumberComparison::LessThan(v) => NumberComparison::LessThan(*v),
            model_af::NumberComparison::LessThanOrEqual(v) => NumberComparison::LessThanOrEqual(*v),
            model_af::NumberComparison::Between { min, max } => NumberComparison::Between {
                min: *min,
                max: *max,
            },
            model_af::NumberComparison::NotEqual(v) => NumberComparison::NotEqual(*v),
        })),
        model_af::FilterCriterion::Date(cmp) => Some(FilterCriterion::Date(match cmp {
            model_af::DateComparison::After(dt) => DateComparison::After(*dt),
            model_af::DateComparison::Before(dt) => DateComparison::Before(*dt),
            model_af::DateComparison::Between { start, end } => DateComparison::Between {
                start: *start,
                end: *end,
            },
            model_af::DateComparison::OnDate(d) => DateComparison::OnDate(*d),
            model_af::DateComparison::Today => DateComparison::Today,
            model_af::DateComparison::Yesterday => DateComparison::Yesterday,
            model_af::DateComparison::Tomorrow => DateComparison::Tomorrow,
        })),
        model_af::FilterCriterion::OpaqueDynamic(d) => match d.filter_type.as_str() {
            "today" => Some(FilterCriterion::Date(DateComparison::Today)),
            "yesterday" => Some(FilterCriterion::Date(DateComparison::Yesterday)),
            "tomorrow" => Some(FilterCriterion::Date(DateComparison::Tomorrow)),
            _ => None,
        },
        model_af::FilterCriterion::OpaqueCustom(c) => {
            // Best-effort: interpret supported OOXML customFilter operator names so filters can be
            // evaluated. Unknown operators remain ignored (but preserved in the model for XLSX
            // round-trip).
            let pattern = c.value.clone().unwrap_or_default();
            match c.operator.as_str() {
                "contains" => Some(FilterCriterion::TextMatch(TextMatch {
                    kind: TextMatchKind::Contains,
                    pattern,
                    case_sensitive: false,
                })),
                "beginsWith" => Some(FilterCriterion::TextMatch(TextMatch {
                    kind: TextMatchKind::BeginsWith,
                    pattern,
                    case_sensitive: false,
                })),
                "endsWith" => Some(FilterCriterion::TextMatch(TextMatch {
                    kind: TextMatchKind::EndsWith,
                    pattern,
                    case_sensitive: false,
                })),
                "doesNotContain" => Some(FilterCriterion::NotTextMatch(TextMatch {
                    kind: TextMatchKind::Contains,
                    pattern,
                    case_sensitive: false,
                })),
                "doesNotBeginWith" => Some(FilterCriterion::NotTextMatch(TextMatch {
                    kind: TextMatchKind::BeginsWith,
                    pattern,
                    case_sensitive: false,
                })),
                "doesNotEndWith" => Some(FilterCriterion::NotTextMatch(TextMatch {
                    kind: TextMatchKind::EndsWith,
                    pattern,
                    case_sensitive: false,
                })),
                _ => None,
            }
        }
    }
}

fn model_value_to_engine(v: &model_af::FilterValue) -> FilterValue {
    match v {
        model_af::FilterValue::Text(s) => FilterValue::Text(s.clone()),
        model_af::FilterValue::Number(n) => FilterValue::Number(*n),
        model_af::FilterValue::DateTime(dt) => FilterValue::DateTime(*dt),
        model_af::FilterValue::Bool(b) => FilterValue::Bool(*b),
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FilterResult {
    /// For each row in the range, `true` if the row is visible.
    pub visible_rows: Vec<bool>,
    /// Sheet row indices that should be hidden.
    pub hidden_sheet_rows: Vec<usize>,
}

#[derive(Debug, Error)]
pub enum FilterError {
    #[error("allocation failed ({0})")]
    AllocationFailure(&'static str),
}

/// A simple filter-view manager for collaboration.
///
/// Each view (collaborator) can have independent filters on the same sheet.
#[derive(Debug, Default)]
pub struct FilterViews {
    views: HashMap<FilterViewId, HashMap<RangeRef, AutoFilter>>,
}

impl FilterViews {
    pub fn set_filter(&mut self, view: FilterViewId, filter: AutoFilter) {
        self.views
            .entry(view)
            .or_default()
            .insert(filter.range, filter);
    }

    pub fn clear_filter(&mut self, view: &FilterViewId, range: RangeRef) {
        if let Some(ranges) = self.views.get_mut(view) {
            ranges.remove(&range);
            if ranges.is_empty() {
                self.views.remove(view);
            }
        }
    }

    pub fn get_filter(&self, view: &FilterViewId, range: RangeRef) -> Option<&AutoFilter> {
        self.views.get(view).and_then(|m| m.get(&range))
    }
}

pub fn apply_autofilter(range: &RangeData, filter: &AutoFilter) -> Result<FilterResult, FilterError> {
    apply_autofilter_with_value_locale(range, filter, ValueLocaleConfig::en_us())
}

pub fn apply_autofilter_with_value_locale(
    range: &RangeData,
    filter: &AutoFilter,
    value_locale: ValueLocaleConfig,
) -> Result<FilterResult, FilterError> {
    let row_count = range.rows.len();
    let mut visible_rows: Vec<bool> = Vec::new();
    if visible_rows.try_reserve_exact(row_count).is_err() {
        debug_assert!(false, "allocation failed (autofilter visible rows={row_count})");
        return Err(FilterError::AllocationFailure("autofilter visible rows"));
    }
    visible_rows.resize(row_count, true);

    // Excel AutoFilter always treats the first row in the ref as the header row.
    if row_count == 0 {
        return Ok(FilterResult {
            visible_rows,
            hidden_sheet_rows: Vec::new(),
        });
    }

    for local_row in 1..row_count {
        let mut row_visible = true;
        for (col_id, col_filter) in &filter.columns {
            let cell = range
                .rows
                .get(local_row)
                .and_then(|r| r.get(*col_id))
                .unwrap_or(&CellValue::Blank);
            if !evaluate_column_filter(cell, col_filter, value_locale) {
                row_visible = false;
                break;
            }
        }
        visible_rows[local_row] = row_visible;
    }

    let mut hidden_sheet_rows: Vec<usize> = Vec::new();
    if hidden_sheet_rows.try_reserve(row_count.saturating_sub(1)).is_err() {
        debug_assert!(false, "allocation failed (autofilter hidden rows)");
        return Err(FilterError::AllocationFailure("autofilter hidden rows"));
    }
    for (local_row, visible) in visible_rows.iter().enumerate().skip(1) {
        if !*visible {
            hidden_sheet_rows.push(range.range.start_row + local_row);
        }
    }

    Ok(FilterResult {
        visible_rows,
        hidden_sheet_rows,
    })
}

fn evaluate_column_filter(
    cell: &CellValue,
    filter: &ColumnFilter,
    value_locale: ValueLocaleConfig,
) -> bool {
    if filter.criteria.is_empty() {
        return true;
    }

    match filter.join {
        FilterJoin::Any => filter
            .criteria
            .iter()
            .any(|c| evaluate_criterion(cell, c, value_locale)),
        FilterJoin::All => filter
            .criteria
            .iter()
            .all(|c| evaluate_criterion(cell, c, value_locale)),
    }
}

fn evaluate_criterion(
    cell: &CellValue,
    criterion: &FilterCriterion,
    value_locale: ValueLocaleConfig,
) -> bool {
    match criterion {
        FilterCriterion::Blanks => is_blank(cell),
        FilterCriterion::NonBlanks => !is_blank(cell),
        FilterCriterion::Equals(value) => equals_value(cell, value, value_locale),
        FilterCriterion::TextMatch(m) => text_match(cell, m, value_locale),
        FilterCriterion::NotTextMatch(m) => !text_match(cell, m, value_locale),
        FilterCriterion::Number(cmp) => number_cmp(cell, cmp, value_locale),
        FilterCriterion::Date(cmp) => date_cmp(cell, cmp, value_locale),
    }
}

fn is_blank(cell: &CellValue) -> bool {
    matches!(cell, CellValue::Blank) || matches!(cell, CellValue::Text(s) if s.trim().is_empty())
}

fn equals_value(cell: &CellValue, value: &FilterValue, value_locale: ValueLocaleConfig) -> bool {
    match value {
        FilterValue::Text(s) => {
            let cell_s = cell_to_string(cell, value_locale);
            crate::value::eq_case_insensitive(cell_s.as_ref(), s)
        }
        FilterValue::Number(n) => coerce_number(cell, value_locale).is_some_and(|v| v == *n),
        FilterValue::Bool(b) => matches!(cell, CellValue::Bool(v) if v == b),
        FilterValue::DateTime(dt) => coerce_datetime(cell, value_locale).is_some_and(|v| v == *dt),
    }
}

fn text_match(cell: &CellValue, m: &TextMatch, value_locale: ValueLocaleConfig) -> bool {
    if m.case_sensitive {
        let cell_s = cell_to_string(cell, value_locale);
        return match m.kind {
            TextMatchKind::Contains => cell_s.contains(m.pattern.as_str()),
            TextMatchKind::BeginsWith => cell_s.starts_with(m.pattern.as_str()),
            TextMatchKind::EndsWith => cell_s.ends_with(m.pattern.as_str()),
        };
    }

    let cell_s = cell_to_string(cell, value_locale);
    if cell_s.is_ascii() && m.pattern.is_ascii() {
        let needle = m.pattern.as_str();
        let hay = cell_s.as_ref();
        return match m.kind {
            TextMatchKind::Contains => ascii_contains_case_insensitive(hay, needle),
            TextMatchKind::BeginsWith => ascii_starts_with_case_insensitive(hay, needle),
            TextMatchKind::EndsWith => ascii_ends_with_case_insensitive(hay, needle),
        };
    }

    let cell_s = crate::value::casefold_owned(cell_to_string(cell, value_locale).into_owned());
    crate::value::with_casefolded_key(m.pattern.as_str(), |pattern_folded| match m.kind {
        TextMatchKind::Contains => cell_s.contains(pattern_folded),
        TextMatchKind::BeginsWith => cell_s.starts_with(pattern_folded),
        TextMatchKind::EndsWith => cell_s.ends_with(pattern_folded),
    })
}

fn ascii_contains_case_insensitive(haystack: &str, needle: &str) -> bool {
    if needle.is_empty() {
        return true;
    }
    if needle.len() > haystack.len() {
        return false;
    }
    for i in 0..=haystack.len() - needle.len() {
        if haystack[i..i + needle.len()].eq_ignore_ascii_case(needle) {
            return true;
        }
    }
    false
}

fn ascii_starts_with_case_insensitive(haystack: &str, needle: &str) -> bool {
    if needle.len() > haystack.len() {
        return false;
    }
    haystack[..needle.len()].eq_ignore_ascii_case(needle)
}

fn ascii_ends_with_case_insensitive(haystack: &str, needle: &str) -> bool {
    if needle.len() > haystack.len() {
        return false;
    }
    haystack[haystack.len() - needle.len()..].eq_ignore_ascii_case(needle)
}

fn number_cmp(cell: &CellValue, cmp: &NumberComparison, value_locale: ValueLocaleConfig) -> bool {
    let Some(n) = coerce_number(cell, value_locale) else {
        return false;
    };

    match cmp {
        NumberComparison::GreaterThan(v) => n > *v,
        NumberComparison::GreaterThanOrEqual(v) => n >= *v,
        NumberComparison::LessThan(v) => n < *v,
        NumberComparison::LessThanOrEqual(v) => n <= *v,
        NumberComparison::Between { min, max } => n >= *min && n <= *max,
        NumberComparison::NotEqual(v) => n != *v,
    }
}

fn date_cmp(cell: &CellValue, cmp: &DateComparison, value_locale: ValueLocaleConfig) -> bool {
    match cmp {
        DateComparison::Today => {
            let today = Local::now().date_naive();
            date_cmp(cell, &DateComparison::OnDate(today), value_locale)
        }
        DateComparison::Yesterday => {
            let today = Local::now().date_naive();
            let yesterday = today.pred_opt().unwrap_or(today);
            date_cmp(cell, &DateComparison::OnDate(yesterday), value_locale)
        }
        DateComparison::Tomorrow => {
            let today = Local::now().date_naive();
            let tomorrow = today.succ_opt().unwrap_or(today);
            date_cmp(cell, &DateComparison::OnDate(tomorrow), value_locale)
        }
        DateComparison::OnDate(d) => coerce_datetime(cell, value_locale)
            .map(|dt| dt.date() == *d)
            .unwrap_or(false),
        DateComparison::After(dt) => coerce_datetime(cell, value_locale).is_some_and(|v| v > *dt),
        DateComparison::Before(dt) => coerce_datetime(cell, value_locale).is_some_and(|v| v < *dt),
        DateComparison::Between { start, end } => {
            coerce_datetime(cell, value_locale).is_some_and(|v| v >= *start && v <= *end)
        }
    }
}

fn cell_to_string<'a>(cell: &'a CellValue, value_locale: ValueLocaleConfig) -> Cow<'a, str> {
    match cell {
        CellValue::Blank => Cow::Borrowed(""),
        CellValue::Number(n) => Cow::Owned(
            formula_format::format_value(
                FormatValue::Number(*n),
                None,
                &FormatOptions {
                    locale: value_locale.separators,
                    date_system: DateSystem::Excel1900,
                },
            )
            .text,
        ),
        CellValue::Text(s) => Cow::Borrowed(s),
        CellValue::Bool(true) => Cow::Borrowed("TRUE"),
        CellValue::Bool(false) => Cow::Borrowed("FALSE"),
        CellValue::Error(err) => Cow::Borrowed(err.as_str()),
        CellValue::DateTime(dt) => Cow::Owned(dt.format("%Y-%m-%d %H:%M:%S").to_string()),
    }
}

fn coerce_number(cell: &CellValue, value_locale: ValueLocaleConfig) -> Option<f64> {
    match cell {
        CellValue::Number(n) => Some(*n),
        CellValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::Text(s) => parse_text_number(s, value_locale),
        CellValue::DateTime(dt) => Some(datetime_to_excel_serial_1900(*dt)),
        CellValue::Error(_) | CellValue::Blank => None,
    }
}

fn coerce_datetime(cell: &CellValue, value_locale: ValueLocaleConfig) -> Option<NaiveDateTime> {
    match cell {
        CellValue::DateTime(dt) => Some(*dt),
        CellValue::Text(s) => parse_text_datetime(s, value_locale),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sort_filter::types::{CellValue, RangeData, RangeRef};
    use formula_model::ErrorValue;
    use pretty_assertions::assert_eq;

    fn range(rows: Vec<Vec<CellValue>>) -> RangeData {
        let range = RangeRef {
            start_row: 0,
            start_col: 0,
            end_row: rows.len() - 1,
            end_col: rows[0].len() - 1,
        };
        RangeData::new(range, rows).unwrap()
    }

    #[test]
    fn cell_to_string_formats_numbers_using_workbook_locale() {
        assert_eq!(
            cell_to_string(&CellValue::Number(1.5), ValueLocaleConfig::de_de()),
            "1,5"
        );
    }

    #[test]
    fn cell_to_string_formats_scientific_numbers_using_workbook_locale() {
        assert_eq!(
            cell_to_string(&CellValue::Number(1.23e11), ValueLocaleConfig::de_de()),
            "1,23E+11"
        );
    }

    #[test]
    fn cell_to_string_formats_numbers_with_excel_general_precision() {
        assert_eq!(
            cell_to_string(
                &CellValue::Number(1_234_567_890_123_456.0),
                ValueLocaleConfig::en_us()
            ),
            "1.23456789012346E+15"
        );
    }

    #[test]
    fn cell_to_string_normalizes_negative_zero() {
        assert_eq!(
            cell_to_string(&CellValue::Number(-0.0), ValueLocaleConfig::en_us()),
            "0"
        );
    }

    #[test]
    fn cell_to_string_formats_booleans_as_excel_true_false() {
        assert_eq!(
            cell_to_string(&CellValue::Bool(true), ValueLocaleConfig::en_us()),
            "TRUE"
        );
        assert_eq!(
            cell_to_string(&CellValue::Bool(false), ValueLocaleConfig::en_us()),
            "FALSE"
        );
    }

    #[test]
    fn text_contains_filter_hides_rows() {
        let data = range(vec![
            vec![CellValue::Text("Name".into())],
            vec![CellValue::Text("Alice".into())],
            vec![CellValue::Text("Bob".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::TextMatch(TextMatch {
                        kind: TextMatchKind::Contains,
                        pattern: "ali".into(),
                        case_sensitive: false,
                    })],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn numeric_between_filter() {
        let data = range(vec![
            vec![CellValue::Text("Score".into())],
            vec![CellValue::Number(10.0)],
            vec![CellValue::Number(5.0)],
            vec![CellValue::Number(1.0)],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Number(NumberComparison::Between {
                        min: 2.0,
                        max: 10.0,
                    })],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![3]);
    }

    #[test]
    fn locale_aware_numeric_filter_respects_decimal_separator() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("1,10".into())],
            vec![CellValue::Text("1,2".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Number(NumberComparison::LessThan(1.15))],
                },
            )]),
        };

        let result = apply_autofilter_with_value_locale(&data, &filter, ValueLocaleConfig::de_de())
            .expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn locale_aware_date_filter_parses_year_first_dot_separated_dates() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("2020.01.01".into())],
            vec![CellValue::Text("2020.01.02".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Date(DateComparison::OnDate(
                        NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
                    ))],
                },
            )]),
        };

        let result = apply_autofilter_with_value_locale(&data, &filter, ValueLocaleConfig::de_de())
            .expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn locale_aware_date_filter_parses_hyphen_separated_dmy_dates() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("31-01-2020".into())],
            vec![CellValue::Text("01-02-2020".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Date(DateComparison::OnDate(
                        NaiveDate::from_ymd_opt(2020, 1, 31).unwrap(),
                    ))],
                },
            )]),
        };

        let result = apply_autofilter_with_value_locale(&data, &filter, ValueLocaleConfig::de_de())
            .expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn locale_aware_date_filter_parses_ampm_datetime_strings() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("1/2/2020 2:00 PM".into())],
            vec![CellValue::Text("1/2/2020 2:00PM".into())],
            vec![CellValue::Text("1/3/2020 2:00 PM".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Date(DateComparison::OnDate(
                        NaiveDate::from_ymd_opt(2020, 1, 2).unwrap(),
                    ))],
                },
            )]),
        };

        let result = apply_autofilter_with_value_locale(&data, &filter, ValueLocaleConfig::en_us())
            .expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![3]);
    }

    #[test]
    fn equals_text_filter_is_unicode_case_insensitive() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("ω".into())],
            vec![CellValue::Text("x".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Equals(FilterValue::Text("Ω".into()))],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn blanks_filter() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Blank],
            vec![CellValue::Text("x".into())],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Blanks],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
    }

    #[test]
    fn blanks_filter_does_not_treat_errors_as_blank() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Error(ErrorValue::Div0)],
            vec![CellValue::Blank],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Blanks],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, false, true]);
        assert_eq!(result.hidden_sheet_rows, vec![1]);
    }

    #[test]
    fn numeric_filters_ignore_errors() {
        let data = range(vec![
            vec![CellValue::Text("Score".into())],
            vec![CellValue::Error(ErrorValue::Div0)],
            vec![CellValue::Number(10.0)],
            vec![CellValue::Number(5.0)],
        ]);

        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Number(NumberComparison::Between {
                        min: 6.0,
                        max: 10.0,
                    })],
                },
            )]),
        };

        let result = apply_autofilter(&data, &filter).expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, false, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![1, 3]);
    }

    #[test]
    fn filter_views_are_independent() {
        let data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("a".into())],
            vec![CellValue::Text("b".into())],
        ]);

        let filter_a = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Equals(FilterValue::Text("a".into()))],
                },
            )]),
        };

        let filter_b = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([(
                0,
                ColumnFilter {
                    join: FilterJoin::Any,
                    criteria: vec![FilterCriterion::Equals(FilterValue::Text("b".into()))],
                },
            )]),
        };

        let mut views = FilterViews::default();
        views.set_filter(FilterViewId("u1".into()), filter_a.clone());
        views.set_filter(FilterViewId("u2".into()), filter_b.clone());

        let result1 = apply_autofilter(
            &data,
            views
                .get_filter(&FilterViewId("u1".into()), data.range)
                .unwrap(),
        )
        .expect("filter should succeed");
        let result2 = apply_autofilter(
            &data,
            views
                .get_filter(&FilterViewId("u2".into()), data.range)
                .unwrap(),
        )
        .expect("filter should succeed");

        assert_eq!(result1.visible_rows, vec![true, true, false]);
        assert_eq!(result2.visible_rows, vec![true, false, true]);
    }

    #[test]
    fn locale_aware_text_filters_format_numbers_and_bools_like_excel() {
        let data = range(vec![
            vec![
                CellValue::Text("Val".into()),
                CellValue::Text("Flag".into()),
            ],
            vec![CellValue::Number(1.5), CellValue::Bool(true)],
        ]);
        let filter = AutoFilter {
            range: data.range,
            columns: BTreeMap::from([
                (
                    0,
                    ColumnFilter {
                        join: FilterJoin::Any,
                        criteria: vec![FilterCriterion::TextMatch(TextMatch {
                            kind: TextMatchKind::Contains,
                            pattern: "1,5".into(),
                            case_sensitive: false,
                        })],
                    },
                ),
                (
                    1,
                    ColumnFilter {
                        join: FilterJoin::Any,
                        criteria: vec![FilterCriterion::TextMatch(TextMatch {
                            kind: TextMatchKind::Contains,
                            pattern: "TRU".into(),
                            case_sensitive: false,
                        })],
                    },
                ),
            ]),
        };

        let result = apply_autofilter_with_value_locale(&data, &filter, ValueLocaleConfig::de_de())
            .expect("filter should succeed");
        assert_eq!(result.visible_rows, vec![true, true]);
        assert_eq!(result.hidden_sheet_rows, Vec::<usize>::new());

        // Ensure boolean coercion uses Excel-style TRUE/FALSE rather than Rust's "true"/"false".
        assert_eq!(
            cell_to_string(&CellValue::Bool(true), ValueLocaleConfig::en_us()),
            "TRUE"
        );
    }
}
