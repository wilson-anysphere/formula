use crate::sort_filter::sort::datetime_to_excel_serial_1900;
use crate::sort_filter::types::{CellValue, RangeData, RangeRef};
use chrono::{Local, NaiveDate, NaiveDateTime};
use std::collections::{BTreeMap, HashMap};

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
    Between { start: NaiveDateTime, end: NaiveDateTime },
    OnDate(NaiveDate),
    Today,
    Yesterday,
    Tomorrow,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FilterCriterion {
    Equals(FilterValue),
    TextMatch(TextMatch),
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

#[derive(Debug, Clone, PartialEq)]
pub struct FilterResult {
    /// For each row in the range, `true` if the row is visible.
    pub visible_rows: Vec<bool>,
    /// Sheet row indices that should be hidden.
    pub hidden_sheet_rows: Vec<usize>,
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

pub fn apply_autofilter(range: &RangeData, filter: &AutoFilter) -> FilterResult {
    let row_count = range.rows.len();
    let mut visible_rows = vec![true; row_count];

    // Excel AutoFilter always treats the first row in the ref as the header row.
    if row_count == 0 {
        return FilterResult {
            visible_rows,
            hidden_sheet_rows: Vec::new(),
        };
    }

    for local_row in 1..row_count {
        let mut row_visible = true;
        for (col_id, col_filter) in &filter.columns {
            let cell = range
                .rows
                .get(local_row)
                .and_then(|r| r.get(*col_id))
                .unwrap_or(&CellValue::Blank);
            if !evaluate_column_filter(cell, col_filter) {
                row_visible = false;
                break;
            }
        }
        visible_rows[local_row] = row_visible;
    }

    let hidden_sheet_rows = visible_rows
        .iter()
        .enumerate()
        .filter_map(|(local_row, visible)| {
            if local_row == 0 || *visible {
                None
            } else {
                Some(range.range.start_row + local_row)
            }
        })
        .collect::<Vec<_>>();

    FilterResult {
        visible_rows,
        hidden_sheet_rows,
    }
}

fn evaluate_column_filter(cell: &CellValue, filter: &ColumnFilter) -> bool {
    if filter.criteria.is_empty() {
        return true;
    }

    match filter.join {
        FilterJoin::Any => filter.criteria.iter().any(|c| evaluate_criterion(cell, c)),
        FilterJoin::All => filter.criteria.iter().all(|c| evaluate_criterion(cell, c)),
    }
}

fn evaluate_criterion(cell: &CellValue, criterion: &FilterCriterion) -> bool {
    match criterion {
        FilterCriterion::Blanks => is_blank(cell),
        FilterCriterion::NonBlanks => !is_blank(cell),
        FilterCriterion::Equals(value) => equals_value(cell, value),
        FilterCriterion::TextMatch(m) => text_match(cell, m),
        FilterCriterion::Number(cmp) => number_cmp(cell, cmp),
        FilterCriterion::Date(cmp) => date_cmp(cell, cmp),
    }
}

fn is_blank(cell: &CellValue) -> bool {
    matches!(cell, CellValue::Blank) || matches!(cell, CellValue::Text(s) if s.trim().is_empty())
}

fn equals_value(cell: &CellValue, value: &FilterValue) -> bool {
    match value {
        FilterValue::Text(s) => {
            let cell_s = cell_to_string(cell);
            cell_s.eq_ignore_ascii_case(s)
        }
        FilterValue::Number(n) => coerce_number(cell).is_some_and(|v| v == *n),
        FilterValue::Bool(b) => matches!(cell, CellValue::Bool(v) if v == b),
        FilterValue::DateTime(dt) => coerce_datetime(cell).is_some_and(|v| v == *dt),
    }
}

fn text_match(cell: &CellValue, m: &TextMatch) -> bool {
    let mut cell_s = cell_to_string(cell);
    let mut pattern = m.pattern.clone();

    if !m.case_sensitive {
        cell_s = cell_s.to_lowercase();
        pattern = pattern.to_lowercase();
    }

    match m.kind {
        TextMatchKind::Contains => cell_s.contains(&pattern),
        TextMatchKind::BeginsWith => cell_s.starts_with(&pattern),
        TextMatchKind::EndsWith => cell_s.ends_with(&pattern),
    }
}

fn number_cmp(cell: &CellValue, cmp: &NumberComparison) -> bool {
    let Some(n) = coerce_number(cell) else {
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

fn date_cmp(cell: &CellValue, cmp: &DateComparison) -> bool {
    match cmp {
        DateComparison::Today => {
            let today = Local::now().date_naive();
            date_cmp(cell, &DateComparison::OnDate(today))
        }
        DateComparison::Yesterday => {
            let yesterday = Local::now().date_naive().pred_opt().unwrap();
            date_cmp(cell, &DateComparison::OnDate(yesterday))
        }
        DateComparison::Tomorrow => {
            let tomorrow = Local::now().date_naive().succ_opt().unwrap();
            date_cmp(cell, &DateComparison::OnDate(tomorrow))
        }
        DateComparison::OnDate(d) => coerce_datetime(cell)
            .map(|dt| dt.date() == *d)
            .unwrap_or(false),
        DateComparison::After(dt) => coerce_datetime(cell).is_some_and(|v| v > *dt),
        DateComparison::Before(dt) => coerce_datetime(cell).is_some_and(|v| v < *dt),
        DateComparison::Between { start, end } => coerce_datetime(cell)
            .is_some_and(|v| v >= *start && v <= *end),
    }
}

fn cell_to_string(cell: &CellValue) -> String {
    match cell {
        CellValue::Blank => String::new(),
        CellValue::Number(n) => n.to_string(),
        CellValue::Text(s) => s.clone(),
        CellValue::Bool(b) => b.to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
    }
}

fn parse_number(text: &str) -> Option<f64> {
    let s = text.trim().replace(',', "");
    if s.is_empty() {
        return None;
    }
    s.parse().ok()
}

fn parse_datetime(text: &str) -> Option<NaiveDateTime> {
    let s = text.trim();
    if s.is_empty() {
        return None;
    }

    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M") {
        return Some(dt);
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }
    None
}

fn coerce_number(cell: &CellValue) -> Option<f64> {
    match cell {
        CellValue::Number(n) => Some(*n),
        CellValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::Text(s) => parse_number(s),
        CellValue::DateTime(dt) => Some(datetime_to_excel_serial_1900(*dt)),
        CellValue::Blank => None,
    }
}

fn coerce_datetime(cell: &CellValue) -> Option<NaiveDateTime> {
    match cell {
        CellValue::DateTime(dt) => Some(*dt),
        CellValue::Text(s) => parse_datetime(s),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sort_filter::types::{CellValue, RangeData, RangeRef};
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

        let result = apply_autofilter(&data, &filter);
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

        let result = apply_autofilter(&data, &filter);
        assert_eq!(result.visible_rows, vec![true, true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![3]);
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

        let result = apply_autofilter(&data, &filter);
        assert_eq!(result.visible_rows, vec![true, true, false]);
        assert_eq!(result.hidden_sheet_rows, vec![2]);
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

        let result1 = apply_autofilter(&data, views.get_filter(&FilterViewId("u1".into()), data.range).unwrap());
        let result2 = apply_autofilter(&data, views.get_filter(&FilterViewId("u2".into()), data.range).unwrap());

        assert_eq!(result1.visible_rows, vec![true, true, false]);
        assert_eq!(result2.visible_rows, vec![true, false, true]);
    }
}
