use crate::sort_filter::types::{CellValue, HeaderOption, RangeData};
use chrono::{NaiveDate, NaiveDateTime, Timelike};
use formula_model::ErrorValue;
use std::cmp::Ordering;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortOrder {
    Ascending,
    Descending,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortValueType {
    Auto,
    Text,
    Number,
    DateTime,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortKey {
    pub column: usize,
    pub order: SortOrder,
    pub value_type: SortValueType,
    pub case_sensitive: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortSpec {
    pub keys: Vec<SortKey>,
    pub header: HeaderOption,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RowPermutation {
    /// new_to_old[new_row] = old_row
    pub new_to_old: Vec<usize>,
    /// old_to_new[old_row] = new_row
    pub old_to_new: Vec<usize>,
}

#[derive(Debug, Clone)]
enum SortKeyValue {
    Blank,
    Number(f64),
    DateTime(NaiveDateTime),
    Text(String),
    Bool(bool),
    Error(ErrorValue),
}

impl SortKeyValue {
    fn kind_rank(&self) -> u8 {
        match self {
            SortKeyValue::Number(_) | SortKeyValue::DateTime(_) => 0,
            SortKeyValue::Text(_) => 1,
            SortKeyValue::Bool(_) => 2,
            SortKeyValue::Error(_) => 3,
            SortKeyValue::Blank => 4,
        }
    }

    fn as_number(&self) -> Option<f64> {
        match self {
            SortKeyValue::Number(v) => Some(*v),
            SortKeyValue::DateTime(dt) => Some(datetime_to_excel_serial_1900(*dt)),
            _ => None,
        }
    }
}

#[derive(Debug)]
struct SortRow {
    original_index: usize,
    key_values: Vec<SortKeyValue>,
}

pub fn sort_range(range: &mut RangeData, spec: &SortSpec) -> RowPermutation {
    let row_count = range.rows.len();
    if row_count <= 1 || spec.keys.is_empty() {
        return identity_permutation(row_count);
    }

    let header_rows = compute_header_rows(row_count, spec.header, &spec.keys, |row, col| {
        range
            .rows
            .get(row)
            .and_then(|r| r.get(col))
            .cloned()
            .unwrap_or(CellValue::Blank)
    });

    let perm = compute_row_permutation(row_count, header_rows, &spec.keys, |row, col| {
        range
            .rows
            .get(row)
            .and_then(|r| r.get(col))
            .cloned()
            .unwrap_or(CellValue::Blank)
    });

    range.rows = perm
        .new_to_old
        .iter()
        .copied()
        .map(|old_row| range.rows[old_row].clone())
        .collect();

    perm
}

fn compare_rows(a: &SortRow, b: &SortRow, keys: &[SortKey]) -> Ordering {
    for (key_index, key) in keys.iter().enumerate() {
        let ord = compare_key_value(&a.key_values[key_index], &b.key_values[key_index], key);
        if ord != Ordering::Equal {
            return ord;
        }
    }
    a.original_index.cmp(&b.original_index)
}

fn compare_key_value(a: &SortKeyValue, b: &SortKeyValue, key: &SortKey) -> Ordering {
    let rank_cmp = a.kind_rank().cmp(&b.kind_rank());
    if rank_cmp != Ordering::Equal {
        // Excel keeps a fixed cross-type ordering (numbers/dates, then text, then booleans, blanks
        // last) regardless of ascending/descending selection. The direction only applies within the
        // same type.
        return rank_cmp;
    }

    let ord = match (a, b) {
        (SortKeyValue::Blank, SortKeyValue::Blank) => Ordering::Equal,
        (SortKeyValue::Error(a), SortKeyValue::Error(b)) => a.code().cmp(&b.code()),
        (SortKeyValue::Text(a), SortKeyValue::Text(b)) => a.cmp(b),
        (SortKeyValue::Bool(a), SortKeyValue::Bool(b)) => a.cmp(b),
        _ => {
            let a = a.as_number().unwrap_or(f64::NAN);
            let b = b.as_number().unwrap_or(f64::NAN);
            a.total_cmp(&b)
        }
    };

    match key.order {
        SortOrder::Ascending => ord,
        SortOrder::Descending => ord.reverse(),
    }
}

pub(crate) fn compute_header_rows(
    row_count: usize,
    header: HeaderOption,
    keys: &[SortKey],
    mut cell_at: impl FnMut(usize, usize) -> CellValue,
) -> usize {
    match header {
        HeaderOption::None => 0,
        HeaderOption::HasHeader => 1.min(row_count),
        HeaderOption::Auto => detect_header_row(row_count, keys, &mut cell_at),
    }
}

pub(crate) fn compute_row_permutation(
    row_count: usize,
    header_rows: usize,
    keys: &[SortKey],
    mut cell_at: impl FnMut(usize, usize) -> CellValue,
) -> RowPermutation {
    if row_count == 0 {
        return RowPermutation {
            new_to_old: Vec::new(),
            old_to_new: Vec::new(),
        };
    }

    if row_count <= 1 || keys.is_empty() {
        return identity_permutation(row_count);
    }

    let mut sortable: Vec<SortRow> = (header_rows..row_count)
        .map(|row_index| SortRow {
            original_index: row_index,
            key_values: keys
                .iter()
                .map(|key| {
                    let cell = cell_at(row_index, key.column);
                    detect_key_value(&cell, key)
                })
                .collect(),
        })
        .collect();

    sortable.sort_by(|a, b| compare_rows(a, b, keys));

    let mut new_to_old: Vec<usize> = Vec::with_capacity(row_count);

    for row in 0..header_rows {
        new_to_old.push(row);
    }
    for entry in &sortable {
        new_to_old.push(entry.original_index);
    }

    let mut old_to_new = vec![0usize; row_count];
    for (new_index, old_index) in new_to_old.iter().copied().enumerate() {
        old_to_new[old_index] = new_index;
    }

    RowPermutation {
        new_to_old,
        old_to_new,
    }
}

fn detect_header_row(
    row_count: usize,
    keys: &[SortKey],
    cell_at: &mut dyn FnMut(usize, usize) -> CellValue,
) -> usize {
    if row_count < 2 {
        return 0;
    }

    // A conservative heuristic:
    // - If the first row contains text and the second row contains numbers/dates for any sort key,
    //   treat the first row as a header.
    // - Otherwise, assume no header.
    for key in keys {
        let v0 = cell_at(0, key.column);
        let v1 = cell_at(1, key.column);

        let is_text0 = matches!(v0, CellValue::Text(s) if !s.trim().is_empty());
        let is_number_or_date1 = matches!(v1, CellValue::Number(_) | CellValue::DateTime(_))
            || matches!(v1, CellValue::Text(s) if parse_number(&s).is_some() || parse_datetime(&s).is_some());

        if is_text0 && is_number_or_date1 {
            return 1;
        }
    }

    0
}

fn detect_key_value(cell: &CellValue, key: &SortKey) -> SortKeyValue {
    match cell {
        CellValue::Blank => return SortKeyValue::Blank,
        // Treat empty/whitespace-only text as blank, matching Excel AutoFilter "Blanks" semantics.
        // (See `sort_filter::filter::is_blank`.)
        CellValue::Text(s) if s.trim().is_empty() => return SortKeyValue::Blank,
        CellValue::Error(err) => return SortKeyValue::Error(*err),
        _ => {}
    }

    match key.value_type {
        SortValueType::Text => {
            SortKeyValue::Text(fold_text(cell_to_string(cell), key.case_sensitive))
        }
        SortValueType::Number => match coerce_number(cell) {
            Some(n) => SortKeyValue::Number(n),
            None => SortKeyValue::Text(fold_text(cell_to_string(cell), key.case_sensitive)),
        },
        SortValueType::DateTime => match coerce_datetime(cell) {
            Some(dt) => SortKeyValue::DateTime(dt),
            None => SortKeyValue::Text(fold_text(cell_to_string(cell), key.case_sensitive)),
        },
        SortValueType::Auto => {
            if let CellValue::Bool(b) = cell {
                return SortKeyValue::Bool(*b);
            }

            if let Some(n) = coerce_number(cell) {
                return SortKeyValue::Number(n);
            }
            if let Some(dt) = coerce_datetime(cell) {
                return SortKeyValue::DateTime(dt);
            }
            match cell {
                _ => SortKeyValue::Text(fold_text(cell_to_string(cell), key.case_sensitive)),
            }
        }
    }
}

fn fold_text(s: String, case_sensitive: bool) -> String {
    if case_sensitive {
        s
    } else {
        s.to_lowercase()
    }
}

fn cell_to_string(cell: &CellValue) -> String {
    match cell {
        CellValue::Blank => String::new(),
        CellValue::Number(n) => {
            // Excel doesn't use scientific notation when generating unique filter/sort items, but we
            // don't have format metadata yet. Keep it simple and deterministic.
            let mut s = n.to_string();
            if s.ends_with(".0") {
                s.truncate(s.len() - 2);
            }
            s
        }
        CellValue::Text(s) => s.clone(),
        CellValue::Bool(b) => b.to_string(),
        CellValue::Error(err) => err.to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
    }
}

fn parse_number(text: &str) -> Option<f64> {
    let mut s = text.trim().replace(',', "");
    if s.is_empty() {
        return None;
    }
    let negative = s.starts_with('(') && s.ends_with(')');
    if negative {
        s = s
            .trim_start_matches('(')
            .trim_end_matches(')')
            .trim()
            .to_string();
    }
    let s = s.strip_prefix('$').unwrap_or(&s);
    let n: f64 = s.parse().ok()?;
    Some(if negative { -n } else { n })
}

fn parse_datetime(text: &str) -> Option<NaiveDateTime> {
    let s = text.trim();
    if s.is_empty() {
        return None;
    }

    // Common ISO-ish formats first.
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M") {
        return Some(dt);
    }

    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }

    // US-style 1/31/2025, with optional time.
    if let Ok(date) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%m/%d/%Y %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%m/%d/%Y %H:%M") {
        return Some(dt);
    }

    None
}

fn coerce_number(cell: &CellValue) -> Option<f64> {
    match cell {
        CellValue::Number(n) => Some(*n),
        CellValue::Text(s) => parse_number(s),
        CellValue::DateTime(dt) => Some(datetime_to_excel_serial_1900(*dt)),
        CellValue::Bool(_) | CellValue::Error(_) | CellValue::Blank => None,
    }
}

fn coerce_datetime(cell: &CellValue) -> Option<NaiveDateTime> {
    match cell {
        CellValue::DateTime(dt) => Some(*dt),
        CellValue::Text(s) => parse_datetime(s),
        _ => None,
    }
}

pub(crate) fn datetime_to_excel_serial_1900(dt: NaiveDateTime) -> f64 {
    // Excel 1900 date system with the 1900 leap year bug.
    // See: https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487
    let base = NaiveDate::from_ymd_opt(1899, 12, 31).unwrap();
    let date = dt.date();
    let mut days = (date - base).num_days() as f64;
    if date >= NaiveDate::from_ymd_opt(1900, 3, 1).unwrap() {
        days += 1.0;
    }
    let seconds = dt.time().num_seconds_from_midnight() as f64;
    days + (seconds / 86_400.0)
}

fn identity_permutation(row_count: usize) -> RowPermutation {
    RowPermutation {
        new_to_old: (0..row_count).collect(),
        old_to_new: (0..row_count).collect(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sort_filter::types::{CellValue, RangeRef};
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
    fn stable_multi_key_sort_preserves_row_integrity() {
        let mut data = range(vec![
            vec![
                CellValue::Text("Name".into()),
                CellValue::Text("Score".into()),
            ],
            vec![CellValue::Text("Alice".into()), CellValue::Number(10.0)],
            vec![CellValue::Text("Bob".into()), CellValue::Number(10.0)],
            vec![CellValue::Text("Charlie".into()), CellValue::Number(7.0)],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![
                SortKey {
                    column: 1,
                    order: SortOrder::Descending,
                    value_type: SortValueType::Auto,
                    case_sensitive: false,
                },
                SortKey {
                    column: 0,
                    order: SortOrder::Ascending,
                    value_type: SortValueType::Auto,
                    case_sensitive: false,
                },
            ],
        };

        let perm = sort_range(&mut data, &spec);

        assert_eq!(
            data.rows
                .iter()
                .map(|r| match &r[0] {
                    CellValue::Text(s) => s.as_str(),
                    _ => "?",
                })
                .collect::<Vec<_>>(),
            vec!["Name", "Alice", "Bob", "Charlie"]
        );

        // Stable: Alice before Bob for equal Score and Name sort key.
        assert_eq!(perm.new_to_old, vec![0, 1, 2, 3]);
    }

    #[test]
    fn numeric_detection_sorts_numbers_not_lexicographically() {
        let mut data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("10".into())],
            vec![CellValue::Text("2".into())],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Ascending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(
            data.rows
                .iter()
                .skip(1)
                .map(|r| match &r[0] {
                    CellValue::Text(s) => s.as_str(),
                    _ => "?",
                })
                .collect::<Vec<_>>(),
            vec!["2", "10"]
        );
    }

    #[test]
    fn blanks_are_sorted_last_even_descending() {
        let mut data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Blank],
            vec![CellValue::Number(1.0)],
            vec![CellValue::Number(2.0)],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Descending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(data.rows[3][0], CellValue::Blank);
    }

    #[test]
    fn empty_text_is_sorted_like_blank() {
        let mut data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Text("".into())],
            vec![CellValue::Number(1.0)],
            vec![CellValue::Number(2.0)],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Descending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(data.rows[3][0], CellValue::Text("".into()));
    }

    #[test]
    fn auto_header_detection() {
        let mut data = range(vec![
            vec![CellValue::Text("Amount".into())],
            vec![CellValue::Number(10.0)],
            vec![CellValue::Number(2.0)],
        ]);

        let spec = SortSpec {
            header: HeaderOption::Auto,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Ascending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(data.rows[0][0], CellValue::Text("Amount".into()));
        assert_eq!(data.rows[1][0], CellValue::Number(2.0));
    }

    #[test]
    fn mixed_types_sort_places_errors_after_booleans() {
        let mut data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Bool(true)],
            vec![CellValue::Error(ErrorValue::Div0)],
            vec![CellValue::Bool(false)],
            vec![CellValue::Number(1.0)],
            vec![CellValue::Text("a".into())],
            vec![CellValue::Blank],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Ascending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(
            data.rows
                .iter()
                .skip(1)
                .map(|r| r[0].clone())
                .collect::<Vec<_>>(),
            vec![
                CellValue::Number(1.0),
                CellValue::Text("a".into()),
                CellValue::Bool(false),
                CellValue::Bool(true),
                CellValue::Error(ErrorValue::Div0),
                CellValue::Blank,
            ]
        );
    }

    #[test]
    fn extended_errors_sort_by_excel_error_code() {
        let mut data = range(vec![
            vec![CellValue::Text("Val".into())],
            vec![CellValue::Error(ErrorValue::Field)],
            vec![CellValue::Error(ErrorValue::GettingData)],
            vec![CellValue::Error(ErrorValue::Div0)],
        ]);

        let spec = SortSpec {
            header: HeaderOption::HasHeader,
            keys: vec![SortKey {
                column: 0,
                order: SortOrder::Ascending,
                value_type: SortValueType::Auto,
                case_sensitive: false,
            }],
        };

        sort_range(&mut data, &spec);

        assert_eq!(
            data.rows
                .iter()
                .skip(1)
                .map(|r| r[0].clone())
                .collect::<Vec<_>>(),
            vec![
                CellValue::Error(ErrorValue::Div0),
                CellValue::Error(ErrorValue::GettingData),
                CellValue::Error(ErrorValue::Field),
            ]
        );
    }
}
