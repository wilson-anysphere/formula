//! Helpers for building pivot cache source ranges from worksheet data.
//!
//! The core pivot engine (`pivot/mod.rs`) operates on a cached rectangular dataset:
//! a header row plus N records, all represented as [`PivotValue`].
//!
//! When building that dataset from a live worksheet range, we need access to cell
//! number formats so we can infer when numeric values represent Excel serial
//! dates/times. Excel stores dates as numbers + number format; without this step
//! date-like fields become plain numbers and cannot be grouped/sorted correctly.

use chrono::NaiveDate;

use crate::date::{serial_to_ymd, ExcelDateSystem};

use super::PivotValue;

/// Build a `Vec<Vec<PivotValue>>` pivot source range (header + records) from callers that can
/// provide both values and number formats.
///
/// For each cell:
/// - If the value is a [`PivotValue::Number`] and the resolved number format looks like a
///   date/time format, it is converted to [`PivotValue::Date`].
/// - Otherwise the value is returned unchanged.
///
/// # Time values (fractional serials)
///
/// Excel stores time-of-day as the fractional part of the serial date number. The pivot engine
/// currently only supports a date type (`NaiveDate`), so this helper **drops the fractional
/// component** and uses only the date component (`floor(serial)`).
pub fn build_pivot_source_range_with_number_formats<'a, ValueAt, NumFmtAt>(
    row_count: usize,
    col_count: usize,
    mut value_at: ValueAt,
    mut number_format_at: NumFmtAt,
    date_system: ExcelDateSystem,
) -> Vec<Vec<PivotValue>>
where
    ValueAt: FnMut(usize, usize) -> PivotValue,
    NumFmtAt: FnMut(usize, usize) -> Option<&'a str>,
{
    let mut out: Vec<Vec<PivotValue>> = Vec::new();
    if out.try_reserve_exact(row_count).is_err() {
        debug_assert!(
            false,
            "allocation failed (pivot source rows={row_count}, cols={col_count})"
        );
        return Vec::new();
    }

    for r in 0..row_count {
        let mut row: Vec<PivotValue> = Vec::new();
        if row.try_reserve_exact(col_count).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot source row cols={col_count})"
            );
            return Vec::new();
        }
        for c in 0..col_count {
            let value = value_at(r, c);
            let value =
                coerce_pivot_value_with_number_format(value, number_format_at(r, c), date_system);
            row.push(value);
        }
        out.push(row);
    }

    out
}

/// Coerce a single pivot value using an Excel number format code.
pub fn coerce_pivot_value_with_number_format(
    value: PivotValue,
    number_format: Option<&str>,
    date_system: ExcelDateSystem,
) -> PivotValue {
    let PivotValue::Number(n) = value else {
        return value;
    };

    let Some(fmt) = number_format else {
        return PivotValue::Number(n);
    };

    if !number_format_looks_like_datetime(fmt) {
        return PivotValue::Number(n);
    }

    match serial_number_to_naive_date(n, date_system) {
        Some(date) => PivotValue::Date(date),
        None => PivotValue::Number(n),
    }
}

/// Resolve the number format code for a `style_id` (if any).
///
/// This is a convenience helper for callers that store formats indirectly via the
/// workbook [`formula_model::StyleTable`].
pub fn resolve_number_format_from_style_id<'a>(
    styles: &'a formula_model::StyleTable,
    style_id: u32,
) -> Option<&'a str> {
    styles
        .get(style_id)
        .and_then(|s| s.number_format.as_deref())
}

fn number_format_looks_like_datetime(code: &str) -> bool {
    let code = code.trim();
    if code.is_empty() {
        return false;
    }

    let resolved = resolve_builtin_placeholder(code);
    // Excel formats can have up to 4 `;`-separated sections. We treat it as a datetime format if
    // **any** section contains date/time tokens.
    resolved
        .split(';')
        .any(|section| formula_format::datetime::looks_like_datetime(section))
}

fn resolve_builtin_placeholder(code: &str) -> &str {
    let Some(rest) = code.strip_prefix(formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)
    else {
        return code;
    };

    let id = rest.trim().parse::<u16>().ok();
    id.and_then(formula_format::builtin_format_code)
        .unwrap_or(code)
}

fn serial_number_to_naive_date(serial: f64, system: ExcelDateSystem) -> Option<NaiveDate> {
    if !serial.is_finite() {
        return None;
    }
    if serial < 0.0 {
        return None;
    }

    // The pivot engine only supports dates (not datetimes), so use the date component only.
    let serial = serial.floor();

    if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
        return None;
    }
    let serial_i32 = serial as i32;

    let excel_date = serial_to_ymd(serial_i32, system).ok()?;
    NaiveDate::from_ymd_opt(
        excel_date.year,
        excel_date.month as u32,
        excel_date.day as u32,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    use crate::date::{ymd_to_serial, ExcelDate};
    use pretty_assertions::assert_eq;

    #[test]
    fn converts_numeric_cells_with_date_formats_to_pivot_dates() {
        let date_serial =
            ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

        let values = |r: usize, c: usize| match (r, c) {
            (0, 0) => PivotValue::Text("Date".to_string()),
            (0, 1) => PivotValue::Text("Amount".to_string()),
            (1, 0) => PivotValue::Number(date_serial),
            (1, 1) => PivotValue::Number(123.0),
            _ => PivotValue::Blank,
        };

        let date_fmt: &'static str = "m/d/yyyy";
        let formats = move |r: usize, c: usize| match (r, c) {
            (1, 0) => Some(date_fmt),
            _ => None,
        };

        let built = build_pivot_source_range_with_number_formats(
            2,
            2,
            values,
            formats,
            ExcelDateSystem::EXCEL_1900,
        );

        assert_eq!(built.len(), 2);
        assert_eq!(built[0].len(), 2);
        assert_eq!(built[0][0], PivotValue::Text("Date".to_string()));
        assert_eq!(built[0][1], PivotValue::Text("Amount".to_string()));

        assert_eq!(
            built[1][0],
            PivotValue::Date(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap())
        );
        assert_eq!(built[1][1], PivotValue::Number(123.0));
    }
}
