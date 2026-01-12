use std::borrow::Cow;
use std::collections::HashMap;
use std::io::BufRead;
use std::sync::Arc;

use csv::ByteRecord;
use encoding_rs::WINDOWS_1252;
use formula_columnar::{
    ColumnSchema, ColumnType as ColumnarType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig,
    TableOptions, Value as ColumnarValue,
};
use thiserror::Error;

use crate::{CellRef, SheetNameError, Workbook, Worksheet, WorksheetId};

#[derive(Clone, Debug)]
pub struct CsvOptions {
    pub delimiter: u8,
    pub has_header: bool,
    pub sample_rows: usize,
    pub page_size_rows: usize,
    pub cache_entries: usize,
    /// How to decode raw CSV bytes into text fields.
    pub encoding: CsvTextEncoding,
    /// Decimal separator used when parsing numbers.
    ///
    /// `.` matches inputs like `1,234.56`. `,` matches inputs like `1.234,56`.
    pub decimal_separator: char,
    /// Preferred order for ambiguous numeric dates like `01/02/2024`.
    pub date_order: CsvDateOrder,
    /// How to handle timestamps with explicit timezone offsets (e.g. `2024-01-02T03:04:05-05:00`).
    pub timestamp_tz_policy: CsvTimestampTzPolicy,
    /// Currency symbols to recognize for currency inference/parsing.
    pub currency_symbols: Vec<char>,
}

impl Default for CsvOptions {
    fn default() -> Self {
        Self {
            delimiter: b',',
            has_header: true,
            sample_rows: 100,
            page_size_rows: 65_536,
            cache_entries: 64,
            encoding: CsvTextEncoding::Auto,
            decimal_separator: '.',
            date_order: CsvDateOrder::default(),
            timestamp_tz_policy: CsvTimestampTzPolicy::default(),
            currency_symbols: vec!['$', '€', '£', '¥'],
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CsvTextEncoding {
    /// Attempt to decode as UTF-8; if a field contains invalid UTF-8, fall back to Windows-1252.
    ///
    /// This matches common Excel behavior when opening CSV files on Windows.
    Auto,
    /// Decode as UTF-8 and reject invalid byte sequences.
    Utf8,
    /// Decode as Windows-1252 (aka CP-1252).
    Windows1252,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CsvDateOrder {
    /// Month / day / year (e.g. `12/31/2024`).
    Mdy,
    /// Day / month / year (e.g. `31/12/2024`).
    Dmy,
    /// Year / month / day (e.g. `2024/12/31`).
    Ymd,
}

impl Default for CsvDateOrder {
    fn default() -> Self {
        CsvDateOrder::Mdy
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CsvTimestampTzPolicy {
    /// Reject timestamps that include a timezone offset (accepts `Z` or no timezone).
    Reject,
    /// Ignore any explicit offset and interpret the clock time as UTC.
    IgnoreOffset,
    /// Convert timestamps with offsets to UTC.
    ConvertToUtc,
}

impl Default for CsvTimestampTzPolicy {
    fn default() -> Self {
        CsvTimestampTzPolicy::Reject
    }
}

#[derive(Debug, Error)]
pub enum CsvImportError {
    #[error("csv input was empty")]
    EmptyInput,
    #[error("csv parse error at row {row}, column {column}: {reason}")]
    Parse { row: u64, column: u64, reason: String },
    #[error(transparent)]
    SheetName(#[from] SheetNameError),
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

/// Import a CSV stream into a [`ColumnarTable`] without materializing a full grid.
pub fn import_csv_to_columnar_table<R: BufRead>(
    reader: R,
    options: CsvOptions,
) -> Result<ColumnarTable, CsvImportError> {
    let mut csv_reader = csv::ReaderBuilder::new()
        .delimiter(options.delimiter)
        // We'll treat headers manually so we can report consistent row/column locations.
        .has_headers(false)
        // Match prior behavior: accept rows with varying column counts.
        .flexible(true)
        .from_reader(reader);

    let mut record = ByteRecord::new();
    let mut record_index: u64 = 0;

    let has_first = csv_reader
        .read_byte_record(&mut record)
        .map_err(|e| map_csv_error(e, record_index + 1))?;
    if !has_first {
        return Err(CsvImportError::EmptyInput);
    }
    record_index += 1;

    let mut header_names: Vec<String> = Vec::new();
    let mut sample_rows: Vec<Vec<String>> = Vec::new();
    let mut column_count: usize;

    if options.has_header {
        header_names = decode_record_to_strings(&record, record_index, options.encoding)?;
        column_count = header_names.len();
    } else {
        let row = decode_record_to_strings(&record, record_index, options.encoding)?;
        column_count = row.len();
        sample_rows.push(row);
    }

    while sample_rows.len() < options.sample_rows {
        record.clear();
        match csv_reader.read_byte_record(&mut record) {
            Ok(false) => break,
            Ok(true) => {
                record_index += 1;
                let row = decode_record_to_strings(&record, record_index, options.encoding)?;
                column_count = column_count.max(row.len());
                sample_rows.push(row);
            }
            Err(e) => return Err(map_csv_error(e, record_index + 1)),
        }
    }

    // Match the previous line-based parser: an empty row implies a single empty field.
    if column_count == 0 {
        column_count = 1;
    }

    if options.has_header {
        if header_names.len() < column_count {
            header_names
                .extend((header_names.len()..column_count).map(|i| format!("Column{}", i + 1)));
        }
    } else {
        header_names = (0..column_count)
            .map(|i| format!("Column{}", i + 1))
            .collect();
    }

    for row in &mut sample_rows {
        if row.len() < column_count {
            row.resize(column_count, String::new());
        }
    }

    let column_types = infer_column_types(&sample_rows, column_count, &options);
    let schema: Vec<ColumnSchema> = header_names
        .into_iter()
        .zip(column_types.iter().copied())
        .map(|(name, column_type)| ColumnSchema { name, column_type })
        .collect();

    let mut builder = ColumnarTableBuilder::new(
        schema,
        TableOptions {
            page_size_rows: options.page_size_rows,
            cache: PageCacheConfig {
                max_entries: options.cache_entries,
            },
        },
    );

    let mut string_pool = StringPool::new();
    let mut row_values: Vec<ColumnarValue> = vec![ColumnarValue::Null; column_count];

    for row in &sample_rows {
        parse_row_to_values(row, &column_types, &options, &mut string_pool, &mut row_values);
        builder.append_row(&row_values);
    }

    // Stream the remainder.
    loop {
        record.clear();
        match csv_reader.read_byte_record(&mut record) {
            Ok(false) => break,
            Ok(true) => {
                record_index += 1;
                for i in 0..column_count {
                    let raw = record.get(i).unwrap_or(b"");
                    let field =
                        decode_field(raw, record_index, i as u64 + 1, options.encoding)?;
                    row_values[i] = parse_typed_value(
                        field.as_ref(),
                        column_types[i],
                        &options,
                        &mut string_pool,
                    );
                }

                builder.append_row(&row_values);
            }
            Err(e) => return Err(map_csv_error(e, record_index + 1)),
        }
    }

    Ok(builder.finalize())
}

/// Import a CSV stream into a new [`Worksheet`] backed by a columnar table.
pub fn import_csv_to_worksheet<R: BufRead>(
    sheet_id: WorksheetId,
    name: impl Into<String>,
    reader: R,
    options: CsvOptions,
) -> Result<Worksheet, CsvImportError> {
    let table = import_csv_to_columnar_table(reader, options)?;
    let mut sheet = Worksheet::new(sheet_id, name);
    sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));
    Ok(sheet)
}

/// Convenience: add a columnar-backed worksheet to an existing workbook.
pub fn import_csv_into_workbook<R: BufRead>(
    workbook: &mut Workbook,
    name: impl Into<String>,
    reader: R,
    options: CsvOptions,
) -> Result<WorksheetId, CsvImportError> {
    let name = name.into();
    let sheet_id = workbook.add_sheet(name.clone())?;
    let table = import_csv_to_columnar_table(reader, options)?;
    if let Some(sheet) = workbook.sheet_mut(sheet_id) {
        sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));
        sheet.name = name;
    }
    Ok(sheet_id)
}

fn parse_row_to_values(
    row: &[String],
    column_types: &[ColumnarType],
    options: &CsvOptions,
    string_pool: &mut StringPool,
    out: &mut [ColumnarValue],
) {
    for (i, column_type) in column_types.iter().copied().enumerate() {
        let field = row.get(i).map(|s| s.as_str()).unwrap_or("");
        out[i] = parse_typed_value(field, column_type, options, string_pool);
    }
}

fn parse_typed_value(
    field: &str,
    column_type: ColumnarType,
    options: &CsvOptions,
    string_pool: &mut StringPool,
) -> ColumnarValue {
    let v = field.trim();
    if v.is_empty() {
        return ColumnarValue::Null;
    }

    match column_type {
        ColumnarType::Number => parse_number_f64(v, options)
            .map(ColumnarValue::Number)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::String => ColumnarValue::String(string_pool.intern(v)),
        ColumnarType::Boolean => parse_bool(v)
            .map(ColumnarValue::Boolean)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::DateTime => parse_datetime_millis(v, options)
            .map(ColumnarValue::DateTime)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::Currency { scale } => parse_currency(v, scale, options)
            .map(ColumnarValue::Currency)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::Percentage { scale } => parse_percentage(v, scale, options)
            .map(ColumnarValue::Percentage)
            .unwrap_or(ColumnarValue::Null),
    }
}

fn parse_bool(v: &str) -> Option<bool> {
    match v.trim().to_ascii_lowercase().as_str() {
        "true" | "t" | "yes" | "y" | "1" => Some(true),
        "false" | "f" | "no" | "n" | "0" => Some(false),
        _ => None,
    }
}

fn parse_currency(v: &str, scale: u8, options: &CsvOptions) -> Option<i64> {
    let (sign, mut body) = split_sign_and_body(v);
    body = body.trim();

    let (stripped, found_symbol) = strip_symbol(body, &options.currency_symbols);
    if !found_symbol {
        return None;
    }

    let scaled = parse_fixed_point(stripped, scale, options)? as i128;
    Some(clamp_i128_to_i64(scaled.saturating_mul(sign as i128)))
}

fn parse_percentage(v: &str, scale: u8, options: &CsvOptions) -> Option<i64> {
    let (sign, body) = split_sign_and_body(v);
    let trimmed = body.trim();
    let num = trimmed.strip_suffix('%')?;
    // Stored as a fraction, not "percent points".
    let scaled = parse_fixed_point(num.trim(), scale, options)? as i128;
    Some(clamp_i128_to_i64(scaled.saturating_mul(sign as i128) / 100))
}

fn parse_fixed_point(v: &str, scale: u8, options: &CsvOptions) -> Option<i64> {
    let (sign, s) = split_sign_and_body(v);
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    let scale_factor: i64 = 10i64.checked_pow(scale as u32)?;

    let mut parts = s.splitn(2, options.decimal_separator);
    let int_part_raw = parts.next().unwrap_or("");
    let frac_part_raw = parts.next().unwrap_or("");

    let mut int_buf = String::new();
    for ch in int_part_raw.chars() {
        if ch.is_ascii_digit() {
            int_buf.push(ch);
        } else if is_grouping_separator(ch, options.decimal_separator) {
            continue;
        } else {
            return None;
        }
    }

    let int_part: i128 = if int_buf.is_empty() {
        0
    } else {
        int_buf.parse().ok()?
    };

    let mut frac_value: i128 = 0;
    let mut digits: u8 = 0;
    for ch in frac_part_raw.chars() {
        if digits >= scale {
            break;
        }
        if !ch.is_ascii_digit() {
            break;
        }
        frac_value = frac_value * 10 + (ch as i128 - '0' as i128);
        digits += 1;
    }

    while digits < scale {
        frac_value *= 10;
        digits += 1;
    }

    let scaled = int_part
        .saturating_mul(scale_factor as i128)
        .saturating_add(frac_value)
        .saturating_mul(sign as i128);
    Some(clamp_i128_to_i64(scaled))
}

fn parse_datetime_millis(v: &str, options: &CsvOptions) -> Option<i64> {
    let s = v.trim();
    let (year, month, day, rest) = parse_date_prefix(s, options.date_order)?;
    let days = days_from_civil(year, month, day)?;

    let mut millis: i64 = days * 86_400_000;
    let rest = rest.trim_start_matches(['T', ' ']);
    if rest.is_empty() {
        return Some(millis);
    }

    let (time_part, tz_part) = match rest.find(|c| c == 'Z' || c == '+' || c == '-') {
        Some(idx) => (&rest[..idx], &rest[idx..]),
        None => (rest, ""),
    };

    let time_part = time_part.trim();
    let (h, m, s_part) = if time_part.is_empty() {
        (0i64, 0i64, "0")
    } else {
        let mut t_iter = time_part.splitn(3, ':');
        let h: i64 = t_iter.next()?.parse().ok()?;
        let m: i64 = t_iter.next()?.parse().ok()?;
        let s_part = t_iter.next().unwrap_or("0").trim();
        (h, m, s_part)
    };

    let (sec_str, frac_str) = match s_part.split_once('.') {
        Some((a, b)) => (a, b),
        None => (s_part, ""),
    };
    let sec: i64 = sec_str.parse().ok()?;

    let mut frac_ms: i64 = 0;
    let mut frac_digits = 0;
    for ch in frac_str.chars() {
        if frac_digits >= 3 {
            break;
        }
        if !ch.is_ascii_digit() {
            break;
        }
        frac_ms = frac_ms * 10 + (ch as i64 - '0' as i64);
        frac_digits += 1;
    }
    while frac_digits < 3 {
        frac_ms *= 10;
        frac_digits += 1;
    }

    millis += ((h * 3600 + m * 60 + sec) * 1000) + frac_ms;

    if tz_part.is_empty() || tz_part.eq_ignore_ascii_case("Z") {
        return Some(millis);
    }

    let offset_millis = parse_tz_offset_millis(tz_part)?;
    match options.timestamp_tz_policy {
        CsvTimestampTzPolicy::Reject => None,
        CsvTimestampTzPolicy::IgnoreOffset => Some(millis),
        CsvTimestampTzPolicy::ConvertToUtc => Some(millis.saturating_sub(offset_millis)),
    }
}

fn days_from_civil(year: i32, month: u32, day: u32) -> Option<i64> {
    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }

    // Howard Hinnant's civil_from_days / days_from_civil algorithm, adapted.
    // Returns days relative to 1970-01-01.
    let y = year as i64 - if month <= 2 { 1 } else { 0 };
    let m = month as i64;
    let d = day as i64;

    let era = if y >= 0 { y } else { y - 399 }.div_euclid(400);
    let yoe = y - era * 400;
    let mp = m + if m > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    Some(era * 146_097 + doe - 719_468)
}

fn infer_column_types(
    sample_rows: &[Vec<String>],
    column_count: usize,
    options: &CsvOptions,
) -> Vec<ColumnarType> {
    let mut out = Vec::with_capacity(column_count);
    for col in 0..column_count {
        let mut is_bool = true;
        let mut saw_text_bool = false;
        let mut is_currency = true;
        let mut is_percentage = true;
        let mut is_datetime = true;
        let mut is_number = true;

        for row in sample_rows {
            let v = row.get(col).map(|s| s.trim()).unwrap_or("");
            if v.is_empty() {
                continue;
            }
            match parse_bool(v) {
                Some(_) => {
                    let lowered = v.trim().to_ascii_lowercase();
                    if lowered != "0" && lowered != "1" {
                        saw_text_bool = true;
                    }
                }
                None => is_bool = false,
            }
            if parse_currency(v, 2, options).is_none() {
                is_currency = false;
            }
            if parse_percentage(v, 4, options).is_none() {
                is_percentage = false;
            }
            if parse_datetime_millis(v, options).is_none() {
                is_datetime = false;
            }
            if parse_number_f64(v, options).is_none() {
                is_number = false;
            }
        }

        let ty = if is_bool && saw_text_bool {
            ColumnarType::Boolean
        } else if is_currency {
            ColumnarType::Currency { scale: 2 }
        } else if is_percentage {
            ColumnarType::Percentage { scale: 4 }
        } else if is_datetime {
            ColumnarType::DateTime
        } else if is_number {
            ColumnarType::Number
        } else {
            ColumnarType::String
        };
        out.push(ty);
    }
    out
}

fn parse_number_f64(v: &str, options: &CsvOptions) -> Option<f64> {
    let (sign, body) = split_sign_and_body(v);
    let body = body.trim();
    if body.is_empty() {
        return None;
    }

    let normalized = normalize_number(body, options.decimal_separator)?;
    let parsed: f64 = normalized.parse().ok()?;
    Some(parsed * sign as f64)
}

fn normalize_number(s: &str, decimal_separator: char) -> Option<String> {
    let mut out = String::with_capacity(s.len());
    let mut saw_digit = false;
    let mut saw_decimal = false;
    let mut saw_exp = false;
    let mut saw_exp_sign = false;

    for ch in s.chars() {
        if ch.is_ascii_digit() {
            saw_digit = true;
            out.push(ch);
            continue;
        }

        if !saw_exp && ch == decimal_separator {
            if saw_decimal {
                return None;
            }
            saw_decimal = true;
            out.push('.');
            continue;
        }

        if is_grouping_separator(ch, decimal_separator) {
            continue;
        }

        if !saw_exp && matches!(ch, 'e' | 'E') {
            if !saw_digit {
                return None;
            }
            saw_exp = true;
            saw_exp_sign = false;
            out.push('e');
            continue;
        }

        if saw_exp && !saw_exp_sign && matches!(ch, '+' | '-') {
            // Exponent sign is only valid immediately after `e` / `E`.
            if out.ends_with('e') {
                saw_exp_sign = true;
                out.push(ch);
                continue;
            }
            return None;
        }

        return None;
    }

    if !saw_digit {
        return None;
    }
    if out.ends_with('e') || out.ends_with("e+") || out.ends_with("e-") {
        return None;
    }
    Some(out)
}

fn split_sign_and_body(mut s: &str) -> (i64, &str) {
    s = s.trim();
    let mut sign: i64 = 1;

    if let Some(inner) = s.strip_prefix('(').and_then(|rest| rest.strip_suffix(')')) {
        sign = -1;
        s = inner.trim();
    }

    if let Some(rest) = s.strip_prefix('-') {
        sign = -sign;
        s = rest.trim_start();
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest.trim_start();
    }

    (sign, s)
}

fn strip_symbol<'a>(mut s: &'a str, symbols: &[char]) -> (&'a str, bool) {
    s = s.trim();
    for sym in symbols {
        if let Some(rest) = s.strip_prefix(*sym) {
            return (rest.trim_start(), true);
        }
        if let Some(rest) = s.strip_suffix(*sym) {
            return (rest.trim_end(), true);
        }
    }
    (s, false)
}

fn is_grouping_separator(ch: char, decimal_separator: char) -> bool {
    match ch {
        ',' => decimal_separator != ',',
        '.' => decimal_separator != '.',
        // Common grouping separators across locales.
        ' ' | '\u{00A0}' | '\u{202F}' | '_' | '\'' | '’' => true,
        _ => false,
    }
}

fn clamp_i128_to_i64(v: i128) -> i64 {
    if v > i64::MAX as i128 {
        i64::MAX
    } else if v < i64::MIN as i128 {
        i64::MIN
    } else {
        v as i64
    }
}

fn parse_date_prefix(s: &str, date_order: CsvDateOrder) -> Option<(i32, u32, u32, &str)> {
    let bytes = s.as_bytes();

    let mut date_end = 0usize;
    for (idx, b) in bytes.iter().enumerate() {
        if b.is_ascii_digit() || *b == b'-' || *b == b'/' {
            date_end = idx + 1;
        } else {
            break;
        }
    }

    if date_end == 0 {
        return None;
    }

    // Safety: date_part is ASCII only.
    let date_part = &s[..date_end];
    let rest = &s[date_end..];

    // YYYYMMDD.
    if date_part.len() >= 8 && date_part.as_bytes()[..8].iter().all(|b| b.is_ascii_digit()) {
        let year: i32 = date_part.get(0..4)?.parse().ok()?;
        let month: u32 = date_part.get(4..6)?.parse().ok()?;
        let day: u32 = date_part.get(6..8)?.parse().ok()?;
        let rest = &s[8..];
        return Some((year, month, day, rest));
    }

    let parts: Vec<&str> = date_part.split(|c| c == '-' || c == '/').collect();
    if parts.len() != 3 {
        return None;
    }

    if parts[0].len() == 4 {
        let year: i32 = parts[0].parse().ok()?;
        let month: u32 = parts[1].parse().ok()?;
        let day: u32 = parts[2].parse().ok()?;
        return Some((year, month, day, rest));
    }

    if parts[2].len() == 4 {
        let year: i32 = parts[2].parse().ok()?;
        let a: u32 = parts[0].parse().ok()?;
        let b: u32 = parts[1].parse().ok()?;

        let (month, day) = if a > 12 && b <= 12 {
            (b, a)
        } else if b > 12 && a <= 12 {
            (a, b)
        } else {
            match date_order {
                CsvDateOrder::Dmy => (b, a),
                CsvDateOrder::Mdy | CsvDateOrder::Ymd => (a, b),
            }
        };
        return Some((year, month, day, rest));
    }

    None
}

fn parse_tz_offset_millis(tz: &str) -> Option<i64> {
    let tz = tz.trim();
    let (sign, rest) = match tz.as_bytes().first().copied() {
        Some(b'+') => (1i64, &tz[1..]),
        Some(b'-') => (-1i64, &tz[1..]),
        _ => return None,
    };

    let (hours_str, mins_str) = if let Some((h, m)) = rest.split_once(':') {
        (h, m)
    } else if rest.len() == 4 {
        (&rest[..2], &rest[2..])
    } else if rest.len() == 2 {
        (rest, "0")
    } else {
        return None;
    };

    let hours: i64 = hours_str.parse().ok()?;
    let mins: i64 = mins_str.parse().ok()?;
    if hours.abs() > 23 || mins.abs() > 59 {
        return None;
    }
    Some(sign * ((hours * 3600 + mins * 60) * 1000))
}

fn decode_record_to_strings(
    record: &ByteRecord,
    row: u64,
    encoding: CsvTextEncoding,
) -> Result<Vec<String>, CsvImportError> {
    if record.len() == 0 {
        return Ok(vec![String::new()]);
    }

    let mut out = Vec::with_capacity(record.len());
    for (idx, field) in record.iter().enumerate() {
        let s = decode_field(field, row, idx as u64 + 1, encoding)?;
        out.push(s.into_owned());
    }
    Ok(out)
}

fn decode_field<'a>(
    field: &'a [u8],
    row: u64,
    column: u64,
    encoding: CsvTextEncoding,
) -> Result<Cow<'a, str>, CsvImportError> {
    // Handle UTF-8 BOM at the start of the file. This commonly appears in Excel-exported CSVs.
    let field = if row == 1 && column == 1 && field.starts_with(&[0xEF, 0xBB, 0xBF]) {
        &field[3..]
    } else {
        field
    };

    match encoding {
        CsvTextEncoding::Utf8 => std::str::from_utf8(field)
            .map(Cow::Borrowed)
            .map_err(|e| CsvImportError::Parse {
                row,
                column,
                reason: format!("invalid UTF-8: {e}"),
            }),
        CsvTextEncoding::Windows1252 => {
            let (cow, _, _) = WINDOWS_1252.decode(field);
            Ok(cow)
        }
        CsvTextEncoding::Auto => match std::str::from_utf8(field) {
            Ok(s) => Ok(Cow::Borrowed(s)),
            Err(_) => {
                let (cow, _, _) = WINDOWS_1252.decode(field);
                Ok(cow)
            }
        },
    }
}

fn map_csv_error(err: csv::Error, fallback_row: u64) -> CsvImportError {
    let reason = err.to_string();
    let pos = err.position().cloned();

    match err.into_kind() {
        csv::ErrorKind::Io(e) => CsvImportError::Io(e),
        _ => {
            let row = pos
                .map(|p| p.record())
                .filter(|r| *r > 0)
                .unwrap_or(fallback_row);
            CsvImportError::Parse {
                row,
                column: 0,
                reason,
            }
        }
    }
}

struct StringPool {
    set: HashMap<Arc<str>, ()>,
}

impl StringPool {
    fn new() -> Self {
        Self { set: HashMap::new() }
    }

    fn intern(&mut self, s: &str) -> Arc<str> {
        if let Some((k, _)) = self.set.get_key_value(s) {
            return k.clone();
        }

        let arc: Arc<str> = Arc::<str>::from(s);
        self.set.insert(arc.clone(), ());
        arc
    }
}
