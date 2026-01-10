use std::collections::HashMap;
use std::io::BufRead;
use std::sync::Arc;

use formula_columnar::{
    ColumnSchema, ColumnType as ColumnarType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig,
    TableOptions, Value as ColumnarValue,
};
use thiserror::Error;

use crate::{CellRef, Workbook, Worksheet, WorksheetId};

#[derive(Clone, Debug)]
pub struct CsvOptions {
    pub delimiter: u8,
    pub has_header: bool,
    pub sample_rows: usize,
    pub page_size_rows: usize,
    pub cache_entries: usize,
}

impl Default for CsvOptions {
    fn default() -> Self {
        Self {
            delimiter: b',',
            has_header: true,
            sample_rows: 100,
            page_size_rows: 65_536,
            cache_entries: 64,
        }
    }
}

#[derive(Debug, Error)]
pub enum CsvImportError {
    #[error("csv input was empty")]
    EmptyInput,
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

/// Import a CSV stream into a [`ColumnarTable`] without materializing a full grid.
pub fn import_csv_to_columnar_table<R: BufRead>(
    mut reader: R,
    options: CsvOptions,
) -> Result<ColumnarTable, CsvImportError> {
    let mut line = String::new();
    let mut fields: Vec<CsvField> = Vec::new();

    let bytes = reader.read_line(&mut line)?;
    if bytes == 0 {
        return Err(CsvImportError::EmptyInput);
    }
    trim_newline(&mut line);
    parse_csv_fields(&line, options.delimiter, &mut fields);

    let mut header_names: Vec<String>;
    let mut sample_rows: Vec<Vec<String>> = Vec::new();
    let mut column_count = fields.len();

    if options.has_header {
        header_names = fields
            .iter()
            .map(|c| c.as_str(&line).to_string())
            .collect();
    } else {
        sample_rows.push(
            fields
                .iter()
                .map(|c| c.as_str(&line).to_string())
                .collect(),
        );
        header_names = (0..column_count)
            .map(|i| format!("Column{}", i + 1))
            .collect();
    }
    fields.clear();

    line.clear();
    for _ in 0..options.sample_rows {
        let bytes = reader.read_line(&mut line)?;
        if bytes == 0 {
            break;
        }
        trim_newline(&mut line);
        parse_csv_fields(&line, options.delimiter, &mut fields);
        column_count = column_count.max(fields.len());
        sample_rows.push(
            fields
                .iter()
                .map(|c| c.as_str(&line).to_string())
                .collect(),
        );
        fields.clear();
        line.clear();
    }

    if header_names.len() < column_count {
        header_names
            .extend((header_names.len()..column_count).map(|i| format!("Column{}", i + 1)));
    }

    for row in &mut sample_rows {
        if row.len() < column_count {
            row.resize(column_count, String::new());
        }
    }

    let column_types = infer_column_types(&sample_rows, column_count);
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
        parse_row_to_values(row, &column_types, &mut string_pool, &mut row_values);
        builder.append_row(&row_values);
    }

    // Stream the remainder.
    fields.clear();
    line.clear();
    while reader.read_line(&mut line)? != 0 {
        trim_newline(&mut line);
        parse_csv_fields(&line, options.delimiter, &mut fields);

        for i in 0..column_count {
            let field = fields.get(i).map(|c| c.as_str(&line)).unwrap_or("");
            row_values[i] = parse_typed_value(field, column_types[i], &mut string_pool);
        }

        builder.append_row(&row_values);
        fields.clear();
        line.clear();
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
    let sheet_id = workbook.add_sheet(name.clone());
    let table = import_csv_to_columnar_table(reader, options)?;
    if let Some(sheet) = workbook.sheet_mut(sheet_id) {
        sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));
        sheet.name = name;
    }
    Ok(sheet_id)
}

fn infer_column_types(sample_rows: &[Vec<String>], column_count: usize) -> Vec<ColumnarType> {
    let mut out = Vec::with_capacity(column_count);
    for col in 0..column_count {
        let mut is_bool = true;
        let mut is_currency = true;
        let mut is_percentage = true;
        let mut is_datetime = true;
        let mut is_number = true;

        for row in sample_rows {
            let v = row.get(col).map(|s| s.trim()).unwrap_or("");
            if v.is_empty() {
                continue;
            }
            if parse_bool(v).is_none() {
                is_bool = false;
            }
            if parse_currency(v, 2).is_none() {
                is_currency = false;
            }
            if parse_percentage(v, 4).is_none() {
                is_percentage = false;
            }
            if parse_datetime_millis(v).is_none() {
                is_datetime = false;
            }
            if v.parse::<f64>().is_err() {
                is_number = false;
            }
        }

        let ty = if is_bool {
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

fn parse_row_to_values(
    row: &[String],
    column_types: &[ColumnarType],
    string_pool: &mut StringPool,
    out: &mut [ColumnarValue],
) {
    for (i, column_type) in column_types.iter().copied().enumerate() {
        let field = row.get(i).map(|s| s.as_str()).unwrap_or("");
        out[i] = parse_typed_value(field, column_type, string_pool);
    }
}

fn parse_typed_value(field: &str, column_type: ColumnarType, string_pool: &mut StringPool) -> ColumnarValue {
    let v = field.trim();
    if v.is_empty() {
        return ColumnarValue::Null;
    }

    match column_type {
        ColumnarType::Number => v
            .parse::<f64>()
            .ok()
            .map(ColumnarValue::Number)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::String => ColumnarValue::String(string_pool.intern(v)),
        ColumnarType::Boolean => parse_bool(v)
            .map(ColumnarValue::Boolean)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::DateTime => parse_datetime_millis(v)
            .map(ColumnarValue::DateTime)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::Currency { scale } => parse_currency(v, scale)
            .map(ColumnarValue::Currency)
            .unwrap_or(ColumnarValue::Null),
        ColumnarType::Percentage { scale } => parse_percentage(v, scale)
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

fn parse_currency(v: &str, scale: u8) -> Option<i64> {
    let trimmed = v.trim();
    let mut s = trimmed;
    let mut sign: i64 = 1;
    if let Some(rest) = s.strip_prefix('-') {
        sign = -1;
        s = rest;
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest;
    }

    let num = s.strip_prefix('$')?;
    parse_fixed_point(num, scale).map(|v| v.saturating_mul(sign))
}

fn parse_percentage(v: &str, scale: u8) -> Option<i64> {
    let trimmed = v.trim();
    let num = trimmed.strip_suffix('%')?;
    // Stored as a fraction, not "percent points".
    parse_fixed_point(num.trim(), scale).map(|scaled_percent| scaled_percent / 100)
}

fn parse_fixed_point(v: &str, scale: u8) -> Option<i64> {
    let mut s = v.trim();
    if s.is_empty() {
        return None;
    }

    let mut sign: i64 = 1;
    if let Some(rest) = s.strip_prefix('-') {
        sign = -1;
        s = rest;
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest;
    }

    let mut parts = s.splitn(2, '.');
    let int_part_raw = parts.next().unwrap_or("");
    let frac_part_raw = parts.next().unwrap_or("");

    let int_part_str: String = int_part_raw.chars().filter(|c| *c != ',').collect();
    let int_part: i64 = if int_part_str.is_empty() {
        0
    } else {
        int_part_str.parse().ok()?
    };

    let mut frac_value: i64 = 0;
    let mut digits: u8 = 0;
    for ch in frac_part_raw.chars() {
        if digits >= scale {
            break;
        }
        if !ch.is_ascii_digit() {
            break;
        }
        frac_value = frac_value * 10 + (ch as i64 - '0' as i64);
        digits += 1;
    }

    while digits < scale {
        frac_value *= 10;
        digits += 1;
    }

    let scale_factor = 10i64.pow(scale as u32);
    Some(sign * (int_part.saturating_mul(scale_factor) + frac_value))
}

fn parse_datetime_millis(v: &str) -> Option<i64> {
    let s = v.trim();
    if s.len() < 10 {
        return None;
    }

    // Fast path for YYYY-MM-DD[...]
    let bytes = s.as_bytes();
    if bytes.get(4) != Some(&b'-') || bytes.get(7) != Some(&b'-') {
        return None;
    }

    let year: i32 = s.get(0..4)?.parse().ok()?;
    let month: u32 = s.get(5..7)?.parse().ok()?;
    let day: u32 = s.get(8..10)?.parse().ok()?;
    let days = days_from_civil(year, month, day)?;

    let mut millis: i64 = days * 86_400_000;
    if s.len() == 10 {
        return Some(millis);
    }

    let rest = s.get(10..)?.trim_start_matches(['T', ' ']);
    if rest.is_empty() {
        return Some(millis);
    }

    let (time_part, tz_part) = match rest.find(|c| c == 'Z' || c == '+' || c == '-') {
        Some(idx) => (&rest[..idx], &rest[idx..]),
        None => (rest, ""),
    };

    let mut t_iter = time_part.splitn(3, ':');
    let h: i64 = t_iter.next()?.parse().ok()?;
    let m: i64 = t_iter.next()?.parse().ok()?;
    let s_part = t_iter.next()?.trim();

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

    // Only support 'Z' or empty TZ for now (keep dependency-free).
    if !tz_part.is_empty() && tz_part != "Z" {
        return None;
    }
    Some(millis)
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

fn trim_newline(s: &mut String) {
    while s.ends_with(['\n', '\r']) {
        s.pop();
    }
}

#[derive(Clone, Debug)]
enum CsvField {
    Range { start: usize, end: usize },
    Owned(String),
}

impl CsvField {
    fn as_str<'a>(&'a self, line: &'a str) -> &'a str {
        match self {
            Self::Range { start, end } => &line[*start..*end],
            Self::Owned(s) => s.as_str(),
        }
    }
}

fn parse_csv_fields(line: &str, delimiter: u8, out: &mut Vec<CsvField>) {
    out.clear();
    let bytes = line.as_bytes();
    if bytes.is_empty() {
        out.push(CsvField::Range { start: 0, end: 0 });
        return;
    }

    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'"' {
            // Quoted field.
            i += 1;
            let mut buf = String::new();
            while i < bytes.len() {
                match bytes[i] {
                    b'"' => {
                        if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                            buf.push('"');
                            i += 2;
                        } else {
                            i += 1;
                            break;
                        }
                    }
                    b => {
                        buf.push(b as char);
                        i += 1;
                    }
                }
            }

            // Skip until delimiter/end.
            while i < bytes.len() && bytes[i] != delimiter {
                i += 1;
            }
            if i < bytes.len() && bytes[i] == delimiter {
                i += 1;
            }
            out.push(CsvField::Owned(buf));
        } else {
            let start = i;
            while i < bytes.len() && bytes[i] != delimiter {
                i += 1;
            }
            let end = i;
            if i < bytes.len() && bytes[i] == delimiter {
                i += 1;
            }
            out.push(CsvField::Range { start, end });
        }
    }

    if bytes.last() == Some(&delimiter) {
        out.push(CsvField::Range {
            start: bytes.len(),
            end: bytes.len(),
        });
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

