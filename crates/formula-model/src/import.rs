use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{BufRead, BufReader, Cursor, Read};
use std::sync::Arc;

use csv::ByteRecord;
use encoding_rs::{CoderResult, UTF_16BE, UTF_16LE, WINDOWS_1252};
use formula_columnar::{
    ColumnSchema, ColumnType as ColumnarType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig,
    TableOptions, Value as ColumnarValue,
};
use thiserror::Error;

use crate::{CellRef, SheetNameError, Workbook, Worksheet, WorksheetId};

#[derive(Clone, Debug)]
pub struct CsvOptions {
    /// Field delimiter byte (`,`/`;`/tab/`|`).
    ///
    /// Use [`CSV_DELIMITER_AUTO`] to automatically detect the delimiter from the first few records
    /// (Excel-like behavior). Auto-detection also honors Excel's `sep=<delimiter>` directive when
    /// it appears as the first line in the file.
    ///
    /// When auto-detecting, [`CsvOptions::decimal_separator`] is used as a locale hint: if the
    /// decimal separator is `,`, delimiter detection will bias toward `;` in ambiguous cases
    /// (matching common Excel locale behavior).
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
            delimiter: CSV_DELIMITER_AUTO,
            has_header: true,
            sample_rows: 100,
            page_size_rows: 65_536,
            cache_entries: 64,
            encoding: CsvTextEncoding::default(),
            decimal_separator: '.',
            date_order: CsvDateOrder::default(),
            timestamp_tz_policy: CsvTimestampTzPolicy::default(),
            currency_symbols: vec!['$', '€', '£', '¥'],
        }
    }
}

/// Sentinel value for [`CsvOptions::delimiter`] indicating that the delimiter should be
/// auto-detected.
///
/// Auto-detection considers (in priority order) `,`, `;`, tab, and `|`.
pub const CSV_DELIMITER_AUTO: u8 = 0;
/// Guess a CSV delimiter from a byte sample.
///
/// This mirrors the importer behavior when using [`CSV_DELIMITER_AUTO`], but operates on an
/// in-memory prefix (useful for callers that need the delimiter before constructing their own CSV
/// reader).
///
/// Note:
/// - This honors Excel's `sep=<delimiter>` directive when present on the first line.
/// - This uses `.` as the decimal separator hint; for locales that use `,` as the decimal
///   separator, prefer [`sniff_csv_delimiter_with_decimal_separator`].
pub fn sniff_csv_delimiter(sample: &[u8]) -> u8 {
    sniff_csv_delimiter_with_decimal_separator(sample, '.')
}

/// Guess a CSV delimiter from a byte sample, using a locale-aware decimal separator.
///
/// When `decimal_separator` is `,`, this biases delimiter selection toward `;` in ambiguous cases,
/// matching common Excel locale behavior (CSV files frequently use `;` as the field delimiter in
/// locales where `,` is used as the decimal separator).
///
/// This still honors Excel's `sep=<delimiter>` directive when present on the first line.
pub fn sniff_csv_delimiter_with_decimal_separator(sample: &[u8], decimal_separator: char) -> u8 {
    // Best-effort: handle UTF-16 samples by transcoding to UTF-8 before sniffing.
    if sample.starts_with(&[0xFF, 0xFE]) || sample.starts_with(&[0xFE, 0xFF]) {
        let decoder = if sample.starts_with(&[0xFF, 0xFE]) {
            UTF_16LE.new_decoder_with_bom_removal()
        } else {
            UTF_16BE.new_decoder_with_bom_removal()
        };
        let utf8 = Utf16ToUtf8Reader::new(Cursor::new(sample), decoder);
        let mut reader = BufReader::new(utf8);
        return sniff_csv_delimiter_prefix(&mut reader, decimal_separator)
            .map(|(_, delimiter)| delimiter)
            .unwrap_or(b',');
    }
    if let Some(encoding) = detect_utf16_bomless_encoding(sample) {
        let utf8 =
            Utf16ToUtf8Reader::new(Cursor::new(sample), encoding.new_decoder_with_bom_removal());
        let mut reader = BufReader::new(utf8);
        return sniff_csv_delimiter_prefix(&mut reader, decimal_separator)
            .map(|(_, delimiter)| delimiter)
            .unwrap_or(b',');
    }

    let mut cursor = Cursor::new(sample);
    sniff_csv_delimiter_prefix(&mut cursor, decimal_separator)
        .map(|(_, delimiter)| delimiter)
        .unwrap_or(b',')
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

impl Default for CsvTextEncoding {
    fn default() -> Self {
        CsvTextEncoding::Auto
    }
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
    Parse {
        row: u64,
        column: u64,
        reason: String,
    },
    #[error(transparent)]
    SheetName(#[from] SheetNameError),
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

/// Import a CSV stream into a [`ColumnarTable`] without materializing a full grid.
pub fn import_csv_to_columnar_table<R: BufRead>(
    mut reader: R,
    options: CsvOptions,
) -> Result<ColumnarTable, CsvImportError> {
    // Excel can export delimited text as UTF-16LE/BE (commonly via "Unicode Text"). The CSV parser
    // expects an 8-bit stream, so detect UTF-16 and transcode to UTF-8 on the fly.
    let buf = reader.fill_buf().map_err(CsvImportError::Io)?;
    if buf.starts_with(&[0xFF, 0xFE]) || buf.starts_with(&[0xFE, 0xFF]) {
        let decoder = if buf.starts_with(&[0xFF, 0xFE]) {
            UTF_16LE.new_decoder_with_bom_removal()
        } else {
            UTF_16BE.new_decoder_with_bom_removal()
        };
        let utf8 = Utf16ToUtf8Reader::new(reader, decoder);
        let mut options = options;
        // The transcoder yields valid UTF-8, so force UTF-8 decoding for field values.
        options.encoding = CsvTextEncoding::Utf8;
        return import_csv_to_columnar_table_impl(BufReader::new(utf8), options);
    }
    if let Some(encoding) = detect_utf16_bomless_encoding(buf) {
        let utf8 = Utf16ToUtf8Reader::new(reader, encoding.new_decoder_with_bom_removal());
        let mut options = options;
        options.encoding = CsvTextEncoding::Utf8;
        return import_csv_to_columnar_table_impl(BufReader::new(utf8), options);
    }

    import_csv_to_columnar_table_impl(reader, options)
}

fn import_csv_to_columnar_table_impl<R: BufRead>(
    mut reader: R,
    options: CsvOptions,
) -> Result<ColumnarTable, CsvImportError> {
    let mut options = options;
    let (sniff_prefix, mut delimiter) = if options.delimiter == CSV_DELIMITER_AUTO {
        sniff_csv_delimiter_prefix(&mut reader, options.decimal_separator)
            .map_err(CsvImportError::Io)?
    } else {
        (Vec::new(), options.delimiter)
    };

    // Excel uses both a list separator (delimiter) and a decimal separator. For files imported with
    // auto-delimiter detection, attempt to infer a decimal separator from the sample when the
    // caller did not explicitly opt into a locale (`CsvOptions::decimal_separator` defaults to
    // `.`). This improves compatibility with semicolon-delimited, decimal-comma CSVs that do not
    // include a `sep=` directive.
    if options.delimiter == CSV_DELIMITER_AUTO && options.decimal_separator == '.' {
        if let Some(inferred) = infer_csv_decimal_separator_from_sample(&sniff_prefix) {
            options.decimal_separator = inferred;
            delimiter = sniff_csv_delimiter_with_decimal_separator(&sniff_prefix, inferred);
        }
    }

    // If we sniffed, re-play the bytes we consumed so the CSV reader sees the full stream.
    let reader = Cursor::new(sniff_prefix).chain(reader);

    let mut csv_reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
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

    // Excel supports a special first line `sep=<delimiter>` which explicitly specifies the
    // delimiter. When present, Excel uses it for delimiter selection and does not surface it as
    // a header/data row.
    if is_csv_sep_directive_record(&record, record_index, delimiter, options.encoding) {
        record.clear();
        let has_next = csv_reader
            .read_byte_record(&mut record)
            .map_err(|e| map_csv_error(e, record_index + 1))?;
        if !has_next {
            return Err(CsvImportError::EmptyInput);
        }
        record_index += 1;
    }

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
        parse_row_to_values(
            row,
            &column_types,
            &options,
            &mut string_pool,
            &mut row_values,
        );
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
                    let field = decode_field(raw, record_index, i as u64 + 1, options.encoding)?;
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

/// Streaming UTF-16 → UTF-8 transcoder used for CSV import.
///
/// This intentionally keeps the implementation small and best-effort: it relies on `encoding_rs`
/// for correctness around surrogate pairs and malformed sequences.
struct Utf16ToUtf8Reader<R> {
    inner: R,
    decoder: encoding_rs::Decoder,
    pending: Vec<u8>,
    out_buf: Vec<u8>,
    out_pos: usize,
    eof: bool,
    finished: bool,
}

impl<R: Read> Utf16ToUtf8Reader<R> {
    fn new(inner: R, decoder: encoding_rs::Decoder) -> Self {
        Self {
            inner,
            decoder,
            pending: Vec::new(),
            out_buf: Vec::new(),
            out_pos: 0,
            eof: false,
            finished: false,
        }
    }
}

impl<R: Read> Read for Utf16ToUtf8Reader<R> {
    fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
        if out.is_empty() {
            return Ok(0);
        }
        if self.finished {
            return Ok(0);
        }

        loop {
            if self.out_pos < self.out_buf.len() {
                let available = self.out_buf.len() - self.out_pos;
                let n = out.len().min(available);
                out[..n].copy_from_slice(&self.out_buf[self.out_pos..self.out_pos + n]);
                self.out_pos += n;
                if self.out_pos == self.out_buf.len() {
                    self.out_buf.clear();
                    self.out_pos = 0;
                }
                return Ok(n);
            }

            // Ensure we have some pending input unless we've hit EOF.
            if self.pending.is_empty() && !self.eof {
                let mut buf = [0u8; 8 * 1024];
                let n = self.inner.read(&mut buf)?;
                if n == 0 {
                    self.eof = true;
                } else {
                    self.pending.extend_from_slice(&buf[..n]);
                }
            }

            let mut decoded = [0u8; 8 * 1024];
            let (result, read, written, _) =
                self.decoder
                    .decode_to_utf8(&self.pending, &mut decoded, self.eof);
            if read > 0 {
                self.pending.drain(..read);
            }
            if written > 0 {
                self.out_buf.clear();
                self.out_pos = 0;
                self.out_buf.extend_from_slice(&decoded[..written]);
                continue;
            }

            match result {
                CoderResult::OutputFull => {
                    // `decoded` is large enough for any single UTF-8 code point, so if the decoder
                    // reports an output buffer overflow with no progress, treat it as an internal
                    // error.
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "UTF-16 to UTF-8 decoder made no progress (output full)",
                    ));
                }
                CoderResult::InputEmpty => {
                    if self.eof {
                        // The decoder has reached EOF and has no more output to flush. Mark it as
                        // finished so subsequent `read` calls return EOF without invoking the
                        // decoder again.
                        self.finished = true;
                        return Ok(0);
                    }
                    // Need more input. This can happen even when `pending` is non-empty (e.g. the
                    // input ended mid-code-unit), so always try to read another chunk.
                    let mut buf = [0u8; 8 * 1024];
                    let n = self.inner.read(&mut buf)?;
                    if n == 0 {
                        self.eof = true;
                    } else {
                        self.pending.extend_from_slice(&buf[..n]);
                    }
                    continue;
                }
            }
        }
    }
}

fn detect_utf16_bomless_encoding(buf: &[u8]) -> Option<&'static encoding_rs::Encoding> {
    // Best-effort detection for UTF-16 inputs that lack a BOM.
    //
    // This is intentionally conservative: only treat the stream as UTF-16 if it contains a high
    // proportion of NUL bytes that are overwhelmingly concentrated on either the even or odd byte
    // positions (ASCII-like UTF-16LE/BE pattern).
    let len = buf.len().min(2048);
    let len = len - (len % 2);
    if len < 4 {
        return None;
    }
    let sample = &buf[..len];

    // Prefer endianness based on common delimiter/newline patterns instead of a raw NUL-byte ratio.
    // This handles BOM-less UTF-16 files where the content is mostly non-ASCII (low NUL ratio), but
    // still contains ASCII record separators like `\t` and `\r\n`.
    let mut le_markers = 0usize;
    let mut be_markers = 0usize;
    let mut even_zero = 0usize;
    let mut odd_zero = 0usize;

    for (idx, b) in sample.iter().enumerate() {
        if *b == 0 {
            if idx % 2 == 0 {
                even_zero += 1;
            } else {
                odd_zero += 1;
            }
        }
    }

    const MARKERS: [u8; 6] = [b',', b';', b'\t', b'|', b'\r', b'\n'];
    for pair in sample.chunks_exact(2) {
        let a = pair[0];
        let b = pair[1];
        if b == 0 && MARKERS.contains(&a) {
            le_markers += 1;
        }
        if a == 0 && MARKERS.contains(&b) {
            be_markers += 1;
        }
    }

    // Require at least a couple marker hits (e.g. tab + CRLF) to avoid misclassifying arbitrary
    // binary data that happens to contain a few NUL bytes.
    const MIN_MARKERS: usize = 2;
    if le_markers >= MIN_MARKERS || be_markers >= MIN_MARKERS {
        if le_markers > be_markers {
            return Some(UTF_16LE);
        }
        if be_markers > le_markers {
            return Some(UTF_16BE);
        }
    }

    // Fall back to NUL-byte skew if marker counts are inconclusive.
    if odd_zero > even_zero.saturating_mul(3) {
        Some(UTF_16LE)
    } else if even_zero > odd_zero.saturating_mul(3) {
        Some(UTF_16BE)
    } else {
        None
    }
}

const CSV_SNIFF_DELIMITERS: [u8; 4] = [b',', b';', b'\t', b'|'];
const CSV_SNIFF_MAX_RECORDS: usize = 20;
const CSV_SNIFF_MAX_BYTES: usize = 64 * 1024;

fn is_csv_sep_directive_record(
    record: &ByteRecord,
    row: u64,
    delimiter: u8,
    encoding: CsvTextEncoding,
) -> bool {
    if row != 1 {
        return false;
    }
    if !CSV_SNIFF_DELIMITERS.contains(&delimiter) {
        return false;
    }
    let first = match record.get(0) {
        Some(bytes) => match decode_field(bytes, row, 1, encoding) {
            Ok(s) => s,
            Err(_) => return false,
        },
        None => return false,
    };
    if !first.trim().eq_ignore_ascii_case("sep=") {
        return false;
    }
    if record.len() < 2 {
        return false;
    }
    for (idx, field) in record.iter().enumerate().skip(1) {
        let col = idx as u64 + 1;
        let value = match decode_field(field, row, col, encoding) {
            Ok(s) => s,
            Err(_) => return false,
        };
        if !value.trim().is_empty() {
            return false;
        }
    }
    true
}

fn parse_csv_sep_directive(prefix: &[u8]) -> Option<u8> {
    // Excel supports a special first line `sep=<delimiter>` which explicitly specifies the CSV
    // delimiter (commonly used for semicolon-delimited CSVs in locales where `,` is the decimal
    // separator). When present, Excel uses it to pick the delimiter and does not treat it as a
    // data/header record.
    //
    // Example: `sep=;` (followed by `\n` or `\r\n`).
    let line_end = prefix.iter().position(|b| matches!(b, b'\n' | b'\r'))?;
    let mut line = &prefix[..line_end];
    if line.starts_with(&[0xEF, 0xBB, 0xBF]) {
        line = &line[3..];
    }

    // Skip leading ASCII whitespace.
    let mut start = 0usize;
    while start < line.len() && matches!(line[start], b' ' | b'\t') {
        start += 1;
    }
    let line = &line[start..];

    // Parse `sep=` prefix (case-insensitive).
    if line.len() < 5 || !line[..4].eq_ignore_ascii_case(b"sep=") {
        return None;
    }

    // Excel's directive is typically `sep=<delim>` with no extra spacing, but be tolerant of an
    // extra space after `=` (e.g. `sep= ;`). Only skip ASCII spaces here so we don't accidentally
    // skip a tab delimiter.
    let mut delim_idx = 4usize;
    while delim_idx < line.len() && line[delim_idx] == b' ' {
        delim_idx += 1;
    }
    let delimiter = *line.get(delim_idx)?;
    if !CSV_SNIFF_DELIMITERS.contains(&delimiter) {
        return None;
    }

    // Allow trailing ASCII whitespace after the delimiter.
    let mut rest_idx = delim_idx + 1;
    while rest_idx < line.len() && matches!(line[rest_idx], b' ' | b'\t') {
        rest_idx += 1;
    }
    if rest_idx != line.len() {
        return None;
    }

    Some(delimiter)
}

fn infer_csv_decimal_separator_from_sample(sample: &[u8]) -> Option<char> {
    // Best-effort locale inference for numeric parsing.
    //
    // Excel opens CSVs using the system's list separator and decimal separator. We don't have
    // access to OS locale here, but we can infer the decimal separator from the data when
    // auto-detecting the delimiter (the common "EU" case is `;` delimiter with `,` decimal).
    //
    // We intentionally keep this conservative: only return `,` when we see evidence of decimal
    // commas *and* no evidence of decimal dots. (Default behavior remains decimal-dot.)
    //
    // Additionally, avoid inferring decimal-comma for comma-delimited CSVs (where commas are more
    // likely to be field separators than decimal separators) by requiring some evidence of a
    // non-comma delimiter (`;`, tab, or `|`) appearing roughly once per line.
    // Estimate the number of records in the prefix. Prefer `\n` count (covers LF + CRLF). Fall
    // back to counting `\r` for old-Mac/CR-only inputs.
    let lf = sample.iter().filter(|b| **b == b'\n').count();
    let cr = sample.iter().filter(|b| **b == b'\r').count();
    let line_breaks = (if lf > 0 { lf } else { cr }).max(1);
    let non_comma_delims = sample
        .iter()
        .filter(|b| matches!(**b, b';' | b'\t' | b'|'))
        .count();
    if non_comma_delims < line_breaks {
        return None;
    }

    let mut comma_evidence = 0usize;
    let mut dot_evidence = 0usize;

    let mut i = 0usize;
    while i < sample.len() {
        if !sample[i].is_ascii_digit() {
            i += 1;
            continue;
        }

        let start = i;
        i += 1;
        while i < sample.len() {
            let b = sample[i];
            if b.is_ascii_digit() || matches!(b, b'.' | b',' | b'e' | b'E' | b'+' | b'-') {
                i += 1;
            } else {
                break;
            }
        }

        let token = &sample[start..i];
        let exp_idx = token.iter().position(|b| matches!(*b, b'e' | b'E'));
        let mantissa = match exp_idx {
            Some(idx) => &token[..idx],
            None => token,
        };

        let dot_count = mantissa.iter().filter(|b| **b == b'.').count();
        let comma_count = mantissa.iter().filter(|b| **b == b',').count();

        if dot_count > 0 && comma_count > 0 {
            let last_dot = mantissa.iter().rposition(|b| *b == b'.');
            let last_comma = mantissa.iter().rposition(|b| *b == b',');
            match (last_dot, last_comma) {
                (Some(d), Some(c)) => {
                    if d > c {
                        dot_evidence += 1;
                    } else {
                        comma_evidence += 1;
                    }
                }
                _ => {}
            }
            continue;
        }

        if dot_count == 1 && comma_count == 0 {
            let sep_idx = mantissa.iter().position(|b| *b == b'.').unwrap_or(0);
            let digits_after = mantissa[sep_idx + 1..]
                .iter()
                .take_while(|b| b.is_ascii_digit())
                .count();
            if digits_after > 0 && digits_after != 3 {
                dot_evidence += 1;
            }
            continue;
        }

        if comma_count == 1 && dot_count == 0 {
            let sep_idx = mantissa.iter().position(|b| *b == b',').unwrap_or(0);
            let digits_after = mantissa[sep_idx + 1..]
                .iter()
                .take_while(|b| b.is_ascii_digit())
                .count();
            if digits_after > 0 && digits_after != 3 {
                comma_evidence += 1;
            }
            continue;
        }
    }

    if comma_evidence > 0 && dot_evidence == 0 {
        return Some(',');
    }
    None
}

fn sniff_csv_delimiter_prefix<R: Read>(
    reader: &mut R,
    decimal_separator: char,
) -> Result<(Vec<u8>, u8), std::io::Error> {
    let mut prefix: Vec<u8> = Vec::new();
    let mut hists: Vec<HashMap<usize, usize>> = CSV_SNIFF_DELIMITERS
        .iter()
        .map(|_| HashMap::new())
        .collect();
    let mut delim_counts: [usize; CSV_SNIFF_DELIMITERS.len()] = [0; CSV_SNIFF_DELIMITERS.len()];

    let mut in_quotes = false;
    let mut pending_quote = false;
    let mut pending_cr = false;
    let mut record_len: usize = 0;
    let mut sampled_records: usize = 0;
    let mut hit_eof = false;

    let mut buf = [0u8; 8 * 1024];

    'read: while prefix.len() < CSV_SNIFF_MAX_BYTES && sampled_records < CSV_SNIFF_MAX_RECORDS {
        let to_read = (CSV_SNIFF_MAX_BYTES - prefix.len()).min(buf.len());
        let n = reader.read(&mut buf[..to_read])?;
        if n == 0 {
            hit_eof = true;
            break;
        }

        prefix.extend_from_slice(&buf[..n]);
        if let Some(delimiter) = parse_csv_sep_directive(&prefix) {
            return Ok((prefix, delimiter));
        }

        for &byte in &buf[..n] {
            if sampled_records >= CSV_SNIFF_MAX_RECORDS {
                break 'read;
            }

            // Process one byte at a time, with simple state to handle `""` escapes and CRLF.
            let b = byte;
            loop {
                if pending_cr && !in_quotes {
                    pending_cr = false;
                    if b == b'\n' {
                        // Swallow the `\n` in a `\r\n` record terminator.
                        break;
                    }
                    // Otherwise, re-process `b` as the start of the next record.
                    continue;
                }

                if pending_quote {
                    if b == b'"' {
                        // Escaped quote (`""`) inside a quoted field.
                        pending_quote = false;
                        record_len += 1;
                        break;
                    }

                    // Closing quote. Re-process this byte outside quotes.
                    pending_quote = false;
                    in_quotes = false;
                    continue;
                }

                if in_quotes {
                    record_len += 1;
                    if b == b'"' {
                        pending_quote = true;
                    }
                    break;
                }

                match b {
                    b'"' => {
                        in_quotes = true;
                        record_len += 1;
                        break;
                    }
                    b'\r' => {
                        commit_csv_sniff_record(
                            record_len,
                            &mut delim_counts,
                            &mut hists,
                            &mut sampled_records,
                        );
                        record_len = 0;
                        delim_counts.fill(0);
                        pending_cr = true;
                        break;
                    }
                    b'\n' => {
                        commit_csv_sniff_record(
                            record_len,
                            &mut delim_counts,
                            &mut hists,
                            &mut sampled_records,
                        );
                        record_len = 0;
                        delim_counts.fill(0);
                        break;
                    }
                    _ => {
                        record_len += 1;
                        for (idx, delim) in CSV_SNIFF_DELIMITERS.iter().enumerate() {
                            if b == *delim {
                                delim_counts[idx] += 1;
                            }
                        }
                        break;
                    }
                }
            }
        }
    }

    // EOF can terminate the last record without a newline.
    if hit_eof && sampled_records < CSV_SNIFF_MAX_RECORDS {
        if pending_quote {
            // Treat a trailing quote at EOF as a closing quote.
            in_quotes = false;
        }
        if record_len > 0 && !in_quotes {
            commit_csv_sniff_record(
                record_len,
                &mut delim_counts,
                &mut hists,
                &mut sampled_records,
            );
        }
    }

    let delimiter = select_sniffed_csv_delimiter(&hists, decimal_separator);
    Ok((prefix, delimiter))
}

fn commit_csv_sniff_record(
    record_len: usize,
    delim_counts: &mut [usize; CSV_SNIFF_DELIMITERS.len()],
    hists: &mut [HashMap<usize, usize>],
    sampled_records: &mut usize,
) {
    if record_len == 0 {
        return;
    }

    for (idx, count) in delim_counts.iter().enumerate() {
        let fields = *count + 1;
        *hists[idx].entry(fields).or_insert(0) += 1;
    }
    *sampled_records += 1;
}

fn select_sniffed_csv_delimiter(hists: &[HashMap<usize, usize>], decimal_separator: char) -> u8 {
    // Pick the delimiter whose sampled records most frequently share the same column count (>1).
    //
    // If two delimiters are equally consistent, prefer the one with the higher consistent column
    // count (mode field count), matching Excel-like behavior.
    //
    // Tie-break:
    // - In locales where `,` is the decimal separator, `,` commonly appears inside numeric values.
    //   If `,` is equally consistent as another delimiter, prefer the non-comma delimiter to avoid
    //   splitting decimal values into separate fields (Excel-like behavior; commonly `;`).
    // - Otherwise, tie-break deterministically in `CSV_SNIFF_DELIMITERS` order: `,` > `;` > tab > `|`.

    #[derive(Clone, Copy, Debug, Default)]
    struct Stats {
        mode_count: usize,
        mode_fields: usize,
    }

    let mut stats: [Stats; CSV_SNIFF_DELIMITERS.len()] =
        [Stats::default(); CSV_SNIFF_DELIMITERS.len()];

    for (idx, hist) in hists.iter().enumerate() {
        let mut mode_count = 0usize;
        let mut mode_fields = 0usize;
        for (fields, count) in hist.iter().filter(|(fields, _)| **fields > 1) {
            if *count > mode_count || (*count == mode_count && *fields > mode_fields) {
                mode_count = *count;
                mode_fields = *fields;
            }
        }
        stats[idx] = Stats {
            mode_count,
            mode_fields,
        };
    }

    let max_mode_count = stats.iter().map(|s| s.mode_count).max().unwrap_or(0);
    if max_mode_count == 0 {
        return b',';
    }

    if decimal_separator == ',' {
        // In decimal-comma locales, `,` commonly appears inside numeric values. When comma-delimited
        // parsing is equally consistent as another delimiter, bias away from `,` so we don't split
        // decimal values into separate fields.
        if let Some(comma_idx) = CSV_SNIFF_DELIMITERS.iter().position(|d| *d == b',') {
            if stats[comma_idx].mode_count == max_mode_count {
                let mut best_non_comma: Option<usize> = None;
                for (idx, stat) in stats.iter().enumerate() {
                    if idx == comma_idx {
                        continue;
                    }
                    if stat.mode_count != max_mode_count {
                        continue;
                    }
                    match best_non_comma {
                        None => best_non_comma = Some(idx),
                        Some(best) => {
                            if stat.mode_fields > stats[best].mode_fields {
                                best_non_comma = Some(idx);
                            } else if stat.mode_fields == stats[best].mode_fields && idx < best {
                                best_non_comma = Some(idx);
                            }
                        }
                    }
                }

                if let Some(idx) = best_non_comma {
                    return CSV_SNIFF_DELIMITERS[idx];
                }
            }
        }
    }

    let mut best_idx: Option<usize> = None;
    for (idx, stat) in stats.iter().enumerate() {
        if stat.mode_count != max_mode_count {
            continue;
        }
        match best_idx {
            None => best_idx = Some(idx),
            Some(best) => {
                if stat.mode_fields > stats[best].mode_fields {
                    best_idx = Some(idx);
                } else if stat.mode_fields == stats[best].mode_fields && idx < best {
                    // Deterministic tie-break in CSV_SNIFF_DELIMITERS order.
                    best_idx = Some(idx);
                }
            }
        }
    }

    best_idx
        .map(|idx| CSV_SNIFF_DELIMITERS[idx])
        .unwrap_or(b',')
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
    match column_type {
        // Preserve whitespace in string values. Unlike numeric parsing (which should ignore
        // surrounding whitespace), CSV fields can legitimately contain leading/trailing spaces that
        // Excel surfaces as part of the cell text.
        ColumnarType::String => {
            if field.is_empty() {
                ColumnarValue::Null
            } else {
                ColumnarValue::String(string_pool.intern(field))
            }
        }
        _ => {
            let v = field.trim();
            if v.is_empty() {
                return ColumnarValue::Null;
            }

            match column_type {
                ColumnarType::Number => parse_number_f64(v, options)
                    .map(ColumnarValue::Number)
                    .unwrap_or(ColumnarValue::Null),
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
                ColumnarType::String => unreachable!("handled above"),
            }
        }
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
    if sample_rows.is_empty() {
        return vec![ColumnarType::String; column_count];
    }

    let mut out = Vec::with_capacity(column_count);
    for col in 0..column_count {
        let mut saw_value = false;
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
            saw_value = true;
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

        let ty = if !saw_value {
            ColumnarType::String
        } else if is_bool && saw_text_bool {
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
        CsvTextEncoding::Utf8 => {
            std::str::from_utf8(field)
                .map(Cow::Borrowed)
                .map_err(|e| CsvImportError::Parse {
                    row,
                    column,
                    reason: format!("invalid UTF-8: {e}"),
                })
        }
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
        Self {
            set: HashMap::new(),
        }
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
