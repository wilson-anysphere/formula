use std::collections::HashMap;

use formula_model::autofilter::{
    FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueCustomFilter,
    TextMatch, TextMatchKind,
};
use formula_model::Range;

use super::{records, strings, BiffVersion};

/// AUTOFILTER (worksheet substream) [MS-XLS 2.4.31].
const RECORD_AUTOFILTER: u16 = 0x009E;

// AUTOFILTER grbit bits (best-effort; based on [MS-XLS] AUTOFILTER).
//
// Note: BIFF stores additional flags (simple filters, Top10, etc). We only
// implement the subset needed to recover basic criteria.
const AUTOFILTER_FLAG_AND: u16 = 0x0001;
const AUTOFILTER_FLAG_TOP10: u16 = 0x0008;
const AUTOFILTER_FLAG_TOP: u16 = 0x0010;
const AUTOFILTER_FLAG_PERCENT: u16 = 0x0020;

// BIFF8 XLUnicodeString flags (mirrors `strings.rs`).
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Default)]
pub(crate) struct ParsedAutoFilterCriteria {
    pub(crate) filter_columns: Vec<FilterColumn>,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort parse of worksheet-level AutoFilter criteria from `AUTOFILTER` records.
///
/// BIFF stores the filtered range separately (via the built-in `_FilterDatabase` defined name);
/// this helper only imports per-column criteria where possible.
///
/// This is intentionally best-effort: malformed records are skipped and reported as warnings.
pub(crate) fn parse_biff_sheet_autofilter_criteria(
    workbook_stream: &[u8],
    start: usize,
    biff: BiffVersion,
    codepage: u16,
    autofilter_range: Range,
) -> Result<ParsedAutoFilterCriteria, String> {
    let mut out = ParsedAutoFilterCriteria::default();

    // `.xls` AutoFilter criteria parsing is currently only implemented for BIFF8.
    if biff != BiffVersion::Biff8 {
        return Ok(out);
    }

    let allows_continuation = |record_id: u16| record_id == RECORD_AUTOFILTER;
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    // Prefer the last AUTOFILTER record per column.
    let mut cols: HashMap<u32, FilterColumn> = HashMap::new();

    // AUTOFILTER records can encode the target column either as:
    // - a 0-based entry index within the AutoFilter range (`iEntry`), or
    // - (observed in some producers) the absolute worksheet column index.
    //
    // We do a first pass to collect raw entry ids and choose a best-effort interpretation for
    // the whole sheet.
    let mut autofilter_records: Vec<records::LogicalBiffRecord<'_>> = Vec::new();
    let mut raw_entries: Vec<u32> = Vec::new();

    for record in iter {
        let record = match record {
            Ok(r) => r,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // Stop at the beginning of the next substream (worksheet BOF).
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_AUTOFILTER => {
                let data = record.data.as_ref();
                if data.len() >= 2 {
                    raw_entries.push(u16::from_le_bytes([data[0], data[1]]) as u32);
                }
                autofilter_records.push(record);
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let entry_mode = choose_entry_mode(&raw_entries, autofilter_range);

    for record in autofilter_records {
        match parse_autofilter_record(
            &record,
            codepage,
            autofilter_range,
            entry_mode,
            &mut out.warnings,
        ) {
            Ok(Some(col)) => {
                cols.insert(col.col_id, col);
            }
            Ok(None) => {}
            Err(err) => out.warnings.push(format!(
                "failed to decode AUTOFILTER record at offset {}: {err}",
                record.offset
            )),
        }
    }

    let mut filter_columns: Vec<FilterColumn> = cols.into_values().collect();
    filter_columns.sort_by_key(|c| c.col_id);
    out.filter_columns = filter_columns;

    Ok(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AutoFilterEntryMode {
    Relative,
    Absolute,
}

fn choose_entry_mode(raw_entries: &[u32], range: Range) -> AutoFilterEntryMode {
    let start_col = range.start.col;
    let end_col = range.end.col;
    let width = end_col.saturating_sub(start_col).saturating_add(1);

    let mut relative_hits = 0usize;
    let mut absolute_hits = 0usize;

    for &entry in raw_entries {
        // Relative encoding: entry is 0..(width-1).
        if entry < width {
            relative_hits += 1;
        }
        // Absolute encoding: entry is an in-range worksheet column index.
        if entry >= start_col && entry <= end_col {
            absolute_hits += 1;
        }
    }

    if absolute_hits > relative_hits {
        AutoFilterEntryMode::Absolute
    } else {
        AutoFilterEntryMode::Relative
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AutoFilterOp {
    None,
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    GreaterThanOrEqual,
    LessThanOrEqual,
    Contains,
    DoesNotContain,
    BeginsWith,
    EndsWith,
    DoesNotBeginWith,
    DoesNotEndWith,
    Unknown(u8),
}

impl AutoFilterOp {
    fn from_biff_code(code: u8) -> Self {
        // Best-effort mapping: this matches the operator ordering used by several BIFF records.
        match code {
            0 => AutoFilterOp::None,
            1 => AutoFilterOp::Between,
            2 => AutoFilterOp::NotBetween,
            3 => AutoFilterOp::Equal,
            4 => AutoFilterOp::NotEqual,
            5 => AutoFilterOp::GreaterThan,
            6 => AutoFilterOp::LessThan,
            7 => AutoFilterOp::GreaterThanOrEqual,
            8 => AutoFilterOp::LessThanOrEqual,
            // Best-effort text operators (OOXML `customFilter/@operator` values).
            //
            // These have been observed in BIFF8 AutoFilter records produced by some writers.
            // Positive operators (`contains`, `beginsWith`, `endsWith`) import as `TextMatch` so the
            // filter is evaluable by the engine; negative operators are preserved as `OpaqueCustom`
            // so the operator/value pair can round-trip to XLSX.
            9 => AutoFilterOp::Contains,
            10 => AutoFilterOp::BeginsWith,
            11 => AutoFilterOp::EndsWith,
            12 => AutoFilterOp::DoesNotContain,
            13 => AutoFilterOp::DoesNotBeginWith,
            14 => AutoFilterOp::DoesNotEndWith,
            other => AutoFilterOp::Unknown(other),
        }
    }

    fn to_ooxml_operator_name(self) -> Option<&'static str> {
        match self {
            AutoFilterOp::Equal => Some("equal"),
            AutoFilterOp::NotEqual => Some("notEqual"),
            AutoFilterOp::GreaterThan => Some("greaterThan"),
            AutoFilterOp::GreaterThanOrEqual => Some("greaterThanOrEqual"),
            AutoFilterOp::LessThan => Some("lessThan"),
            AutoFilterOp::LessThanOrEqual => Some("lessThanOrEqual"),
            AutoFilterOp::Contains => Some("contains"),
            AutoFilterOp::DoesNotContain => Some("doesNotContain"),
            AutoFilterOp::BeginsWith => Some("beginsWith"),
            AutoFilterOp::EndsWith => Some("endsWith"),
            AutoFilterOp::DoesNotBeginWith => Some("doesNotBeginWith"),
            AutoFilterOp::DoesNotEndWith => Some("doesNotEndWith"),
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
enum DoperValue {
    Empty,
    Bool(bool),
    Number(f64),
    Text { value: String, type_known: bool },
    Unknown,
}

#[derive(Debug, Clone)]
struct ParsedDoper {
    op: AutoFilterOp,
    vt: u8,
    value: DoperValue,
}

fn parse_autofilter_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    autofilter_range: Range,
    entry_mode: AutoFilterEntryMode,
    warnings: &mut Vec<String>,
) -> Result<Option<FilterColumn>, String> {
    let data = record.data.as_ref();
    if data.len() < 20 {
        return Err(format!(
            "AUTOFILTER record too short (expected >=20 bytes, got {})",
            data.len()
        ));
    }

    // AUTOFILTER layout (best-effort):
    // - iEntry (2 bytes): column index (0-based column in the sheet)
    // - grbit (2 bytes): flags
    // - DOPER1 (8 bytes)
    // - DOPER2 (8 bytes)
    // - optional strings for string DOPER values (XLUnicodeString)
    let entry = u16::from_le_bytes([data[0], data[1]]) as u32;
    let grbit = u16::from_le_bytes([data[2], data[3]]);

    let col_id = resolve_col_id(entry, autofilter_range, entry_mode)?;

    let join = if (grbit & AUTOFILTER_FLAG_AND) != 0 {
        FilterJoin::All
    } else {
        FilterJoin::Any
    };

    let doper1 = parse_doper(&data[4..12]);
    let doper2 = parse_doper(&data[12..20]);

    // Parse any trailing XLUnicodeString payloads (for string DOPER values). AUTOFILTER strings can
    // be split across CONTINUE records, so we must respect fragment boundaries.
    let mut trailing_strings: Vec<String> = Vec::new();
    if data.len() > 20 {
        let fragments: Vec<&[u8]> = record.fragments().collect();
        if let Some((frag_idx, frag_off)) = locate_fragment_offset(&record.fragment_sizes, 20usize)
        {
            let mut cursor = FragmentCursor::new(&fragments, frag_idx, frag_off);
            for _ in 0..2 {
                if cursor.is_at_end() {
                    break;
                }
                match cursor.read_biff8_unicode_string(codepage) {
                    Ok(s) => trailing_strings.push(s),
                    Err(_) => break,
                }
            }
        }
    }

    let mut str_iter = trailing_strings.into_iter();
    let doper1 = attach_string_to_doper(doper1, &mut str_iter, true, codepage)?;
    let doper2 = attach_string_to_doper(doper2, &mut str_iter, false, codepage)?;

    // Top10 and other advanced filter types are not currently modeled; preserve via raw XML.
    if (grbit & AUTOFILTER_FLAG_TOP10) != 0 {
        let top = if (grbit & AUTOFILTER_FLAG_TOP) != 0 {
            1
        } else {
            0
        };
        let percent = if (grbit & AUTOFILTER_FLAG_PERCENT) != 0 {
            1
        } else {
            0
        };
        let val = match doper1.value {
            DoperValue::Number(n) => n,
            _ => 10.0,
        };
        return Ok(Some(FilterColumn {
            col_id,
            join: FilterJoin::Any,
            criteria: Vec::new(),
            values: Vec::new(),
            raw_xml: vec![format!(
                "<top10 top=\"{top}\" percent=\"{percent}\" val=\"{val}\"/>"
            )],
        }));
    }

    // Best-effort mapping for "between"/"not between" operators. These are represented in OOXML
    // as a pair of comparisons combined with AND/OR; model those directly so they round-trip
    // through XLSX without losing structure.
    if matches!(doper1.op, AutoFilterOp::Between | AutoFilterOp::NotBetween) {
        if let (DoperValue::Number(a), DoperValue::Number(b)) = (&doper1.value, &doper2.value) {
            let (min, max) = if a <= b { (*a, *b) } else { (*b, *a) };
            let (join, criteria) = match doper1.op {
                AutoFilterOp::Between => (
                    FilterJoin::All,
                    vec![
                        FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(min)),
                        FilterCriterion::Number(NumberComparison::LessThanOrEqual(max)),
                    ],
                ),
                AutoFilterOp::NotBetween => (
                    FilterJoin::Any,
                    vec![
                        FilterCriterion::Number(NumberComparison::LessThan(min)),
                        FilterCriterion::Number(NumberComparison::GreaterThan(max)),
                    ],
                ),
                _ => unreachable!(),
            };
            return Ok(Some(FilterColumn {
                col_id,
                join,
                criteria,
                values: Vec::new(),
                raw_xml: Vec::new(),
            }));
        }
    }

    let mut criteria: Vec<FilterCriterion> = Vec::new();
    for doper in [&doper1, &doper2] {
        if let AutoFilterOp::Unknown(code) = doper.op {
            // Unknown/invalid operator codes have been observed in the wild. Only preserve operator
            // values that are valid in OOXML; otherwise skip and surface a warning for corpus triage.
            warnings.push(format!(
                "skipping unknown AUTOFILTER operator code 0x{code:02X} at offset {}",
                record.offset
            ));
            continue;
        }
        if let Some(c) = criterion_from_doper(doper) {
            criteria.push(c);
        }
    }

    // Only emit a FilterColumn when we recovered some criteria or raw XML payload.
    if criteria.is_empty() {
        return Ok(None);
    }

    Ok(Some(FilterColumn {
        col_id,
        join: if criteria.len() > 1 {
            join
        } else {
            FilterJoin::Any
        },
        criteria,
        values: Vec::new(),
        raw_xml: Vec::new(),
    }))
}

fn resolve_col_id(
    entry: u32,
    range: Range,
    entry_mode: AutoFilterEntryMode,
) -> Result<u32, String> {
    let start_col = range.start.col;
    let end_col = range.end.col;
    let width = end_col.saturating_sub(start_col).saturating_add(1);

    match entry_mode {
        AutoFilterEntryMode::Relative => {
            if entry >= width {
                return Err(format!(
                    "AUTOFILTER iEntry {entry} out of range for filter width {width}"
                ));
            }
            Ok(entry)
        }
        AutoFilterEntryMode::Absolute => {
            if entry < start_col || entry > end_col {
                return Err(format!(
                    "AUTOFILTER column {entry} out of range for filter range cols {start_col}..={end_col}"
                ));
            }
            Ok(entry - start_col)
        }
    }
}

fn parse_doper(bytes: &[u8]) -> ParsedDoper {
    // DOPER [MS-XLS 2.5.69] (best-effort).
    //
    // The canonical BIFF8 layout is:
    // - vt (1 byte)
    // - grbit (1 byte)
    // - wOper (2 bytes, little-endian)
    // - operand value (4 bytes, type-dependent; numbers are typically stored as RK)
    //
    // Some producers have been observed to store the operator in the second byte instead of
    // `wOper`; decode both forms best-effort.
    let vt = *bytes.first().unwrap_or(&0);
    let op_u16 = u16::from_le_bytes([*bytes.get(2).unwrap_or(&0), *bytes.get(3).unwrap_or(&0)]);
    let op_byte = *bytes.get(1).unwrap_or(&0);

    let op_code = if op_u16 <= 14 {
        let op = op_u16 as u8;
        // If `wOper` is unset but the second byte looks like an operator, fall back.
        if op == 0 && op_byte != 0 && (op_byte <= 14 || op_byte >= 0x80) {
            op_byte
        } else {
            op
        }
    } else if op_byte != 0 && (op_byte <= 14 || op_byte >= 0x80) {
        op_byte
    } else {
        0
    };

    let value_raw = u32::from_le_bytes([
        *bytes.get(4).unwrap_or(&0),
        *bytes.get(5).unwrap_or(&0),
        *bytes.get(6).unwrap_or(&0),
        *bytes.get(7).unwrap_or(&0),
    ]);

    let op = AutoFilterOp::from_biff_code(op_code);

    ParsedDoper {
        op,
        vt,
        value: decode_doper_value(vt, value_raw),
    }
}

fn decode_doper_value(vt: u8, raw: u32) -> DoperValue {
    // Best-effort decoding:
    // - Numbers are encoded as RK (4 bytes).
    // - Booleans are stored as 0/1 in the low byte.
    // - Strings are stored separately; `attach_string_to_doper` overwrites `value` accordingly.
    match vt {
        0 => DoperValue::Empty,
        // Common boolean encodings: some writers use vt=6, others use vt=11 (VT_BOOL).
        6 | 0x0B => DoperValue::Bool(raw != 0),
        // Treat vt=8 as string when a trailing string is present; otherwise best-effort empty.
        8 => DoperValue::Unknown,
        // Default: treat as RK number.
        _ => DoperValue::Number(decode_rk_number(raw)),
    }
}

fn decode_rk_number(rk: u32) -> f64 {
    let is_integer = (rk & 0x02) != 0;
    let is_x100 = (rk & 0x01) != 0;

    let mut value = if is_integer {
        // Signed 30-bit integer.
        let i = (rk as i32) >> 2;
        i as f64
    } else {
        // High 30 bits of an IEEE754 f64, low 34 bits are zero.
        let bits = (rk & 0xFFFF_FFFC) as u64;
        f64::from_bits(bits << 32)
    };

    if is_x100 {
        value /= 100.0;
    }
    value
}

fn attach_string_to_doper(
    mut doper: ParsedDoper,
    strings: &mut impl Iterator<Item = String>,
    // When the DOPER type is ambiguous, prefer attaching strings to the first condition.
    is_first: bool,
    _codepage: u16,
) -> Result<ParsedDoper, String> {
    // Best-effort: only attach a string when the DOPER type strongly suggests it, or when the
    // DOPER payload looks like "unused" but a trailing string exists.
    let needs_string = matches!(doper.vt, 4 | 8 | 0x10 | 0x11 | 0x12)
        || (is_first && matches!(doper.value, DoperValue::Unknown));
    if !needs_string {
        return Ok(doper);
    }

    if let Some(s) = strings.next() {
        // When we attach a string payload due to an explicit BIFF type tag, treat the type as
        // known (text filter). When we attach it only as a best-effort fallback (e.g. vt is
        // unknown but a trailing string exists), mark it as unknown so we can preserve the
        // operator/value pair as `OpaqueCustom` instead of forcing it into a text equality filter.
        let type_known = matches!(doper.vt, 4 | 8);
        doper.value = DoperValue::Text {
            value: s,
            type_known,
        };
    }

    Ok(doper)
}

fn criterion_from_doper(doper: &ParsedDoper) -> Option<FilterCriterion> {
    if matches!(doper.op, AutoFilterOp::None) {
        return None;
    }

    // Helper to preserve criteria as opaque custom filters **only when the OOXML operator name is
    // known**. Unknown operator strings would produce invalid OOXML.
    let opaque = |op: AutoFilterOp, value: Option<String>| -> Option<FilterCriterion> {
        let operator = op.to_ooxml_operator_name()?.to_string();
        Some(FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
            operator,
            value,
        }))
    };

    match doper.op {
        AutoFilterOp::Equal => match &doper.value {
            DoperValue::Empty => Some(FilterCriterion::Blanks),
            DoperValue::Bool(b) => Some(FilterCriterion::Equals(FilterValue::Bool(*b))),
            DoperValue::Number(n) => Some(FilterCriterion::Equals(FilterValue::Number(*n))),
            DoperValue::Text { value, type_known } => {
                if value.is_empty() {
                    Some(FilterCriterion::Blanks)
                } else if *type_known {
                    Some(FilterCriterion::Equals(FilterValue::Text(value.clone())))
                } else {
                    opaque(AutoFilterOp::Equal, Some(value.clone()))
                }
            }
            DoperValue::Unknown => opaque(AutoFilterOp::Equal, None),
        },
        AutoFilterOp::NotEqual => match &doper.value {
            DoperValue::Empty => Some(FilterCriterion::NonBlanks),
            DoperValue::Bool(b) => opaque(AutoFilterOp::NotEqual, Some(b.to_string())),
            DoperValue::Number(n) => Some(FilterCriterion::Number(NumberComparison::NotEqual(*n))),
            DoperValue::Text { value, .. } => {
                if value.is_empty() {
                    Some(FilterCriterion::NonBlanks)
                } else {
                    opaque(AutoFilterOp::NotEqual, Some(value.clone()))
                }
            }
            DoperValue::Unknown => opaque(AutoFilterOp::NotEqual, None),
        },
        AutoFilterOp::GreaterThan => match doper.value {
            DoperValue::Number(n) => {
                Some(FilterCriterion::Number(NumberComparison::GreaterThan(n)))
            }
            DoperValue::Text { ref value, .. } => {
                opaque(AutoFilterOp::GreaterThan, Some(value.clone()))
            }
            _ => opaque(AutoFilterOp::GreaterThan, None),
        },
        AutoFilterOp::GreaterThanOrEqual => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(
                NumberComparison::GreaterThanOrEqual(n),
            )),
            DoperValue::Text { ref value, .. } => {
                opaque(AutoFilterOp::GreaterThanOrEqual, Some(value.clone()))
            }
            _ => opaque(AutoFilterOp::GreaterThanOrEqual, None),
        },
        AutoFilterOp::LessThan => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(NumberComparison::LessThan(n))),
            DoperValue::Text { ref value, .. } => {
                opaque(AutoFilterOp::LessThan, Some(value.clone()))
            }
            _ => opaque(AutoFilterOp::LessThan, None),
        },
        AutoFilterOp::LessThanOrEqual => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(
                NumberComparison::LessThanOrEqual(n),
            )),
            DoperValue::Text { ref value, .. } => {
                opaque(AutoFilterOp::LessThanOrEqual, Some(value.clone()))
            }
            _ => opaque(AutoFilterOp::LessThanOrEqual, None),
        },
        AutoFilterOp::Contains | AutoFilterOp::BeginsWith | AutoFilterOp::EndsWith => {
            let kind = match doper.op {
                AutoFilterOp::Contains => TextMatchKind::Contains,
                AutoFilterOp::BeginsWith => TextMatchKind::BeginsWith,
                AutoFilterOp::EndsWith => TextMatchKind::EndsWith,
                _ => unreachable!(),
            };
            match &doper.value {
                DoperValue::Text { value, .. } => Some(FilterCriterion::TextMatch(TextMatch {
                    kind,
                    pattern: value.clone(),
                    case_sensitive: false,
                })),
                DoperValue::Empty => Some(FilterCriterion::TextMatch(TextMatch {
                    kind,
                    pattern: String::new(),
                    case_sensitive: false,
                })),
                // If we can't confidently decode a string, preserve as opaque.
                DoperValue::Unknown => opaque(doper.op, None),
                DoperValue::Bool(b) => opaque(doper.op, Some(b.to_string())),
                DoperValue::Number(n) => opaque(doper.op, Some(n.to_string())),
            }
        }
        AutoFilterOp::DoesNotContain
        | AutoFilterOp::DoesNotBeginWith
        | AutoFilterOp::DoesNotEndWith => match &doper.value {
            DoperValue::Text { value, .. } => opaque(doper.op, Some(value.clone())),
            DoperValue::Empty => opaque(doper.op, Some(String::new())),
            DoperValue::Unknown => opaque(doper.op, None),
            DoperValue::Bool(b) => opaque(doper.op, Some(b.to_string())),
            DoperValue::Number(n) => opaque(doper.op, Some(n.to_string())),
        },
        // Between/NotBetween should have been handled earlier (when numeric); otherwise skip.
        AutoFilterOp::Between | AutoFilterOp::NotBetween | AutoFilterOp::Unknown(_) => None,
        AutoFilterOp::None => None,
    }
}

fn locate_fragment_offset(
    fragment_sizes: &[usize],
    global_offset: usize,
) -> Option<(usize, usize)> {
    let mut remaining = global_offset;
    for (idx, &size) in fragment_sizes.iter().enumerate() {
        if remaining < size {
            return Some((idx, remaining));
        }
        remaining = remaining.saturating_sub(size);
    }
    None
}

/// A small cursor for parsing BIFF8 XLUnicodeString values across `CONTINUE` boundaries.
///
/// This is copied from `strings.rs` and extended to support arbitrary starting offsets; the
/// `strings` module currently only exposes a one-shot helper.
struct FragmentCursor<'a> {
    fragments: &'a [&'a [u8]],
    frag_idx: usize,
    offset: usize,
}

impl<'a> FragmentCursor<'a> {
    fn new(fragments: &'a [&'a [u8]], frag_idx: usize, offset: usize) -> Self {
        Self {
            fragments,
            frag_idx,
            offset,
        }
    }

    fn is_at_end(&self) -> bool {
        if self.frag_idx >= self.fragments.len() {
            return true;
        }
        self.remaining_in_fragment() == 0
            && self
                .fragments
                .iter()
                .skip(self.frag_idx.saturating_add(1))
                .all(|f| f.is_empty())
    }

    fn remaining_in_fragment(&self) -> usize {
        self.fragments
            .get(self.frag_idx)
            .map(|f| f.len().saturating_sub(self.offset))
            .unwrap_or(0)
    }

    fn advance_fragment(&mut self) -> Result<(), String> {
        self.frag_idx = self
            .frag_idx
            .checked_add(1)
            .ok_or_else(|| "fragment index overflow".to_string())?;
        self.offset = 0;
        if self.frag_idx >= self.fragments.len() {
            return Err("unexpected end of record".to_string());
        }
        Ok(())
    }

    fn read_u8(&mut self) -> Result<u8, String> {
        loop {
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            if self.offset < frag.len() {
                let b = frag[self.offset];
                self.offset += 1;
                return Ok(b);
            }
            self.advance_fragment()?;
        }
    }

    fn read_u16_le(&mut self) -> Result<u16, String> {
        let lo = self.read_u8()?;
        let hi = self.read_u8()?;
        Ok(u16::from_le_bytes([lo, hi]))
    }

    fn read_exact_from_current(&mut self, n: usize) -> Result<&'a [u8], String> {
        let frag = self
            .fragments
            .get(self.frag_idx)
            .ok_or_else(|| "unexpected end of record".to_string())?;
        let end = self
            .offset
            .checked_add(n)
            .ok_or_else(|| "offset overflow".to_string())?;
        if end > frag.len() {
            return Err("unexpected end of record".to_string());
        }
        let out = &frag[self.offset..end];
        self.offset = end;
        Ok(out)
    }

    fn advance_fragment_in_biff8_string(&mut self, is_unicode: &mut bool) -> Result<(), String> {
        self.advance_fragment()?;
        // When a BIFF8 string spans a CONTINUE boundary, Excel inserts a 1-byte option flags prefix
        // at the start of the continued fragment. The only relevant bit is `fHighByte` (unicode vs
        // compressed).
        let cont_flags = self.read_u8()?;
        *is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
        Ok(())
    }

    fn read_biff8_string_bytes(
        &mut self,
        mut n: usize,
        is_unicode: &mut bool,
    ) -> Result<Vec<u8>, String> {
        // Read `n` canonical bytes from a BIFF8 continued string payload, skipping the 1-byte
        // continuation flags prefix that appears at the start of each continued fragment.
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
        Ok(out)
    }

    fn skip_biff8_string_bytes(
        &mut self,
        mut n: usize,
        is_unicode: &mut bool,
    ) -> Result<(), String> {
        // Skip `n` canonical bytes from a BIFF8 continued string payload, consuming any inserted
        // continuation flags bytes at fragment boundaries.
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
            let take = n.min(available);
            self.offset += take;
            n -= take;
        }
        Ok(())
    }

    fn read_biff8_unicode_string(&mut self, codepage: u16) -> Result<String, String> {
        // XLUnicodeString [MS-XLS 2.5.268]
        let cch = self.read_u16_le()? as usize;
        let flags = self.read_u8()?;

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;

        let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
            let bytes = self.read_biff8_string_bytes(2, &mut is_unicode)?;
            u16::from_le_bytes([bytes[0], bytes[1]]) as usize
        } else {
            0
        };

        let ext_size = if flags & STR_FLAG_EXT != 0 {
            let bytes = self.read_biff8_string_bytes(4, &mut is_unicode)?;
            u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize
        } else {
            0
        };

        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first byte is option
                // flags for the continued segment (fHighByte).
                self.advance_fragment_in_biff8_string(&mut is_unicode)?;
                continue;
            }

            let bytes_per_char = if is_unicode { 2 } else { 1 };
            let available_bytes = self.remaining_in_fragment();
            let available_chars = available_bytes / bytes_per_char;
            if available_chars == 0 {
                return Err("string continuation split mid-character".to_string());
            }

            let take_chars = remaining_chars.min(available_chars);
            let take_bytes = take_chars * bytes_per_char;
            let bytes = self.read_exact_from_current(take_bytes)?;

            if is_unicode {
                let mut u16s = Vec::with_capacity(take_chars);
                for chunk in bytes.chunks_exact(2) {
                    u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                out.push_str(&String::from_utf16_lossy(&u16s));
            } else {
                out.push_str(&strings::decode_ansi(codepage, bytes));
            }

            remaining_chars -= take_chars;
        }

        let richtext_bytes = richtext_runs
            .checked_mul(4)
            .ok_or_else(|| "rich text run count overflow".to_string())?;
        let extra_len = richtext_bytes
            .checked_add(ext_size)
            .ok_or_else(|| "string ext payload length overflow".to_string())?;
        self.skip_biff8_string_bytes(extra_len, &mut is_unicode)?;

        Ok(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::autofilter::FilterCriterion;
    use formula_model::{CellRef, Range};

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn bof_worksheet() -> Vec<u8> {
        // Minimal BIFF8 BOF payload: version (0x0600) + dt (worksheet=0x0010) + padding.
        let mut out = Vec::new();
        out.extend_from_slice(&0x0600u16.to_le_bytes());
        out.extend_from_slice(&0x0010u16.to_le_bytes());
        out.extend_from_slice(&[0u8; 12]);
        out
    }

    fn xl_unicode_string_compressed(s: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(s.len() as u16).to_le_bytes());
        out.push(0); // flags (compressed)
        out.extend_from_slice(s.as_bytes());
        out
    }

    fn xl_unicode_string_unicode(s: &str) -> Vec<u8> {
        let mut out = Vec::new();
        let utf16le: Vec<u8> = s
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<u8>>();
        let cch: u16 = (utf16le.len() / 2).try_into().expect("test string fits in u16");
        out.extend_from_slice(&cch.to_le_bytes());
        out.push(STR_FLAG_HIGH_BYTE); // flags (uncompressed/unicode)
        out.extend_from_slice(&utf16le);
        out
    }

    fn rk_number(n: i32) -> u32 {
        // Encode a signed integer RK value.
        ((n as u32) << 2) | 0x02
    }

    #[test]
    fn parses_equals_text_filter() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 2)); // A..C

        // AUTOFILTER col=1 (B), grbit=0, doper1 = (vt=8 string, wOper=3 equal), doper2 unused.
        let mut af = Vec::new();
        af.extend_from_slice(&1u16.to_le_bytes()); // col
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1.
        af.push(8); // vt (string)
        af.push(0); // grbit
        af.extend_from_slice(&3u16.to_le_bytes()); // wOper (equal)
        af.extend_from_slice(&0u32.to_le_bytes()); // value_raw

        // DOPER2 (unused).
        af.extend_from_slice(&[0u8; 8]);

        // Trailing string.
        af.extend_from_slice(&xl_unicode_string_compressed("Alice"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(col.col_id, 1);
        assert_eq!(col.join, FilterJoin::Any);
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("Alice".into()))]
        );
    }

    #[test]
    fn parses_numeric_comparison_and_join() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        // AUTOFILTER col=0 (A), grbit=AND, doper1 >=2, doper2 <=5.
        let mut af = Vec::new();
        af.extend_from_slice(&0u16.to_le_bytes()); // col
        af.extend_from_slice(&AUTOFILTER_FLAG_AND.to_le_bytes()); // grbit (AND)

        // DOPER1: vt=1 (number), wOper=7 (>=), rk=2.
        af.push(1); // vt
        af.push(0); // grbit
        af.extend_from_slice(&7u16.to_le_bytes()); // wOper
        af.extend_from_slice(&rk_number(2).to_le_bytes());

        // DOPER2: vt=1 (number), wOper=8 (<=), rk=5.
        af.push(1); // vt
        af.push(0); // grbit
        af.extend_from_slice(&8u16.to_le_bytes()); // wOper
        af.extend_from_slice(&rk_number(5).to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(col.col_id, 0);
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(2.0)),
                FilterCriterion::Number(NumberComparison::LessThanOrEqual(5.0))
            ]
        );
    }

    #[test]
    fn parses_continued_string_across_continue_records() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        // Build AUTOFILTER record with a string that will be split across CONTINUE.
        let mut af_full = Vec::new();
        af_full.extend_from_slice(&0u16.to_le_bytes()); // col
        af_full.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1.
        af_full.push(8); // vt string
        af_full.push(0); // grbit
        af_full.extend_from_slice(&3u16.to_le_bytes()); // wOper equal
        af_full.extend_from_slice(&0u32.to_le_bytes());
        // DOPER2 unused.
        af_full.extend_from_slice(&[0u8; 8]);

        // XLUnicodeString "ABCDE".
        let s = "ABCDE";
        let mut str_bytes = Vec::new();
        str_bytes.extend_from_slice(&(s.len() as u16).to_le_bytes());
        str_bytes.push(0); // flags (compressed)
        str_bytes.extend_from_slice(s.as_bytes());
        af_full.extend_from_slice(&str_bytes);

        // Split the logical AUTOFILTER payload so that the string's character bytes span fragments.
        // Keep the header + part of the character data in the first record; rest in CONTINUE with
        // the required "continued segment flags" byte.
        let string_start = 20; // after header + 2 dopers
        let split_at = string_start + 3 + 2; // header (3) + 2 chars ("AB")

        let first_payload = &af_full[..split_at];
        let remaining_chars = &af_full[split_at..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment flags (compressed)
        continue_payload.extend_from_slice(remaining_chars);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, first_payload),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("ABCDE".into()))]
        );
    }

    #[test]
    fn parses_continued_richtext_string_with_crun_split_across_continue_records() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        // Build AUTOFILTER record with a rich-text string that will be split across CONTINUE such
        // that the `cRun` field spans the boundary.
        let mut af_full = Vec::new();
        af_full.extend_from_slice(&0u16.to_le_bytes()); // col
        af_full.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1.
        af_full.push(8); // vt string
        af_full.push(0); // grbit
        af_full.extend_from_slice(&3u16.to_le_bytes()); // wOper equal
        af_full.extend_from_slice(&0u32.to_le_bytes());
        // DOPER2 unused.
        af_full.extend_from_slice(&[0u8; 8]);

        // XLUnicodeString "ABCDE" with richtext flag set.
        let s = "ABCDE";
        let mut str_bytes = Vec::new();
        str_bytes.extend_from_slice(&(s.len() as u16).to_le_bytes());
        str_bytes.push(STR_FLAG_RICH_TEXT); // flags (compressed + rich text)
        str_bytes.extend_from_slice(&1u16.to_le_bytes()); // cRun=1
        str_bytes.extend_from_slice(s.as_bytes());
        str_bytes.extend_from_slice(&[0x11, 0x22, 0x33, 0x44]); // rgRun (4 bytes)
        af_full.extend_from_slice(&str_bytes);

        // Split such that we cut between the two bytes of `cRun`. The CONTINUE fragment begins with
        // the required continued-segment option flags byte.
        let string_start = 20usize;
        let split_at = string_start + 3 + 1; // header (3) + first byte of cRun
        let first_payload = &af_full[..split_at];
        let remaining = &af_full[split_at..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, first_payload),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("ABCDE".into()))]
        );
    }

    #[test]
    fn parses_continued_ext_string_with_ext_payload_split_across_continue_records() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let mut af_full = Vec::new();
        af_full.extend_from_slice(&0u16.to_le_bytes()); // col
        af_full.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: string equals.
        af_full.push(8); // vt string
        af_full.push(0); // grbit
        af_full.extend_from_slice(&3u16.to_le_bytes()); // wOper equal
        af_full.extend_from_slice(&0u32.to_le_bytes());

        // DOPER2: another string equals.
        af_full.push(8); // vt string
        af_full.push(0); // grbit
        af_full.extend_from_slice(&3u16.to_le_bytes()); // wOper equal
        af_full.extend_from_slice(&0u32.to_le_bytes());

        // First XLUnicodeString "abc" with ext payload.
        let s1 = "abc";
        let ext = [0xDEu8, 0xAD, 0xBE, 0xEF];
        let mut str1_bytes = Vec::new();
        str1_bytes.extend_from_slice(&(s1.len() as u16).to_le_bytes());
        str1_bytes.push(STR_FLAG_EXT); // flags (compressed + ext)
        str1_bytes.extend_from_slice(&(ext.len() as u32).to_le_bytes()); // cbExtRst
        str1_bytes.extend_from_slice(s1.as_bytes());
        str1_bytes.extend_from_slice(&ext);

        // Second XLUnicodeString "Z" (simple).
        let s2 = "Z";
        let str2_bytes = xl_unicode_string_compressed(s2);

        af_full.extend_from_slice(&str1_bytes);
        af_full.extend_from_slice(&str2_bytes);

        // Split so the first string's ext payload spans a CONTINUE record (after 2 ext bytes).
        let string_start = 20usize;
        let split_at = string_start + 3 + 4 + s1.len() + 2; // header(3) + cbExtRst(4) + chars + 2 ext bytes
        let first_payload = &af_full[..split_at];
        let remaining = &af_full[split_at..];

        let mut continue_payload = Vec::new();
        continue_payload.push(0); // continued segment flags (compressed)
        continue_payload.extend_from_slice(remaining);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, first_payload),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Equals(FilterValue::Text(s1.into())),
                FilterCriterion::Equals(FilterValue::Text(s2.into()))
            ]
        );
    }

    #[test]
    fn parses_continued_unicode_string_across_continue_records_with_high_byte_flag() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let mut af_full = Vec::new();
        af_full.extend_from_slice(&0u16.to_le_bytes()); // col
        af_full.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: string equals.
        af_full.push(4); // vt string
        af_full.push(3); // op equal
        af_full.extend_from_slice(&0u16.to_le_bytes());
        af_full.extend_from_slice(&0u32.to_le_bytes());
        // DOPER2 unused.
        af_full.extend_from_slice(&[0u8; 8]);

        let s = "Hello";
        af_full.extend_from_slice(&xl_unicode_string_unicode(s));

        // Split mid-string so the UTF-16 character bytes span a CONTINUE record.
        let string_start = 20usize;
        let split_at = string_start + 3 + 4; // header (3) + 2 UTF-16 chars (4 bytes)
        let first_payload = &af_full[..split_at];
        let remaining_bytes = &af_full[split_at..];

        let mut continue_payload = Vec::new();
        continue_payload.push(STR_FLAG_HIGH_BYTE); // continued segment flags (unicode)
        continue_payload.extend_from_slice(remaining_bytes);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, first_payload),
            record(records::RECORD_CONTINUE, &continue_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::Equals(FilterValue::Text(s.into()))]
        );
    }

    #[test]
    fn truncated_autofilter_payload_emits_warning_and_skips_record() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let stream = [
            record(RECORD_AUTOFILTER, &[1, 2, 3]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");
        assert!(parsed.filter_columns.is_empty());
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("failed to decode AUTOFILTER record") && w.contains("offset 0")),
            "expected warning with record offset, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn unknown_operator_codes_fall_back_to_opaque_custom_or_warn_and_skip() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let mut af = Vec::new();
        af.extend_from_slice(&0u16.to_le_bytes()); // col
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: vt=string, op=doesNotContain (0x0C).
        af.push(4);
        af.push(12);
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&0u32.to_le_bytes());

        // DOPER2: vt=string, op=unknown (0xFF) => warning + skipped.
        af.push(4);
        af.push(0xFF);
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&0u32.to_le_bytes());

        af.extend_from_slice(&xl_unicode_string_compressed("foo"));
        af.extend_from_slice(&xl_unicode_string_compressed("bar"));

        let stream = [
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
                operator: "doesNotContain".to_string(),
                value: Some("foo".to_string())
            })]
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("unknown AUTOFILTER operator code 0xFF")
                    && w.contains("offset 0")),
            "expected unknown-op warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_entry_index_relative_to_autofilter_range_start() {
        // AutoFilter range starts at column D (index 3). AUTOFILTER iEntry uses a 0-based index
        // within the range in BIFF8, so entry 0 should map to colId=0 (column D).
        let range = Range::new(CellRef::new(0, 3), CellRef::new(10, 5)); // D..F

        let mut af = Vec::new();
        af.extend_from_slice(&0u16.to_le_bytes()); // iEntry (relative)
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: string equals "X".
        af.push(8);
        af.push(0);
        af.extend_from_slice(&3u16.to_le_bytes());
        af.extend_from_slice(&0u32.to_le_bytes());
        af.extend_from_slice(&[0u8; 8]); // DOPER2 unused
        af.extend_from_slice(&xl_unicode_string_compressed("X"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        assert_eq!(parsed.filter_columns[0].col_id, 0);
    }

    #[test]
    fn parses_entry_as_absolute_col_index_when_it_matches_range() {
        // Some writers store the AUTOFILTER "entry" field as an absolute worksheet column.
        // Validate the best-effort detection path when the AutoFilter range does not start at A.
        let range = Range::new(CellRef::new(0, 3), CellRef::new(10, 5)); // D..F (start_col=3)

        let mut af = Vec::new();
        af.extend_from_slice(&3u16.to_le_bytes()); // entry encoded as absolute column D
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: string equals "X".
        af.push(8);
        af.push(0);
        af.extend_from_slice(&3u16.to_le_bytes());
        af.extend_from_slice(&0u32.to_le_bytes());
        af.extend_from_slice(&[0u8; 8]); // DOPER2 unused
        af.extend_from_slice(&xl_unicode_string_compressed("X"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        assert_eq!(parsed.filter_columns[0].col_id, 0);
    }

    #[test]
    fn parses_between_operator_into_two_numeric_comparisons() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let mut af = Vec::new();
        af.extend_from_slice(&0u16.to_le_bytes()); // entry
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: between 2..5 (min).
        af.push(1); // vt
        af.push(1); // op between
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(2).to_le_bytes());
        // DOPER2: max value, op unused.
        af.push(1); // vt
        af.push(0); // op none
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(5).to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        let col = &parsed.filter_columns[0];
        assert_eq!(col.join, FilterJoin::All);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::GreaterThanOrEqual(2.0)),
                FilterCriterion::Number(NumberComparison::LessThanOrEqual(5.0)),
            ]
        );
    }

    #[test]
    fn parses_not_between_operator_into_or_comparisons() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 0)); // A

        let mut af = Vec::new();
        af.extend_from_slice(&0u16.to_le_bytes()); // entry
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1: not between 2..5 (min).
        af.push(1); // vt
        af.push(2); // op not between
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(2).to_le_bytes());
        // DOPER2: max value, op unused.
        af.push(1); // vt
        af.push(0); // op none
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(5).to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_sheet_autofilter_criteria(&stream, 0, BiffVersion::Biff8, 1252, range)
                .expect("parse");

        let col = &parsed.filter_columns[0];
        assert_eq!(col.join, FilterJoin::Any);
        assert_eq!(
            col.criteria,
            vec![
                FilterCriterion::Number(NumberComparison::LessThan(2.0)),
                FilterCriterion::Number(NumberComparison::GreaterThan(5.0)),
            ]
        );
    }
}
