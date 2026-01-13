use std::collections::HashMap;

use formula_model::autofilter::{
    FilterColumn, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueCustomFilter,
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
            RECORD_AUTOFILTER => match parse_autofilter_record(&record, codepage, autofilter_range)
            {
                Ok(Some(col)) => {
                    cols.insert(col.col_id, col);
                }
                Ok(None) => {}
                Err(err) => out.warnings.push(format!(
                    "failed to decode AUTOFILTER record at offset {}: {err}",
                    record.offset
                )),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let mut filter_columns: Vec<FilterColumn> = cols.into_values().collect();
    filter_columns.sort_by_key(|c| c.col_id);
    out.filter_columns = filter_columns;

    Ok(out)
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
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
enum DoperValue {
    Empty,
    Bool(bool),
    Number(f64),
    Text {
        value: String,
        type_known: bool,
    },
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
) -> Result<Option<FilterColumn>, String> {
    let data = record.data.as_ref();
    if data.len() < 20 {
        return Err(format!("AUTOFILTER record too short (expected >=20 bytes, got {})", data.len()));
    }

    // AUTOFILTER layout (best-effort):
    // - iEntry (2 bytes): column index (0-based column in the sheet)
    // - grbit (2 bytes): flags
    // - DOPER1 (8 bytes)
    // - DOPER2 (8 bytes)
    // - optional strings for string DOPER values (XLUnicodeString)
    let col = u16::from_le_bytes([data[0], data[1]]) as u32;
    let grbit = u16::from_le_bytes([data[2], data[3]]);

    if col < autofilter_range.start.col || col > autofilter_range.end.col {
        return Ok(None);
    }
    let col_id = col - autofilter_range.start.col;

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
        if let Some((frag_idx, frag_off)) =
            locate_fragment_offset(&record.fragment_sizes, 20usize)
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
        let top = if (grbit & AUTOFILTER_FLAG_TOP) != 0 { 1 } else { 0 };
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

    let mut criteria: Vec<FilterCriterion> = Vec::new();
    if let Some(c) = criterion_from_doper(&doper1) {
        criteria.push(c);
    }
    if let Some(c) = criterion_from_doper(&doper2) {
        criteria.push(c);
    }

    // Only emit a FilterColumn when we recovered some criteria or raw XML payload.
    if criteria.is_empty() {
        return Ok(None);
    }

    Ok(Some(FilterColumn {
        col_id,
        join: if criteria.len() > 1 { join } else { FilterJoin::Any },
        criteria,
        values: Vec::new(),
        raw_xml: Vec::new(),
    }))
}

fn parse_doper(bytes: &[u8]) -> ParsedDoper {
    // DOPER [MS-XLS 2.5.69] (best-effort).
    //
    // The exact vt/op encoding differs across BIFF producers and record variants. We treat the
    // first two bytes as (vt, op) but attempt to auto-detect swapped layouts.
    let b0 = *bytes.first().unwrap_or(&0);
    let b1 = *bytes.get(1).unwrap_or(&0);

    let (vt, op_code) = classify_doper_header(b0, b1);

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

fn classify_doper_header(b0: u8, b1: u8) -> (u8, u8) {
    let op_set = b0 <= 8;
    let op_set_b1 = b1 <= 8;

    // Prefer treating the byte that looks like an operator (0..=8) as the operator.
    if op_set && !op_set_b1 {
        return (b1, b0);
    }
    if op_set_b1 && !op_set {
        return (b0, b1);
    }

    // Ambiguous: default to (vt=b0, op=b1) which matches the [MS-XLS] ordering.
    (b0, b1)
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
        let type_known = matches!(doper.vt, 4);
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

    // Helper to preserve unsupported criteria as opaque custom filter operators.
    let opaque = |op: AutoFilterOp, value: Option<String>| -> FilterCriterion {
        FilterCriterion::OpaqueCustom(OpaqueCustomFilter {
            operator: op
                .to_ooxml_operator_name()
                .unwrap_or_else(|| "unknown")
                .to_string(),
            value,
        })
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
                    Some(opaque(AutoFilterOp::Equal, Some(value.clone())))
                }
            }
            DoperValue::Unknown => Some(opaque(AutoFilterOp::Equal, None)),
        },
        AutoFilterOp::NotEqual => match &doper.value {
            DoperValue::Empty => Some(FilterCriterion::NonBlanks),
            DoperValue::Bool(b) => Some(opaque(AutoFilterOp::NotEqual, Some(b.to_string()))),
            DoperValue::Number(n) => Some(FilterCriterion::Number(NumberComparison::NotEqual(*n))),
            DoperValue::Text { value, .. } => {
                if value.is_empty() {
                    Some(FilterCriterion::NonBlanks)
                } else {
                    Some(opaque(AutoFilterOp::NotEqual, Some(value.clone())))
                }
            }
            DoperValue::Unknown => Some(opaque(AutoFilterOp::NotEqual, None)),
        },
        AutoFilterOp::GreaterThan => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(NumberComparison::GreaterThan(n))),
            DoperValue::Text { ref value, .. } => {
                Some(opaque(AutoFilterOp::GreaterThan, Some(value.clone())))
            }
            _ => Some(opaque(AutoFilterOp::GreaterThan, None)),
        },
        AutoFilterOp::GreaterThanOrEqual => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(
                NumberComparison::GreaterThanOrEqual(n),
            )),
            DoperValue::Text { ref value, .. } => Some(opaque(
                AutoFilterOp::GreaterThanOrEqual,
                Some(value.clone()),
            )),
            _ => Some(opaque(AutoFilterOp::GreaterThanOrEqual, None)),
        },
        AutoFilterOp::LessThan => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(NumberComparison::LessThan(n))),
            DoperValue::Text { ref value, .. } => Some(opaque(AutoFilterOp::LessThan, Some(value.clone()))),
            _ => Some(opaque(AutoFilterOp::LessThan, None)),
        },
        AutoFilterOp::LessThanOrEqual => match doper.value {
            DoperValue::Number(n) => Some(FilterCriterion::Number(
                NumberComparison::LessThanOrEqual(n),
            )),
            DoperValue::Text { ref value, .. } => Some(opaque(
                AutoFilterOp::LessThanOrEqual,
                Some(value.clone()),
            )),
            _ => Some(opaque(AutoFilterOp::LessThanOrEqual, None)),
        },
        AutoFilterOp::Between | AutoFilterOp::NotBetween | AutoFilterOp::Unknown(_) => {
            Some(opaque(doper.op, None))
        }
        AutoFilterOp::None => None,
    }
}

fn locate_fragment_offset(fragment_sizes: &[usize], global_offset: usize) -> Option<(usize, usize)> {
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

    fn read_u32_le(&mut self) -> Result<u32, String> {
        let b0 = self.read_u8()?;
        let b1 = self.read_u8()?;
        let b2 = self.read_u8()?;
        let b3 = self.read_u8()?;
        Ok(u32::from_le_bytes([b0, b1, b2, b3]))
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

    fn skip_bytes(&mut self, mut n: usize) -> Result<(), String> {
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
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

        let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
            self.read_u16_le()? as usize
        } else {
            0
        };

        let ext_size = if flags & STR_FLAG_EXT != 0 {
            self.read_u32_le()? as usize
        } else {
            0
        };

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first byte is option
                // flags for the continued segment (fHighByte).
                self.advance_fragment()?;
                let cont_flags = self.read_u8()?;
                is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
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
        self.skip_bytes(richtext_bytes + ext_size)?;

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

    fn rk_number(n: i32) -> u32 {
        // Encode a signed integer RK value.
        ((n as u32) << 2) | 0x02
    }

    #[test]
    fn parses_equals_text_filter() {
        let range = Range::new(CellRef::new(0, 0), CellRef::new(10, 2)); // A..C

        // AUTOFILTER col=1 (B), grbit=0, doper1 = (vt=4 string, op=3 equal), doper2 unused.
        let mut af = Vec::new();
        af.extend_from_slice(&1u16.to_le_bytes()); // col
        af.extend_from_slice(&0u16.to_le_bytes()); // grbit

        // DOPER1.
        af.push(4); // vt (string)
        af.push(3); // op (equal)
        af.extend_from_slice(&0u16.to_le_bytes()); // reserved
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

        let parsed = parse_biff_sheet_autofilter_criteria(
            &stream,
            0,
            BiffVersion::Biff8,
            1252,
            range,
        )
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

        // DOPER1: vt=1 (number), op=7 (>=), rk=2.
        af.push(1);
        af.push(7);
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(2).to_le_bytes());

        // DOPER2: vt=1 (number), op=8 (<=), rk=5.
        af.push(1);
        af.push(8);
        af.extend_from_slice(&0u16.to_le_bytes());
        af.extend_from_slice(&rk_number(5).to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_worksheet()),
            record(RECORD_AUTOFILTER, &af),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_autofilter_criteria(
            &stream,
            0,
            BiffVersion::Biff8,
            1252,
            range,
        )
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
        af_full.push(4); // vt string
        af_full.push(3); // op equal
        af_full.extend_from_slice(&0u16.to_le_bytes());
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

        let parsed = parse_biff_sheet_autofilter_criteria(
            &stream,
            0,
            BiffVersion::Biff8,
            1252,
            range,
        )
        .expect("parse");

        assert_eq!(parsed.filter_columns.len(), 1);
        let col = &parsed.filter_columns[0];
        assert_eq!(
            col.criteria,
            vec![FilterCriterion::Equals(FilterValue::Text("ABCDE".into()))]
        );
    }
}
