use formula_model::autofilter::{SortCondition, SortState};
use formula_model::{CellRef, Range, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

use super::records;

// Worksheet substream record ids.
// See [MS-XLS] 2.4.261 (SORT).
const RECORD_SORT: u16 = 0x0090;
/// ContinueFrt12 [MS-XLS] 2.4.?? (Future Record Type continuation; BIFF8 only)
///
/// Sort12/SortData12 payloads can span multiple BIFF records via one or more `ContinueFrt12`
/// fragments. The record begins with an `FrtHeader`; bytes after the header should be appended to
/// the previous FRT record payload.
const RECORD_CONTINUEFRT12: u16 = 0x087F;

// BIFF8 "future record" variants used by newer Excel versions.
// See [MS-XLS] 2.4.278 (Sort12) and 2.4.277 (SortData12).
//
// Note: Some producers may emit Sort12/SortData12 instead of the classic SORT record.
// These records are Future Record Type (FRT) records and begin with an `FrtHeader`. Excel writers
// typically use `record_id == FrtHeader.rt`, but we key off `rt` to be robust.
const RT_SORT12: u16 = 0x0890;
const RT_SORTDATA12: u16 = 0x0895;
// Some producers appear to use alternative `rt` values for Sort12/SortData12. Accept these as
// aliases in best-effort decoding.
const RT_SORT12_ALT: u16 = 0x0880;
const RT_SORTDATA12_ALT: u16 = 0x0881;

// BIFF record id range used by Excel for future records (FRT).
const RECORD_FRT_MIN: u16 = 0x0850;
const RECORD_FRT_MAX: u16 = 0x08FF;

// BIFF8 `.xls` worksheets are limited to 256 columns, but some producers use `0x3FFF` as a
// sentinel "max column". Masking to 8 bits maps that to `0x00FF` (IV), matching Excel's limits.
const BIFF8_COL_INDEX_MASK: u16 = 0x00FF;

// Sort option flags.
//
// [MS-XLS] uses a `grbit` field with multiple sort options. The exact bit assignments differ
// between classic and future records; we only need the "header row present" signal.
//
// Empirically, Excel writers have used both low and high bits for the header flag across record
// variants. Accept either to maximize interoperability.
const SORT_GRBIT_HEADER_LOW: u16 = 0x0001;
const SORT_GRBIT_HEADER_HIGH: u16 = 0x0010;

#[derive(Debug, Default)]
pub(crate) struct ParsedSheetSortState {
    pub(crate) sort_state: Option<SortState>,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort parse of worksheet sort state relevant to an AutoFilter range.
///
/// This scans the worksheet substream for sort-related records and attempts to recover the
/// classic BIFF8 `SORT` record into a [`SortState`] payload.
///
/// Malformed/truncated records are surfaced as warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_sort_state(
    workbook_stream: &[u8],
    start: usize,
    auto_filter_range: Range,
) -> Result<ParsedSheetSortState, String> {
    let mut out = ParsedSheetSortState::default();

    let mut pending_frt_sort: Option<PendingFrtSort> = None;

    // SORT and BIFF8 future record types can legally be split across `CONTINUE` records. Use the
    // logical iterator so we can reassemble those payloads before decoding.
    let allows_continuation = |record_id: u16| {
        record_id == RECORD_SORT || (RECORD_FRT_MIN..=RECORD_FRT_MAX).contains(&record_id)
    };
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(r) => r,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            flush_pending_frt_sort(pending_frt_sort.take(), auto_filter_range, &mut out);
            break;
        }

        // Flush a pending Sort12/SortData12 record before processing any non-continuation record.
        // This ensures `ContinueFrt12` fragments are associated only with the immediately preceding
        // FRT record.
        if record.record_id != RECORD_CONTINUEFRT12 {
            flush_pending_frt_sort(pending_frt_sort.take(), auto_filter_range, &mut out);
        }

        match record.record_id {
            RECORD_SORT => {
                if let Some(sort_state) = parse_sort_record_best_effort(
                    record.data.as_ref(),
                    record.offset,
                    auto_filter_range,
                    &mut out.warnings,
                ) {
                    out.sort_state = Some(sort_state);
                }
            }
            RECORD_CONTINUEFRT12 => {
                let Some(pending) = pending_frt_sort.as_mut() else {
                    continue;
                };
                let data = record.data.as_ref();
                let payload = parse_frt_header(data)
                    .map(|(_, p)| p)
                    .unwrap_or(data);
                if payload.is_empty() {
                    continue;
                }
                if pending.fragments >= records::MAX_LOGICAL_RECORD_FRAGMENTS {
                    let kind = match pending.rt {
                        RT_SORTDATA12 | RT_SORTDATA12_ALT => "unsupported SortData12",
                        _ => "unsupported Sort12",
                    };
                    push_warning_once(&mut out.warnings, kind);
                    pending_frt_sort = None;
                    continue;
                }
                if pending
                    .payload
                    .len()
                    .saturating_add(payload.len())
                    > records::MAX_LOGICAL_RECORD_BYTES
                {
                    let kind = match pending.rt {
                        RT_SORTDATA12 | RT_SORTDATA12_ALT => "unsupported SortData12",
                        _ => "unsupported Sort12",
                    };
                    push_warning_once(&mut out.warnings, kind);
                    pending_frt_sort = None;
                    continue;
                }
                pending.payload.extend_from_slice(payload);
                pending.fragments = pending.fragments.saturating_add(1);
            }
            id if (RECORD_FRT_MIN..=RECORD_FRT_MAX).contains(&id) => {
                let data = record.data.as_ref();
                let (rt, payload) = parse_frt_header(data).unwrap_or((id, data));
                match rt {
                    RT_SORT12 | RT_SORT12_ALT => {
                        pending_frt_sort = Some(PendingFrtSort {
                            rt,
                            record_offset: record.offset,
                            payload: payload.to_vec(),
                            fragments: 1,
                        });
                    }
                    RT_SORTDATA12 | RT_SORTDATA12_ALT => {
                        pending_frt_sort = Some(PendingFrtSort {
                            rt,
                            record_offset: record.offset,
                            payload: payload.to_vec(),
                            fragments: 1,
                        });
                    }
                    _ => {}
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    flush_pending_frt_sort(pending_frt_sort.take(), auto_filter_range, &mut out);
    Ok(out)
}

#[derive(Debug)]
struct PendingFrtSort {
    rt: u16,
    record_offset: usize,
    payload: Vec<u8>,
    fragments: usize,
}

fn flush_pending_frt_sort(
    pending: Option<PendingFrtSort>,
    auto_filter_range: Range,
    out: &mut ParsedSheetSortState,
) {
    let Some(pending) = pending else {
        return;
    };

    if let Some(sort_state) = parse_sort12_like_payload_best_effort(
        &pending.payload,
        pending.record_offset,
        auto_filter_range,
    ) {
        out.sort_state = Some(sort_state);
        return;
    }

    if !payload_has_relevant_ref(&pending.payload, auto_filter_range) {
        return;
    }

    match pending.rt {
        RT_SORT12 | RT_SORT12_ALT => {
            push_warning_once(&mut out.warnings, "unsupported Sort12")
        }
        RT_SORTDATA12 | RT_SORTDATA12_ALT => {
            push_warning_once(&mut out.warnings, "unsupported SortData12")
        }
        _ => {}
    }
}

#[derive(Debug, Clone)]
struct ParsedSortRecord {
    range: Range,
    grbit: u16,
    keys: Vec<SortKey>,
}

#[derive(Debug, Clone, Copy)]
struct SortKey {
    col_raw: u16,
    descending: bool,
}

fn parse_sort_record_best_effort(
    data: &[u8],
    record_offset: usize,
    auto_filter_range: Range,
    warnings: &mut Vec<String>,
) -> Option<SortState> {
    // The canonical BIFF8 `SORT` record layout is 24 bytes and is by far the most common
    // encoding produced by Excel.
    //
    // If it looks like the canonical layout, treat it as authoritative: do not attempt to decode
    // alternative layouts from the same payload (which risks inventing bogus key columns when the
    // record is present but contains no usable keys).
    if let Some(parsed) = parse_sort_record_canonical_24(data) {
        if range_matches_or_contained_by(auto_filter_range, parsed.range) {
            if let Some(sort_state) = build_sort_state_from_parsed(parsed, auto_filter_range) {
                return Some(sort_state);
            }
            warnings.push(format!(
                "failed to decode SORT record at offset {record_offset}: no usable sort keys"
            ));
        }
        return None;
    }

    let candidates = [
        parse_sort_record_grbit_then_count(data),
        parse_sort_record_count_then_grbit(data),
        parse_sort_record_fixed_3_keys(data),
    ];

    let mut saw_relevant_range = false;

    for parsed in candidates.into_iter().flatten() {
        if !range_matches_or_contained_by(auto_filter_range, parsed.range) {
            continue;
        }

        saw_relevant_range = true;

        if let Some(sort_state) = build_sort_state_from_parsed(parsed, auto_filter_range) {
            return Some(sort_state);
        }
    }

    if saw_relevant_range {
        warnings.push(format!(
            "failed to decode SORT record at offset {record_offset}: no usable sort keys"
        ));
    }

    None
}

/// Parse an `FrtHeader` structure and return `(rt, payload_after_header)`.
///
/// `FrtHeader` is an 8-byte structure: `rt` (u16), `grbitFrt` (u16), and a reserved u32.
fn parse_frt_header(data: &[u8]) -> Option<(u16, &[u8])> {
    if data.len() < 8 {
        return None;
    }
    let rt = u16::from_le_bytes([data[0], data[1]]);
    Some((rt, &data[8..]))
}

fn push_warning_once(warnings: &mut Vec<String>, msg: &'static str) {
    if warnings.iter().any(|w| w == msg) {
        return;
    }
    warnings.push(msg.to_string());
}

/// Returns true if the payload appears to contain a Ref8 range that matches or is contained by the
/// AutoFilter range.
fn payload_has_relevant_ref(payload: &[u8], auto_filter_range: Range) -> bool {
    // Try a few common starting offsets for embedded Ref8 ranges within future-record payloads.
    for start in [0usize, 2, 4, 6, 8, 10, 12, 14, 16] {
        let Some(slice) = payload.get(start..) else {
            continue;
        };
        let Some(range) = parse_ref8(slice) else {
            continue;
        };
        if range_matches_or_contained_by(auto_filter_range, range) {
            return true;
        }
    }
    false
}

/// Best-effort decode of a Sort12/SortData12 payload (after the `FrtHeader`).
///
/// The exact on-disk structures for these future records are complex. As a conservative fallback,
/// attempt to interpret the payload as a classic BIFF8 `SORT` record starting at a handful of
/// common offsets. This recovers basic column-order sort state for some Excel-generated `.xls`
/// files.
fn parse_sort12_like_payload_best_effort(
    payload: &[u8],
    record_offset: usize,
    auto_filter_range: Range,
) -> Option<SortState> {
    for start in [0usize, 2, 4, 6, 8, 10, 12, 14, 16] {
        let Some(slice) = payload.get(start..) else {
            continue;
        };
        // Require a valid Ref8 at the start to avoid inventing bogus sort state.
        if parse_ref8(slice).is_none() {
            continue;
        }
        let mut tmp_warnings = Vec::new();
        if let Some(state) =
            parse_sort_record_best_effort(slice, record_offset, auto_filter_range, &mut tmp_warnings)
        {
            return Some(state);
        }
    }
    None
}

fn parse_sort_record_canonical_24(data: &[u8]) -> Option<ParsedSortRecord> {
    // Canonical BIFF8 SORT layout (24 bytes):
    // - Ref8U (rwFirst, rwLast, colFirst, colLast): 8 bytes
    // - grbit: 2 bytes
    // - cKeys: 2 bytes
    // - rgKey[3]: 3 * u16 (key column indices, 0xFFFF for unused)
    // - rgOrder[3]: 3 * u16 (0=ascending, 1=descending)
    if data.len() < 24 {
        return None;
    }

    let range = parse_ref8(data)?;
    let grbit = u16::from_le_bytes([data[8], data[9]]);
    let c_keys = u16::from_le_bytes([data[10], data[11]]) as usize;
    // The classic BIFF8 SORT record supports up to 3 keys. If the value is larger, treat this as
    // a different/unrecognized layout and let other candidate parsers try.
    if c_keys > 3 {
        return None;
    }

    let key_cols = [
        u16::from_le_bytes([data[12], data[13]]),
        u16::from_le_bytes([data[14], data[15]]),
        u16::from_le_bytes([data[16], data[17]]),
    ];
    let orders = [
        u16::from_le_bytes([data[18], data[19]]),
        u16::from_le_bytes([data[20], data[21]]),
        u16::from_le_bytes([data[22], data[23]]),
    ];

    let mut keys = Vec::new();
    for i in 0..c_keys {
        let col_raw = key_cols[i];
        if col_raw == 0xFFFF {
            continue;
        }
        let descending = orders[i] != 0;
        keys.push(SortKey {
            col_raw,
            descending,
        });
    }

    Some(ParsedSortRecord { range, grbit, keys })
}

fn build_sort_state_from_parsed(
    parsed: ParsedSortRecord,
    auto_filter_range: Range,
) -> Option<SortState> {
    let has_header = sort_record_has_header(parsed.grbit, parsed.range, auto_filter_range);
    let start_row = if has_header {
        parsed.range.start.row.saturating_add(1)
    } else {
        parsed.range.start.row
    };
    if start_row > parsed.range.end.row {
        return None;
    }

    let mut conditions: Vec<SortCondition> = Vec::new();
    for key in parsed.keys {
        let Some(col) = resolve_key_col(key.col_raw, parsed.range) else {
            continue;
        };
        let key_range = Range::new(
            CellRef::new(start_row, col),
            CellRef::new(parsed.range.end.row, col),
        );
        // Only keep keys that still fall within the AutoFilter range after header adjustments.
        if !range_matches_or_contained_by(auto_filter_range, key_range) {
            continue;
        }
        conditions.push(SortCondition {
            range: key_range,
            descending: key.descending,
        });
    }

    (!conditions.is_empty()).then_some(SortState { conditions })
}

fn parse_sort_record_grbit_then_count(data: &[u8]) -> Option<ParsedSortRecord> {
    // Layout:
    // - Ref: rwFirst, rwLast, colFirst, colLast (8 bytes)
    // - grbit (2 bytes)
    // - cKey (2 bytes)
    // - repeated (col, grbitKey) pairs (4 bytes each)
    if data.len() < 12 {
        return None;
    }

    let range = parse_ref8(data)?;
    let grbit = u16::from_le_bytes([data[8], data[9]]);
    let c_key = u16::from_le_bytes([data[10], data[11]]) as usize;
    if c_key == 0 || c_key > 64 {
        return None;
    }

    let needed = 12usize.checked_add(c_key.checked_mul(4)?)?;
    if data.len() < needed {
        return None;
    }

    let mut keys = Vec::with_capacity(c_key);
    let mut off = 12usize;
    for _ in 0..c_key {
        let col_raw = u16::from_le_bytes([data[off], data[off + 1]]);
        let grbit_key = u16::from_le_bytes([data[off + 2], data[off + 3]]);
        off += 4;
        keys.push(SortKey {
            col_raw,
            descending: sort_key_is_descending(grbit_key),
        });
    }

    Some(ParsedSortRecord { range, grbit, keys })
}

fn parse_sort_record_count_then_grbit(data: &[u8]) -> Option<ParsedSortRecord> {
    // Some writers swap the `grbit` and `cKey` fields. Accept both.
    if data.len() < 12 {
        return None;
    }

    let range = parse_ref8(data)?;
    let c_key = u16::from_le_bytes([data[8], data[9]]) as usize;
    let grbit = u16::from_le_bytes([data[10], data[11]]);
    if c_key == 0 || c_key > 64 {
        return None;
    }

    let needed = 12usize.checked_add(c_key.checked_mul(4)?)?;
    if data.len() < needed {
        return None;
    }

    let mut keys = Vec::with_capacity(c_key);
    let mut off = 12usize;
    for _ in 0..c_key {
        let col_raw = u16::from_le_bytes([data[off], data[off + 1]]);
        let grbit_key = u16::from_le_bytes([data[off + 2], data[off + 3]]);
        off += 4;
        keys.push(SortKey {
            col_raw,
            descending: sort_key_is_descending(grbit_key),
        });
    }

    Some(ParsedSortRecord { range, grbit, keys })
}

fn parse_sort_record_fixed_3_keys(data: &[u8]) -> Option<ParsedSortRecord> {
    // Legacy fixed-size layout used by some BIFF writers:
    // - Ref: rwFirst, rwLast, colFirst, colLast (8 bytes)
    // - colKey1, colKey2, colKey3 (6 bytes)
    // - grbit (2 bytes)
    //
    // Descending flags are stored as bits in `grbit`:
    // - bit0 => key1 descending
    // - bit1 => key2 descending
    // - bit2 => key3 descending
    if data.len() < 16 {
        return None;
    }

    let range = parse_ref8(data)?;
    let key_cols = [
        u16::from_le_bytes([data[8], data[9]]),
        u16::from_le_bytes([data[10], data[11]]),
        u16::from_le_bytes([data[12], data[13]]),
    ];
    let grbit = u16::from_le_bytes([data[14], data[15]]);

    let mut keys = Vec::new();
    for (idx, col_raw) in key_cols.into_iter().enumerate() {
        if col_raw == 0xFFFF {
            continue;
        }
        let descending = (grbit & (1u16 << idx)) != 0;
        keys.push(SortKey {
            col_raw,
            descending,
        });
    }

    if keys.is_empty() {
        return None;
    }

    Some(ParsedSortRecord { range, grbit, keys })
}

fn parse_ref8(data: &[u8]) -> Option<Range> {
    if data.len() < 8 {
        return None;
    }
    let rw_first = u16::from_le_bytes([data[0], data[1]]) as u32;
    let rw_last = u16::from_le_bytes([data[2], data[3]]) as u32;
    let col_first = (u16::from_le_bytes([data[4], data[5]]) & BIFF8_COL_INDEX_MASK) as u32;
    let col_last = (u16::from_le_bytes([data[6], data[7]]) & BIFF8_COL_INDEX_MASK) as u32;

    if rw_first >= EXCEL_MAX_ROWS
        || rw_last >= EXCEL_MAX_ROWS
        || col_first >= EXCEL_MAX_COLS
        || col_last >= EXCEL_MAX_COLS
    {
        return None;
    }

    Some(Range::new(
        CellRef::new(rw_first, col_first),
        CellRef::new(rw_last, col_last),
    ))
}

fn range_matches_or_contained_by(outer: Range, inner: Range) -> bool {
    outer == inner || (outer.contains(inner.start) && outer.contains(inner.end))
}

fn resolve_key_col(col_raw: u16, sort_range: Range) -> Option<u32> {
    let col_masked = (col_raw & BIFF8_COL_INDEX_MASK) as u32;
    // Prefer interpreting `col_raw` as an absolute sheet column index.
    if col_masked >= sort_range.start.col && col_masked <= sort_range.end.col {
        return Some(col_masked);
    }

    // Some writers store the sort key as an offset within the sorted range.
    let abs = sort_range.start.col.saturating_add(col_masked);
    (abs >= sort_range.start.col && abs <= sort_range.end.col).then_some(abs)
}

fn sort_key_is_descending(grbit_key: u16) -> bool {
    // Best-effort: many writers use 0/1, others use a bitfield.
    if grbit_key == 0 {
        return false;
    }
    if grbit_key == 1 {
        return true;
    }
    (grbit_key & 0x0001) != 0
}

fn sort_record_has_header(grbit: u16, sort_range: Range, auto_filter_range: Range) -> bool {
    let header_bit = (grbit & (SORT_GRBIT_HEADER_LOW | SORT_GRBIT_HEADER_HIGH)) != 0;

    // When no explicit header bit is present, infer header presence for AutoFilter sorts when
    // the sorted range starts at the AutoFilter header row (the common Excel encoding).
    header_bit || sort_range.start.row == auto_filter_range.start.row
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn frt_record(record_id: u16, rt: u16, payload: &[u8]) -> Vec<u8> {
        // `FrtHeader` is an 8-byte structure:
        // - rt (u16)
        // - grbitFrt (u16)
        // - reserved (u32)
        let mut out = Vec::with_capacity(8 + payload.len());
        out.extend_from_slice(&rt.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out.extend_from_slice(&0u32.to_le_bytes()); // reserved
        out.extend_from_slice(payload);
        record(record_id, &out)
    }

    const RECORD_BOF: u16 = 0x0809;
    const RECORD_EOF: u16 = 0x000A;

    fn bof_payload() -> Vec<u8> {
        // BIFF8 BOF payload is 16 bytes; for our parser any bytes are fine.
        vec![0u8; 16]
    }

    fn canonical_sort_payload_with_header() -> Vec<u8> {
        // SORT range: A1:C5, 2 keys: B ascending, A descending, header present.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        payload.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        payload.extend_from_slice(&2u16.to_le_bytes()); // colLast
        payload.extend_from_slice(&(SORT_GRBIT_HEADER_LOW).to_le_bytes()); // grbit
        payload.extend_from_slice(&2u16.to_le_bytes()); // cKey
        // rgKey[3]
        payload.extend_from_slice(&1u16.to_le_bytes()); // col B
        payload.extend_from_slice(&0u16.to_le_bytes()); // col A
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        // rgOrder[3]
        payload.extend_from_slice(&0u16.to_le_bytes()); // B ascending
        payload.extend_from_slice(&1u16.to_le_bytes()); // A descending
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused
        payload
    }

    fn canonical_sort_payload_one_key_b_desc_with_header() -> Vec<u8> {
        // SORT range: A1:C5, 1 key: B descending, header present.
        let mut payload = ref8_payload_a1_c5();
        payload.extend_from_slice(&(SORT_GRBIT_HEADER_LOW).to_le_bytes()); // grbit
        payload.extend_from_slice(&1u16.to_le_bytes()); // cKey
        // rgKey[3]
        payload.extend_from_slice(&1u16.to_le_bytes()); // col B
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        // rgOrder[3]
        payload.extend_from_slice(&1u16.to_le_bytes()); // B descending
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused
        payload
    }

    fn ref8_payload_a1_c5() -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        payload.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        payload.extend_from_slice(&2u16.to_le_bytes()); // colLast
        payload
    }

    #[test]
    fn parses_sort_record_with_header_excludes_header_row() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // SORT range: A1:C5, 2 keys: B ascending, A descending, header present.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        payload.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        payload.extend_from_slice(&2u16.to_le_bytes()); // colLast
        payload.extend_from_slice(&(SORT_GRBIT_HEADER_LOW).to_le_bytes()); // grbit
        payload.extend_from_slice(&2u16.to_le_bytes()); // cKey
        // rgKey[3]
        payload.extend_from_slice(&1u16.to_le_bytes()); // col B
        payload.extend_from_slice(&0u16.to_le_bytes()); // col A
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        // rgOrder[3]
        payload.extend_from_slice(&0u16.to_le_bytes()); // B ascending
        payload.extend_from_slice(&1u16.to_le_bytes()); // A descending
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            record(RECORD_SORT, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![
                SortCondition {
                    range: Range::from_a1("B2:B5").unwrap(),
                    descending: false,
                },
                SortCondition {
                    range: Range::from_a1("A2:A5").unwrap(),
                    descending: true,
                },
            ]
        );
    }

    #[test]
    fn parses_sort_record_without_header_keeps_first_row() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // SORT range: A2:C5 (data-only), 1 key: C descending, header flag not set.
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes()); // rwFirst (row 2)
        payload.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        payload.extend_from_slice(&2u16.to_le_bytes()); // colLast
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&1u16.to_le_bytes()); // cKey
        // rgKey[3]
        payload.extend_from_slice(&2u16.to_le_bytes()); // col C
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // unused
        // rgOrder[3]
        payload.extend_from_slice(&1u16.to_le_bytes()); // C descending
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused
        payload.extend_from_slice(&0u16.to_le_bytes()); // unused

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            record(RECORD_SORT, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("C2:C5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sort12_frt_record_with_embedded_canonical_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORT12, RT_SORT12, &canonical_sort_payload_with_header()),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![
                SortCondition {
                    range: Range::from_a1("B2:B5").unwrap(),
                    descending: false,
                },
                SortCondition {
                    range: Range::from_a1("A2:A5").unwrap(),
                    descending: true,
                },
            ]
        );
    }

    #[test]
    fn warns_once_on_unsupported_sort12_with_relevant_ref8() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // Payload contains a relevant Ref8, but is too short to decode as a classic SORT record.
        let payload = ref8_payload_a1_c5();
        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORT12, RT_SORT12, &payload),
            frt_record(RT_SORT12, RT_SORT12, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.sort_state.is_none());
        assert_eq!(parsed.warnings, vec!["unsupported Sort12".to_string()]);
    }

    #[test]
    fn parses_sort12_frt_record_with_embedded_one_key_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(
                RT_SORT12,
                RT_SORT12,
                &canonical_sort_payload_one_key_b_desc_with_header(),
            ),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sort12_alt_frt_record_with_embedded_one_key_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(
                RT_SORT12_ALT,
                RT_SORT12_ALT,
                &canonical_sort_payload_one_key_b_desc_with_header(),
            ),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sort12_frt_record_continued_via_continuefrt12() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // Split the embedded SORT payload across Sort12 + ContinueFrt12 records.
        let sort_payload = canonical_sort_payload_one_key_b_desc_with_header();
        let (first, rest) = sort_payload.split_at(8); // split after Ref8

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORT12, RT_SORT12, first),
            frt_record(RECORD_CONTINUEFRT12, RECORD_CONTINUEFRT12, rest),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sortdata12_frt_record_with_embedded_canonical_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORTDATA12, RT_SORTDATA12, &canonical_sort_payload_with_header()),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![
                SortCondition {
                    range: Range::from_a1("B2:B5").unwrap(),
                    descending: false,
                },
                SortCondition {
                    range: Range::from_a1("A2:A5").unwrap(),
                    descending: true,
                },
            ]
        );
    }

    #[test]
    fn warns_once_on_unsupported_sortdata12_with_relevant_ref8() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // Payload contains a relevant Ref8, but is too short to decode as a classic SORT record.
        let payload = ref8_payload_a1_c5();
        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORTDATA12, RT_SORTDATA12, &payload),
            frt_record(RT_SORTDATA12, RT_SORTDATA12, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.sort_state.is_none());
        assert_eq!(parsed.warnings, vec!["unsupported SortData12".to_string()]);
    }

    #[test]
    fn parses_sortdata12_frt_record_with_embedded_one_key_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(
                RT_SORTDATA12,
                RT_SORTDATA12,
                &canonical_sort_payload_one_key_b_desc_with_header(),
            ),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sortdata12_alt_frt_record_with_embedded_one_key_sort() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(
                RT_SORTDATA12_ALT,
                RT_SORTDATA12_ALT,
                &canonical_sort_payload_one_key_b_desc_with_header(),
            ),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn parses_sortdata12_frt_record_continued_via_continuefrt12() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // Split the embedded SORT payload across SortData12 + ContinueFrt12 records.
        let sort_payload = canonical_sort_payload_one_key_b_desc_with_header();
        let (first, rest) = sort_payload.split_at(8); // split after Ref8

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORTDATA12, RT_SORTDATA12, first),
            frt_record(RECORD_CONTINUEFRT12, RECORD_CONTINUEFRT12, rest),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.warnings.is_empty(), "unexpected warnings: {:?}", parsed.warnings);

        let sort_state = parsed.sort_state.expect("expected sort_state");
        assert_eq!(
            sort_state.conditions,
            vec![SortCondition {
                range: Range::from_a1("B2:B5").unwrap(),
                descending: true,
            }]
        );
    }

    #[test]
    fn does_not_warn_on_unsupported_sort12_with_irrelevant_ref8() {
        // AutoFilter range: A1:C5.
        let af = Range::from_a1("A1:C5").unwrap();

        // Ref8 range: D1:F5 (outside A1:C5).
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        payload.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        payload.extend_from_slice(&3u16.to_le_bytes()); // colFirst (D)
        payload.extend_from_slice(&5u16.to_le_bytes()); // colLast (F)

        let stream = [
            record(RECORD_BOF, &bof_payload()),
            frt_record(RT_SORT12, RT_SORT12, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_sort_state(&stream, 0, af).unwrap();
        assert!(parsed.sort_state.is_none());
        assert!(parsed.warnings.is_empty());
    }
}
