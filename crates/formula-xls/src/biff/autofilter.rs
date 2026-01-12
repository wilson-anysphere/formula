use std::collections::HashMap;

use formula_model::{CellRef, Range, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, XLNM_FILTER_DATABASE};

use super::{records, strings, BiffVersion};

// Workbook-global record ids.
// See [MS-XLS]:
// - EXTERNSHEET: 2.4.105 (0x0017)
// - NAME: 2.4.150 (0x0018)
// - SUPBOOK: 2.4.271 (0x01AE)
const RECORD_EXTERNSHEET: u16 = 0x0017;
const RECORD_NAME: u16 = 0x0018;
const RECORD_SUPBOOK: u16 = 0x01AE;

// NAME record option flags [MS-XLS 2.4.150].
// We only care about `fBuiltin` (built-in name).
const NAME_FLAG_BUILTIN: u16 = 0x0020;

// Built-in defined name ids [MS-XLS 2.5.66].
// These are stored in the NAME record when `fBuiltin == 1`.
const BUILTIN_NAME_FILTER_DATABASE: u8 = 0x0D;

// BIFF8 string flags (mirrors `strings.rs`).
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

// BIFF8 formula tokens for 2D/3D references.
const PTG_REF: [u8; 3] = [0x24, 0x44, 0x64];
const PTG_AREA: [u8; 3] = [0x25, 0x45, 0x65];
const PTG_REF3D: [u8; 3] = [0x3A, 0x5A, 0x7A];
const PTG_AREA3D: [u8; 3] = [0x3B, 0x5B, 0x7B];

/// BIFF8 `XTI` entry from the `EXTERNSHEET` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Xti {
    pub(crate) i_sup_book: u16,
    pub(crate) itab_first: u16,
    pub(crate) itab_last: u16,
}

#[derive(Debug, Default)]
pub(crate) struct ParsedFilterDatabaseRanges {
    /// Mapping from 0-based BIFF sheet index to the AutoFilter range.
    pub(crate) by_sheet: HashMap<usize, Range>,
    /// Non-fatal parse warnings.
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone)]
struct FilterDatabaseName {
    record_offset: usize,
    /// NAME.itab (1-based sheet index, or 0 for workbook-scope).
    itab: u16,
    rgce: Vec<u8>,
}

/// Parse workbook-global `EXTERNSHEET` and `NAME` records to recover AutoFilter ranges stored in
/// the built-in `_FilterDatabase` name.
///
/// This is intentionally best-effort: malformed records are skipped and surfaced as warnings.
pub(crate) fn parse_biff_filter_database_ranges(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
    // Optional count of worksheets in the workbook, used to validate resolved sheet indices.
    // When `None`, sheet index bounds checks are skipped.
    sheet_count: Option<usize>,
) -> Result<ParsedFilterDatabaseRanges, String> {
    let mut out = ParsedFilterDatabaseRanges::default();

    if biff != BiffVersion::Biff8 {
        // BIFF5 uses a different formula token layout for 3D references; we currently only
        // implement the BIFF8 mapping used by `.xls` files written by Excel 97+.
        return Ok(out);
    }

    let allows_continuation = |record_id: u16| {
        record_id == RECORD_NAME || record_id == RECORD_SUPBOOK || record_id == RECORD_EXTERNSHEET
    };
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    let mut saw_eof = false;

    let mut supbook_count: u16 = 0;
    let mut internal_supbook_index: Option<u16> = None;

    // There is typically only one EXTERNSHEET record; keep the last one we see.
    let mut externsheets: Vec<Xti> = Vec::new();

    let mut filter_database_names: Vec<FilterDatabaseName> = Vec::new();

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // Stop at the beginning of the next substream (worksheet BOF).
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_SUPBOOK => {
                if internal_supbook_index.is_none()
                    && supbook_record_is_internal(record.data.as_ref(), codepage)
                {
                    internal_supbook_index = Some(supbook_count);
                }
                supbook_count = supbook_count.saturating_add(1);
            }
            RECORD_EXTERNSHEET => {
                externsheets = parse_externsheet_record_best_effort(record.data.as_ref());
            }
            RECORD_NAME => {
                match parse_name_record_best_effort(record.data.as_ref(), biff, codepage) {
                    Ok(Some(parsed)) => {
                        if parsed.name == XLNM_FILTER_DATABASE {
                            filter_database_names.push(FilterDatabaseName {
                                record_offset: record.offset,
                                itab: parsed.itab,
                                rgce: parsed.rgce,
                            });
                        }
                    }
                    Ok(None) => {}
                    Err(err) => out.warnings.push(format!(
                        "failed to decode NAME record at offset {}: {err}",
                        record.offset
                    )),
                }
            }
            records::RECORD_EOF => {
                saw_eof = true;
                break;
            }
            _ => {}
        }
    }

    if !saw_eof {
        // Consistent with other BIFF helpers: tolerate missing EOF but surface a warning so
        // callers understand the parse was partial.
        out.warnings
            .push("unexpected end of workbook globals stream (missing EOF)".to_string());
    }

    for name in filter_database_names {
        let base_sheet = if name.itab != 0 {
            // NAME.itab is 1-based; 0 indicates workbook-scope.
            Some(name.itab.saturating_sub(1) as usize)
        } else {
            None
        };

        match decode_filter_database_rgce(
            &name.rgce,
            base_sheet,
            &externsheets,
            internal_supbook_index,
            sheet_count,
        ) {
            Ok(Some((sheet_idx, range))) => {
                out.by_sheet.insert(sheet_idx, range);
            }
            Ok(None) => out.warnings.push(format!(
                "skipping `_FilterDatabase` NAME record at offset {}: unsupported formula",
                name.record_offset
            )),
            Err(err) => out.warnings.push(format!(
                "skipping `_FilterDatabase` NAME record at offset {}: {err}",
                name.record_offset
            )),
        }
    }

    Ok(out)
}

fn parse_externsheet_record_best_effort(data: &[u8]) -> Vec<Xti> {
    if data.len() < 2 {
        return Vec::new();
    }
    let cxti = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut out = Vec::with_capacity(cxti);

    let mut pos = 2usize;
    let available = data.len().saturating_sub(pos);
    let max_entries = available / 6;
    let take = cxti.min(max_entries);

    for _ in 0..take {
        let i_sup_book = u16::from_le_bytes([data[pos], data[pos + 1]]);
        let itab_first = u16::from_le_bytes([data[pos + 2], data[pos + 3]]);
        let itab_last = u16::from_le_bytes([data[pos + 4], data[pos + 5]]);
        out.push(Xti {
            i_sup_book,
            itab_first,
            itab_last,
        });
        pos += 6;
    }

    out
}

fn supbook_record_is_internal(data: &[u8], codepage: u16) -> bool {
    // [MS-XLS 2.4.271] describes a special "internal references" SUPBOOK record.
    //
    // Different producers appear to encode the "internal workbook" marker differently. Prefer
    // best-effort detection over strict validation so we can recover 3D NAME references in the
    // wild.
    // Best-effort: detect the internal tag in `virtPath` (an XLUnicodeString after `ctab`).
    if data.len() >= 5 {
        if let Ok((s, _)) = strings::parse_biff8_unicode_string(&data[2..], codepage) {
            if matches!(s.as_str(), "\u{1}" | "\0" | "\u{1}\u{4}") {
                return true;
            }
        }
    }

    false
}

#[derive(Debug, Clone)]
struct ParsedNameRecord {
    name: String,
    itab: u16,
    rgce: Vec<u8>,
}

fn parse_name_record_best_effort(
    data: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<Option<ParsedNameRecord>, String> {
    // NAME [MS-XLS 2.4.150]
    if data.len() < 14 {
        return Err("NAME record too short".to_string());
    }

    let grbit = u16::from_le_bytes([data[0], data[1]]);
    let cch = data[3] as usize;
    let cce = u16::from_le_bytes([data[4], data[5]]) as usize;
    let itab = u16::from_le_bytes([data[8], data[9]]);

    let cch_cust_menu = data[10] as usize;
    let cch_description = data[11] as usize;
    let cch_help_topic = data[12] as usize;
    let cch_status_text = data[13] as usize;

    let mut pos = 14usize;

    let is_builtin = (grbit & NAME_FLAG_BUILTIN) != 0;

    let name = if is_builtin {
        // Built-in name: `cch` is usually 1 and the name field is a 1-byte built-in id.
        let code = *data
            .get(pos)
            .ok_or_else(|| "truncated built-in NAME".to_string())?;
        pos += 1;
        match code {
            BUILTIN_NAME_FILTER_DATABASE => XLNM_FILTER_DATABASE.to_string(),
            _ => {
                // Ignore unknown built-in names; callers only care about FilterDatabase today.
                return Ok(None);
            }
        }
    } else {
        match biff {
            BiffVersion::Biff5 => {
                let bytes = data
                    .get(pos..pos + cch)
                    .ok_or_else(|| "truncated NAME string".to_string())?;
                pos += cch;
                strings::decode_ansi(codepage, bytes)
            }
            BiffVersion::Biff8 => {
                let (s, consumed) = parse_biff8_unicode_string_no_cch(&data[pos..], cch, codepage)?;
                pos = pos
                    .checked_add(consumed)
                    .ok_or_else(|| "NAME offset overflow".to_string())?;
                s
            }
        }
    };

    let rest = data.get(pos..).unwrap_or_default();

    // The on-disk NAME record layout has evolved, and different producers appear to place the
    // optional strings either before or after the formula token stream. For our purposes (built-in
    // FilterDatabase), those strings are typically empty; still, be defensive and accept both
    // layouts.
    let try_layout = |strings_first: bool| -> Result<Option<Vec<u8>>, String> {
        let mut cursor = rest;

        if strings_first {
            cursor = skip_name_optional_strings(
                cursor,
                biff,
                codepage,
                cch_cust_menu,
                cch_description,
                cch_help_topic,
                cch_status_text,
            )?;
        }

        if cursor.len() < cce {
            return Ok(None);
        }
        let rgce = cursor[..cce].to_vec();
        cursor = &cursor[cce..];

        if !strings_first {
            let _ = skip_name_optional_strings(
                cursor,
                biff,
                codepage,
                cch_cust_menu,
                cch_description,
                cch_help_topic,
                cch_status_text,
            )?;
        }

        Ok(Some(rgce))
    };

    let rgce = try_layout(false)?.or(try_layout(true)?).ok_or_else(|| {
        "truncated NAME rgce (formula token stream extends past end of record)".to_string()
    })?;

    Ok(Some(ParsedNameRecord { name, itab, rgce }))
}

fn skip_name_optional_strings<'a>(
    mut input: &'a [u8],
    biff: BiffVersion,
    codepage: u16,
    cch_cust_menu: usize,
    cch_description: usize,
    cch_help_topic: usize,
    cch_status_text: usize,
) -> Result<&'a [u8], String> {
    for &cch in &[
        cch_cust_menu,
        cch_description,
        cch_help_topic,
        cch_status_text,
    ] {
        if cch == 0 {
            continue;
        }
        match biff {
            BiffVersion::Biff5 => {
                input = input
                    .get(cch..)
                    .ok_or_else(|| "truncated NAME optional string".to_string())?;
            }
            BiffVersion::Biff8 => {
                let (_, consumed) = parse_biff8_unicode_string_no_cch(input, cch, codepage)?;
                input = input
                    .get(consumed..)
                    .ok_or_else(|| "truncated NAME optional string".to_string())?;
            }
        }
    }
    Ok(input)
}

fn parse_biff8_unicode_string_no_cch(
    input: &[u8],
    cch: usize,
    codepage: u16,
) -> Result<(String, usize), String> {
    if input.is_empty() {
        return Err("unexpected end of string".to_string());
    }
    let flags = input[0];
    let mut offset = 1usize;

    let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
        if input.len() < offset + 2 {
            return Err("unexpected end of string".to_string());
        }
        let runs = u16::from_le_bytes([input[offset], input[offset + 1]]) as usize;
        offset += 2;
        runs
    } else {
        0
    };

    let ext_size = if flags & STR_FLAG_EXT != 0 {
        if input.len() < offset + 4 {
            return Err("unexpected end of string".to_string());
        }
        let size = u32::from_le_bytes([
            input[offset],
            input[offset + 1],
            input[offset + 2],
            input[offset + 3],
        ]) as usize;
        offset += 4;
        size
    } else {
        0
    };

    let is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
    let char_bytes = if is_unicode {
        cch.checked_mul(2)
            .ok_or_else(|| "string length overflow".to_string())?
    } else {
        cch
    };

    let chars = input
        .get(offset..offset + char_bytes)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    offset += char_bytes;

    let value = if is_unicode {
        let mut u16s = Vec::with_capacity(cch);
        for chunk in chars.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        strings::decode_ansi(codepage, chars)
    };

    let richtext_bytes = richtext_runs
        .checked_mul(4)
        .ok_or_else(|| "rich text run count overflow".to_string())?;

    if input.len() < offset + richtext_bytes + ext_size {
        return Err("unexpected end of string".to_string());
    }
    offset += richtext_bytes + ext_size;

    Ok((value, offset))
}

fn decode_filter_database_rgce(
    rgce: &[u8],
    base_sheet: Option<usize>,
    externsheets: &[Xti],
    internal_supbook_index: Option<u16>,
    sheet_count: Option<usize>,
) -> Result<Option<(usize, Range)>, String> {
    if rgce.is_empty() {
        return Ok(None);
    }

    let ptg = rgce[0];
    let mut pos = 1usize;

    let (sheet_override, range, consumed) = if PTG_AREA.contains(&ptg) {
        let (range, used) = decode_ptg_area(rgce.get(pos..).unwrap_or_default())?;
        (None, range, used)
    } else if PTG_REF.contains(&ptg) {
        let (range, used) = decode_ptg_ref(rgce.get(pos..).unwrap_or_default())?;
        (None, range, used)
    } else if PTG_AREA3D.contains(&ptg) {
        let (sheet, range, used) = decode_ptg_area3d(
            rgce.get(pos..).unwrap_or_default(),
            externsheets,
            internal_supbook_index,
            sheet_count,
        )?;
        (sheet, range, used)
    } else if PTG_REF3D.contains(&ptg) {
        let (sheet, range, used) = decode_ptg_ref3d(
            rgce.get(pos..).unwrap_or_default(),
            externsheets,
            internal_supbook_index,
            sheet_count,
        )?;
        (sheet, range, used)
    } else {
        return Ok(None);
    };

    pos = pos
        .checked_add(consumed)
        .ok_or_else(|| "rgce offset overflow".to_string())?;

    if pos != rgce.len() {
        // FilterDatabase names in valid workbooks are usually a single reference token. If extra
        // tokens are present, treat as unsupported so we don't misinterpret the formula.
        return Ok(None);
    }

    let sheet_idx = sheet_override.or(base_sheet).ok_or_else(|| {
        "workbook-scope `_FilterDatabase` NAME formula does not specify a sheet".to_string()
    })?;

    // Validate range bounds against Excel limits.
    if range.end.row >= EXCEL_MAX_ROWS || range.end.col >= EXCEL_MAX_COLS {
        return Err(format!(
            "AutoFilter range `{range}` exceeds Excel bounds (max row {}, max col {})",
            EXCEL_MAX_ROWS, EXCEL_MAX_COLS
        ));
    }

    if let Some(count) = sheet_count {
        if sheet_idx >= count {
            return Err(format!(
                "AutoFilter sheet index {sheet_idx} out of range (sheet_count={count})"
            ));
        }
    }

    Ok(Some((sheet_idx, range)))
}

fn decode_ptg_ref(payload: &[u8]) -> Result<(Range, usize), String> {
    if payload.len() < 4 {
        return Err("truncated PtgRef token".to_string());
    }
    let row = u16::from_le_bytes([payload[0], payload[1]]) as u32;
    let col_field = u16::from_le_bytes([payload[2], payload[3]]);
    let col = (col_field & 0x3FFF) as u32;
    Ok((
        Range::new(CellRef::new(row, col), CellRef::new(row, col)),
        4,
    ))
}

fn decode_ptg_area(payload: &[u8]) -> Result<(Range, usize), String> {
    if payload.len() < 8 {
        return Err("truncated PtgArea token".to_string());
    }
    let row_first = u16::from_le_bytes([payload[0], payload[1]]) as u32;
    let row_last = u16::from_le_bytes([payload[2], payload[3]]) as u32;
    let col_first_field = u16::from_le_bytes([payload[4], payload[5]]);
    let col_last_field = u16::from_le_bytes([payload[6], payload[7]]);
    let col_first = (col_first_field & 0x3FFF) as u32;
    let col_last = (col_last_field & 0x3FFF) as u32;
    Ok((
        Range::new(
            CellRef::new(row_first, col_first),
            CellRef::new(row_last, col_last),
        ),
        8,
    ))
}

fn decode_ptg_ref3d(
    payload: &[u8],
    externsheets: &[Xti],
    internal_supbook_index: Option<u16>,
    sheet_count: Option<usize>,
) -> Result<(Option<usize>, Range, usize), String> {
    if payload.len() < 6 {
        return Err("truncated PtgRef3d token".to_string());
    }
    let ixti = u16::from_le_bytes([payload[0], payload[1]]);
    let row = u16::from_le_bytes([payload[2], payload[3]]) as u32;
    let col_field = u16::from_le_bytes([payload[4], payload[5]]);
    let col = (col_field & 0x3FFF) as u32;

    let sheet =
        resolve_ixti_to_internal_sheet(ixti, externsheets, internal_supbook_index, sheet_count);
    let range = Range::new(CellRef::new(row, col), CellRef::new(row, col));
    Ok((sheet, range, 6))
}

fn decode_ptg_area3d(
    payload: &[u8],
    externsheets: &[Xti],
    internal_supbook_index: Option<u16>,
    sheet_count: Option<usize>,
) -> Result<(Option<usize>, Range, usize), String> {
    if payload.len() < 10 {
        return Err("truncated PtgArea3d token".to_string());
    }
    let ixti = u16::from_le_bytes([payload[0], payload[1]]);
    let row_first = u16::from_le_bytes([payload[2], payload[3]]) as u32;
    let row_last = u16::from_le_bytes([payload[4], payload[5]]) as u32;
    let col_first_field = u16::from_le_bytes([payload[6], payload[7]]);
    let col_last_field = u16::from_le_bytes([payload[8], payload[9]]);
    let col_first = (col_first_field & 0x3FFF) as u32;
    let col_last = (col_last_field & 0x3FFF) as u32;

    let sheet =
        resolve_ixti_to_internal_sheet(ixti, externsheets, internal_supbook_index, sheet_count);
    let range = Range::new(
        CellRef::new(row_first, col_first),
        CellRef::new(row_last, col_last),
    );
    Ok((sheet, range, 10))
}

fn resolve_ixti_to_internal_sheet(
    ixti: u16,
    externsheets: &[Xti],
    internal_supbook_index: Option<u16>,
    sheet_count: Option<usize>,
) -> Option<usize> {
    let xti = *externsheets.get(ixti as usize)?;

    // BIFF8 conventions in the wild:
    // - Some writers use `iSupBook==0` to mean "internal workbook", regardless of the SUPBOOK
    //   table contents.
    // - Other writers reference the internal workbook SUPBOOK explicitly via its index.
    //
    // Be permissive and treat either as internal when possible.
    let is_internal = if xti.i_sup_book == 0 {
        true
    } else if let Some(internal_idx) = internal_supbook_index {
        xti.i_sup_book == internal_idx
    } else {
        false
    };
    if !is_internal {
        return None;
    }
    if xti.itab_first != xti.itab_last {
        return None;
    }

    let sheet_idx = xti.itab_first as usize;
    if let Some(count) = sheet_count {
        if sheet_idx >= count {
            return None;
        }
    }
    Some(sheet_idx)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    fn bof_globals() -> Vec<u8> {
        // Minimal BIFF8 BOF payload (16 bytes). Only version + dt matter for our parsers.
        let mut out = [0u8; 16];
        out[0..2].copy_from_slice(&0x0600u16.to_le_bytes()); // BIFF8
        out[2..4].copy_from_slice(&0x0005u16.to_le_bytes()); // workbook globals
        out.to_vec()
    }

    #[test]
    fn decodes_workbook_scope_filter_database_via_externsheet_ptgarea3d() {
        // Two SUPBOOK records so the internal workbook reference is not at index 0.
        // This ensures we resolve iSupBook via SUPBOOK rather than relying on the "iSupBook==0"
        // heuristic.
        let supbook_external = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&3u16.to_le_bytes()); // cch
            data.push(0); // flags (compressed)
            data.extend_from_slice(b"ext"); // virtPath
            record(RECORD_SUPBOOK, &data)
        };
        let supbook_internal = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&1u16.to_le_bytes()); // cch
            data.push(0); // flags (compressed)
            data.push(0x01); // virtPath marker
            record(RECORD_SUPBOOK, &data)
        };

        // EXTERNSHEET with one XTI entry mapping ixti=0 -> internal sheet 0 (iSupBook=1).
        //
        // EXTERNSHEET records can be split across CONTINUE records; split the payload mid-u16 to
        // ensure the parser coalesces continuations.
        let externsheet_full = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // cXTI
            data.extend_from_slice(&1u16.to_le_bytes()); // iSupBook (internal SUPBOOK index)
            data.extend_from_slice(&0u16.to_le_bytes()); // itabFirst
            data.extend_from_slice(&0u16.to_le_bytes()); // itabLast
            data
        };
        let externsheet_first = record(RECORD_EXTERNSHEET, &externsheet_full[..3]);
        let externsheet_continue = record(records::RECORD_CONTINUE, &externsheet_full[3..]);

        // NAME record: built-in _FilterDatabase, workbook-scope (itab=0), rgce = PtgArea3d.
        let name = {
            // rgce = [ptgArea3d][ixti=0][rwFirst=0][rwLast=4][colFirst=0][colLast=2]
            let mut rgce = Vec::new();
            rgce.push(0x3B); // PtgArea3d
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
            rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
            rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
            rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
            rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

            let mut data = Vec::new();
            data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
            data.push(0); // chKey
            data.push(1); // cch (builtin id length)
            data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
            data.extend_from_slice(&0u16.to_le_bytes()); // ixals
            data.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook-scope)
            data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
            data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
            data.extend_from_slice(&rgce);
            record(RECORD_NAME, &data)
        };

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            supbook_external,
            supbook_internal,
            externsheet_first,
            externsheet_continue,
            name,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_filter_database_ranges(&stream, BiffVersion::Biff8, 1252, Some(1))
            .expect("parse");

        assert_eq!(
            parsed.by_sheet.get(&0).copied(),
            Some(Range::new(CellRef::new(0, 0), CellRef::new(4, 2)))
        );
    }

    #[test]
    fn treats_isupbook_zero_as_internal_even_when_internal_supbook_index_is_nonzero() {
        // SUPBOOK[0]: add-in marker (not internal workbook).
        let supbook_addin = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&1u16.to_le_bytes()); // cch
            data.push(0); // flags (compressed)
            data.push(0x02); // virtPath marker for add-in
            record(RECORD_SUPBOOK, &data)
        };
        // SUPBOOK[1]: internal workbook marker.
        let supbook_internal = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&1u16.to_le_bytes()); // cch
            data.push(0); // flags (compressed)
            data.push(0x01); // virtPath marker for internal workbook
            record(RECORD_SUPBOOK, &data)
        };

        // EXTERNSHEET uses iSupBook==0 for internal refs, even though the internal workbook SUPBOOK
        // record is at index 1. Some writers do this; we should still resolve it.
        let externsheet = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // cXTI
            data.extend_from_slice(&0u16.to_le_bytes()); // iSupBook == 0 (internal)
            data.extend_from_slice(&0u16.to_le_bytes()); // itabFirst
            data.extend_from_slice(&0u16.to_le_bytes()); // itabLast
            record(RECORD_EXTERNSHEET, &data)
        };

        // NAME record: built-in _FilterDatabase, workbook-scope, rgce = PtgArea3d.
        let name = {
            let mut rgce = Vec::new();
            rgce.push(0x3B); // PtgArea3d
            rgce.extend_from_slice(&0u16.to_le_bytes()); // ixti
            rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
            rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
            rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
            rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

            let mut data = Vec::new();
            data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
            data.push(0); // chKey
            data.push(1); // cch (builtin id length)
            data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
            data.extend_from_slice(&0u16.to_le_bytes()); // ixals
            data.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook-scope)
            data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
            data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
            data.extend_from_slice(&rgce);
            record(RECORD_NAME, &data)
        };

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            supbook_addin,
            supbook_internal,
            externsheet,
            name,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_filter_database_ranges(&stream, BiffVersion::Biff8, 1252, Some(1))
            .expect("parse");

        assert_eq!(
            parsed.by_sheet.get(&0).copied(),
            Some(Range::new(CellRef::new(0, 0), CellRef::new(4, 2)))
        );
    }
}
