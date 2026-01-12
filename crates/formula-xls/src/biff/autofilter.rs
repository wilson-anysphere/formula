use std::collections::HashMap;

use formula_model::{CellRef, Range, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, XLNM_FILTER_DATABASE};

use super::{externsheet, records, strings, BiffVersion};

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
// Some writers appear to store the FilterDatabase name as a normal string rather than as a
// built-in id. Excel's visible built-in name omits the `_xlnm.` prefix used in XLSX.
const FILTER_DATABASE_NAME_ALIAS: &str = "_FilterDatabase";
// Some decoders (notably calamine) have been observed to surface `_FilterDatabase` missing the
// trailing `e`. This is likely due to quirks in some BIFF NAME encodings; accept it as a
// best-effort alias so we can still recover the filter range.
const FILTER_DATABASE_NAME_ALIAS_TRUNCATED: &str = "_FilterDatabas";

// BIFF8 string flags (mirrors `strings.rs`).
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

// BIFF8 formula tokens for 2D/3D references.
const PTG_REF: [u8; 3] = [0x24, 0x44, 0x64];
const PTG_AREA: [u8; 3] = [0x25, 0x45, 0x65];
const PTG_REF3D: [u8; 3] = [0x3A, 0x5A, 0x7A];
const PTG_AREA3D: [u8; 3] = [0x3B, 0x5B, 0x7B];

// BIFF8 worksheets (`.xls`) are limited to 256 columns (A..IV). Column indices are stored in a
// 2-byte field that also contains relative/absolute flags; in practice only the low 8 bits are
// meaningful for `.xls` column indices. Some producers also use `0x3FFF` as a "max column"
// sentinel; masking to 8 bits maps that to `0x00FF` (IV), matching Excel's limits.
const BIFF8_COL_INDEX_MASK: u16 = 0x00FF;
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
    let mut externsheets: Vec<externsheet::ExternSheetEntry> = Vec::new();

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
                let parsed = externsheet::parse_biff8_externsheet_record_data(
                    record.data.as_ref(),
                    record.offset,
                );
                externsheets = parsed.entries;
                out.warnings.extend(parsed.warnings);
            }
            RECORD_NAME => {
                match parse_name_record_best_effort(record.data.as_ref(), biff, codepage) {
                    Ok(Some(parsed)) => {
                        if is_filter_database_name(&parsed.name) {
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

fn is_filter_database_name(name: &str) -> bool {
    name.eq_ignore_ascii_case(XLNM_FILTER_DATABASE)
        || name.eq_ignore_ascii_case(FILTER_DATABASE_NAME_ALIAS)
        || name.eq_ignore_ascii_case(FILTER_DATABASE_NAME_ALIAS_TRUNCATED)
}

fn supbook_record_is_internal(data: &[u8], codepage: u16) -> bool {
    // [MS-XLS 2.4.271] describes a special "internal references" SUPBOOK record.
    //
    // Different producers appear to encode the "internal workbook" marker differently. Prefer
    // best-effort detection over strict validation so we can recover 3D NAME references in the
    // wild.
    //
    // Some producers appear to emit a minimal 4-byte SUPBOOK payload where the second u16 is a
    // sentinel marker (commonly `0x0401`) rather than an XLUnicodeString `virtPath`.
    if data.len() == 4 {
        let marker = u16::from_le_bytes([data[2], data[3]]);
        if marker == 0x0401 {
            return true;
        }
    }

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

    let mut name = if is_builtin {
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

    // Best-effort: BIFF Unicode strings can contain embedded NUL bytes in the wild; strip them so
    // name matching behaves like Excel UI semantics.
    if name.contains('\0') {
        name.retain(|ch| ch != '\0');
    }

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
    externsheets: &[externsheet::ExternSheetEntry],
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

    // FilterDatabase names in valid workbooks are usually a single reference token, but some
    // producers may include additional "wrapper" tokens (e.g. PtgParen or PtgAttr with only
    // formatting flags). Be permissive and ignore those where possible.
    //
    // If the trailing token stream contains anything else, treat the formula as unsupported so we
    // don't misinterpret it.
    while pos < rgce.len() {
        match rgce[pos] {
            // PtgParen (explicit parentheses): no payload.
            0x15 => pos = pos.saturating_add(1),
            // PtgAttr: [grbit: u8][wAttr: u16]
            0x19 => {
                pos = pos.saturating_add(1);
                if rgce.len() < pos + 3 {
                    return Err("truncated PtgAttr token".to_string());
                }
                let grbit = rgce[pos];
                let _w_attr = u16::from_le_bytes([rgce[pos + 1], rgce[pos + 2]]);
                pos = pos.saturating_add(3);

                // Some PtgAttr bits affect evaluation (notably tAttrSum / tAttrChoose). If present,
                // the formula is no longer a simple range reference; treat as unsupported.
                const T_ATTR_CHOOSE: u8 = 0x04;
                const T_ATTR_SUM: u8 = 0x10;
                if (grbit & (T_ATTR_CHOOSE | T_ATTR_SUM)) != 0 {
                    return Ok(None);
                }
            }
            _ => return Ok(None),
        }
    }

    let sheet_idx = match sheet_override.or(base_sheet) {
        Some(idx) => idx,
        None => {
            // Some files in the wild appear to store workbook-scoped `_FilterDatabase` names even
            // though the formula token stream does not include an explicit 3D sheet reference.
            // When there is only a single sheet, treat that as the implied target.
            if sheet_count == Some(1) {
                0
            } else {
                return Err(
                    "workbook-scope `_FilterDatabase` NAME formula does not specify a sheet"
                        .to_string(),
                );
            }
        }
    };

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
    let col = (col_field & BIFF8_COL_INDEX_MASK) as u32;
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
    let col_first = (col_first_field & BIFF8_COL_INDEX_MASK) as u32;
    let col_last = (col_last_field & BIFF8_COL_INDEX_MASK) as u32;
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
    externsheets: &[externsheet::ExternSheetEntry],
    internal_supbook_index: Option<u16>,
    sheet_count: Option<usize>,
) -> Result<(Option<usize>, Range, usize), String> {
    if payload.len() < 6 {
        return Err("truncated PtgRef3d token".to_string());
    }
    let ixti = u16::from_le_bytes([payload[0], payload[1]]);
    let row = u16::from_le_bytes([payload[2], payload[3]]) as u32;
    let col_field = u16::from_le_bytes([payload[4], payload[5]]);
    let col = (col_field & BIFF8_COL_INDEX_MASK) as u32;

    let sheet =
        resolve_ixti_to_internal_sheet(ixti, externsheets, internal_supbook_index, sheet_count);
    let range = Range::new(CellRef::new(row, col), CellRef::new(row, col));
    Ok((sheet, range, 6))
}

fn decode_ptg_area3d(
    payload: &[u8],
    externsheets: &[externsheet::ExternSheetEntry],
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
    let col_first = (col_first_field & BIFF8_COL_INDEX_MASK) as u32;
    let col_last = (col_last_field & BIFF8_COL_INDEX_MASK) as u32;

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
    externsheets: &[externsheet::ExternSheetEntry],
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
    let is_internal = if xti.supbook == 0 {
        true
    } else if let Some(internal_idx) = internal_supbook_index {
        xti.supbook == internal_idx
    } else {
        false
    };
    if !is_internal {
        return None;
    }
    if xti.itab_first != xti.itab_last {
        return None;
    }

    if xti.itab_first < 0 {
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
    fn decodes_filter_database_area_with_high_col_bits_set() {
        // Some producers set unused high bits in the BIFF8 col fields. `.xls` worksheets are limited
        // to 256 columns, so we treat column indices as 8-bit and ignore those higher bits.
        //
        // Encode `$A$1:$C$5`, but with colLast having an extra bit set (0x0400).
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&0x0402u16.to_le_bytes()); // colLast (C with a high bit set)

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
        name_data.push(0); // chKey
        name_data.push(1); // cch (builtin id length)
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn decodes_filter_database_area_with_trailing_ptg_paren() {
        // Some producers may wrap the reference with an explicit PtgParen token.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast
        rgce.push(0x15); // PtgParen

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
        name_data.push(0); // chKey
        name_data.push(1); // cch (builtin id length)
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn decodes_filter_database_area_with_trailing_ptg_attr() {
        // Some producers may include a trailing PtgAttr token with only formatting/evaluation-hint
        // flags. Accept it as best-effort.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast
        // PtgAttr: [0x19][grbit][wAttr]
        rgce.push(0x19);
        rgce.push(0x00); // grbit
        rgce.extend_from_slice(&0u16.to_le_bytes()); // wAttr

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
        name_data.push(0); // chKey
        name_data.push(1); // cch (builtin id length)
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn decodes_filter_database_from_non_builtin_name_using_filterdatabase_alias() {
        // Some writers store `_FilterDatabase` as a normal string NAME rather than a built-in id.
        // Accept the alias and recover the filter range.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let name_str = FILTER_DATABASE_NAME_ALIAS;
        let cch = name_str.len() as u8;

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&0u16.to_le_bytes()); // grbit (not builtin)
        name_data.push(0); // chKey
        name_data.push(cch); // cch
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(0); // flags (compressed XLUnicodeStringNoCch)
        name_data.extend_from_slice(name_str.as_bytes());
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn strips_embedded_nuls_from_filter_database_name_string() {
        // BIFF string payloads can contain embedded NUL bytes in the wild (and calamine can surface
        // them). Strip them so `_FilterDatabase` name matching works.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let name_str = "_FilterDatabase\0";
        let cch = name_str.len() as u8;

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&0u16.to_le_bytes()); // grbit (not builtin)
        name_data.push(0); // chKey
        name_data.push(cch); // cch
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(0); // flags (compressed XLUnicodeStringNoCch)
        name_data.extend_from_slice(name_str.as_bytes());
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn decodes_filter_database_from_truncated_filterdatabase_alias() {
        // Some BIFF NAME encodings (or decoders) can lose the final `e` in `_FilterDatabase`.
        // Accept the truncated alias so the AutoFilter range is still recovered.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let name_str = FILTER_DATABASE_NAME_ALIAS_TRUNCATED;
        let cch = name_str.len() as u8;

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&0u16.to_le_bytes()); // grbit (not builtin)
        name_data.push(0); // chKey
        name_data.push(cch); // cch
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&1u16.to_le_bytes()); // itab (sheet 1)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(0); // flags (compressed XLUnicodeStringNoCch)
        name_data.extend_from_slice(name_str.as_bytes());
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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
    fn decodes_workbook_scope_filter_database_without_sheet_when_single_sheet() {
        // Some writers store `_FilterDatabase` as workbook-scoped (itab=0) but still use a 2D area
        // token. When there is only one sheet, treat it as the implied target.
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea
        rgce.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        rgce.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        rgce.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        rgce.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let mut name_data = Vec::new();
        name_data.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
        name_data.push(0); // chKey
        name_data.push(1); // cch (builtin id length)
        name_data.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        name_data.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_data.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        name_data.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu..cchStatusText
        name_data.push(BUILTIN_NAME_FILTER_DATABASE); // built-in name id
        name_data.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &bof_globals()),
            record(RECORD_NAME, &name_data),
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

    #[test]
    fn detects_internal_supbook_marker_0401() {
        // SUPBOOK[0]: external workbook marker (minimal, but not internal).
        let supbook_external = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&3u16.to_le_bytes()); // cch
            data.push(0); // flags (compressed)
            data.extend_from_slice(b"ext"); // virtPath
            record(RECORD_SUPBOOK, &data)
        };

        // SUPBOOK[1]: internal workbook marker using a 4-byte payload with sentinel 0x0401.
        let supbook_internal_marker = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // ctab
            data.extend_from_slice(&0x0401u16.to_le_bytes()); // marker
            record(RECORD_SUPBOOK, &data)
        };

        // EXTERNSHEET entry that references SUPBOOK index 1.
        let externsheet = {
            let mut data = Vec::new();
            data.extend_from_slice(&1u16.to_le_bytes()); // cXTI
            data.extend_from_slice(&1u16.to_le_bytes()); // iSupBook (internal supbook index)
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
            supbook_external,
            supbook_internal_marker,
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
