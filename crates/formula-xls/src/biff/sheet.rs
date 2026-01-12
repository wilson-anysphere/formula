use std::collections::{BTreeMap, HashMap};

use formula_model::{
    CellRef, ColProperties, Hyperlink, HyperlinkTarget, Range, RowProperties, EXCEL_MAX_COLS,
    EXCEL_MAX_ROWS,
};

use super::records;
use super::strings;

// Record ids used by worksheet parsing.
// See [MS-XLS] sections:
// - ROW: 2.4.184
// - COLINFO: 2.4.48
// - Cell records: 2.5.14
// - MULRK: 2.4.141
// - MULBLANK: 2.4.140
const RECORD_ROW: u16 = 0x0208;
const RECORD_COLINFO: u16 = 0x007D;
/// MERGEDCELLS [MS-XLS 2.4.139]
const RECORD_MERGEDCELLS: u16 = 0x00E5;

const RECORD_FORMULA: u16 = 0x0006;
const RECORD_BLANK: u16 = 0x0201;
const RECORD_NUMBER: u16 = 0x0203;
const RECORD_LABEL_BIFF5: u16 = 0x0204;
const RECORD_BOOLERR: u16 = 0x0205;
const RECORD_RK: u16 = 0x027E;
const RECORD_RSTRING: u16 = 0x00D6;
const RECORD_LABELSST: u16 = 0x00FD;
const RECORD_MULRK: u16 = 0x00BD;
const RECORD_MULBLANK: u16 = 0x00BE;
/// HLINK [MS-XLS 2.4.110]
const RECORD_HLINK: u16 = 0x01B8;

const ROW_HEIGHT_TWIPS_MASK: u16 = 0x7FFF;
const ROW_HEIGHT_DEFAULT_FLAG: u16 = 0x8000;
const ROW_OPTION_HIDDEN: u32 = 0x0000_0020;

const COLINFO_OPTION_HIDDEN: u16 = 0x0001;

#[derive(Debug, Default)]
pub(crate) struct SheetRowColProperties {
    pub(crate) rows: BTreeMap<u32, RowProperties>,
    pub(crate) cols: BTreeMap<u32, ColProperties>,
}

pub(crate) fn parse_biff_sheet_row_col_properties(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetRowColProperties, String> {
    let mut props = SheetRowColProperties::default();

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        match record.record_id {
            // ROW [MS-XLS 2.4.184]
            RECORD_ROW => {
                let data = record.data;
                if data.len() < 16 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let height_options = u16::from_le_bytes([data[6], data[7]]);
                let height_twips = height_options & ROW_HEIGHT_TWIPS_MASK;
                let default_height = (height_options & ROW_HEIGHT_DEFAULT_FLAG) != 0;
                let options = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let hidden = (options & ROW_OPTION_HIDDEN) != 0;

                let height =
                    (!default_height && height_twips > 0).then_some(height_twips as f32 / 20.0);

                if hidden || height.is_some() {
                    let entry = props.rows.entry(row).or_default();
                    if let Some(height) = height {
                        entry.height = Some(height);
                    }
                    if hidden {
                        entry.hidden = true;
                    }
                }
            }
            // COLINFO [MS-XLS 2.4.48]
            RECORD_COLINFO => {
                let data = record.data;
                if data.len() < 12 {
                    continue;
                }
                let first_col = u16::from_le_bytes([data[0], data[1]]) as u32;
                let last_col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let width_raw = u16::from_le_bytes([data[4], data[5]]);
                let options = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (options & COLINFO_OPTION_HIDDEN) != 0;

                let width = (width_raw > 0).then_some(width_raw as f32 / 256.0);

                if hidden || width.is_some() {
                    for col in first_col..=last_col {
                        let entry = props.cols.entry(col).or_default();
                        if let Some(width) = width {
                            entry.width = Some(width);
                        }
                        if hidden {
                            entry.hidden = true;
                        }
                    }
                }
            }
            // EOF terminates the sheet substream.
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(props)
}

/// Parse merged cell regions from a worksheet BIFF substream.
///
/// `calamine` usually exposes merge ranges via `worksheet_merge_cells()`, but some `.xls` files in
/// the wild contain `MERGEDCELLS` records that are not surfaced (or surfaced incompletely). This is
/// a best-effort fallback that scans the sheet substream directly and recovers any merge ranges it
/// can.
pub(crate) fn parse_biff_sheet_merged_cells(
    workbook_stream: &[u8],
    start: usize,
) -> Result<Vec<Range>, String> {
    let mut out = Vec::new();

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        match record.record_id {
            RECORD_MERGEDCELLS => {
                // MERGEDCELLS [MS-XLS 2.4.139]
                // - cAreas (2 bytes): number of Ref8 structures
                // - Ref8 (8 bytes each): rwFirst, rwLast, colFirst, colLast (all u16)
                let data = record.data;
                if data.len() < 2 {
                    continue;
                }

                let c_areas = u16::from_le_bytes([data[0], data[1]]) as usize;
                let mut pos = 2usize;
                for _ in 0..c_areas {
                    let Some(chunk) = data.get(pos..pos + 8) else {
                        break;
                    };
                    pos = pos.saturating_add(8);

                    let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]) as u32;
                    let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]) as u32;
                    let col_first = u16::from_le_bytes([chunk[4], chunk[5]]) as u32;
                    let col_last = u16::from_le_bytes([chunk[6], chunk[7]]) as u32;

                    if rw_first >= EXCEL_MAX_ROWS
                        || rw_last >= EXCEL_MAX_ROWS
                        || col_first >= EXCEL_MAX_COLS
                        || col_last >= EXCEL_MAX_COLS
                    {
                        // Ignore out-of-bounds ranges to avoid corrupt coordinates.
                        continue;
                    }

                    out.push(Range::new(
                        CellRef::new(rw_first, col_first),
                        CellRef::new(rw_last, col_last),
                    ));
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

pub(crate) fn parse_biff_sheet_cell_xf_indices_filtered(
    workbook_stream: &[u8],
    start: usize,
    xf_is_interesting: Option<&[bool]>,
) -> Result<HashMap<CellRef, u16>, String> {
    let mut out = HashMap::new();

    let mut maybe_insert = |row: u32, col: u32, xf: u16| {
        if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
            return;
        }
        if let Some(mask) = xf_is_interesting {
            let idx = xf as usize;
            // Retain out-of-range XF indices so callers can surface an aggregated warning.
            if idx >= mask.len() {
                out.insert(CellRef::new(row, col), xf);
                return;
            }
            if !mask[idx] {
                return;
            }
        }
        out.insert(CellRef::new(row, col), xf);
    };

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        let data = record.data;
        match record.record_id {
            // Cell records with a `Cell` header (rw, col, ixfe) [MS-XLS 2.5.14].
            //
            // We only care about extracting the XF index (`ixfe`) so we can resolve
            // number formats from workbook globals.
            RECORD_FORMULA | RECORD_BLANK | RECORD_NUMBER | RECORD_LABEL_BIFF5 | RECORD_BOOLERR
            | RECORD_RK | RECORD_RSTRING | RECORD_LABELSST => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let xf = u16::from_le_bytes([data[4], data[5]]);
                maybe_insert(row, col, xf);
            }
            // MULRK [MS-XLS 2.4.141]
            RECORD_MULRK => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last =
                    u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]) as u32;
                let rk_data = &data[4..data.len().saturating_sub(2)];
                for (idx, chunk) in rk_data.chunks_exact(6).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf);
                }
            }
            // MULBLANK [MS-XLS 2.4.140]
            RECORD_MULBLANK => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last =
                    u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]) as u32;
                let xf_data = &data[4..data.len().saturating_sub(2)];
                for (idx, chunk) in xf_data.chunks_exact(2).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf);
                }
            }
            // EOF terminates the sheet substream.
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

#[derive(Debug, Default)]
pub(crate) struct SheetHyperlinks {
    pub(crate) hyperlinks: Vec<Hyperlink>,
    pub(crate) warnings: Vec<String>,
}

// Hyperlink record bits (linkOpts / grbit). These come from the MS-XLS HLINK record spec, but we
// treat them as best-effort: Excel files in the wild sometimes contain slightly different flag
// combinations depending on link type.
const HLINK_FLAG_HAS_MONIKER: u32 = 0x0000_0001;
const HLINK_FLAG_HAS_LOCATION: u32 = 0x0000_0008;
const HLINK_FLAG_HAS_DISPLAY: u32 = 0x0000_0010;
const HLINK_FLAG_HAS_TOOLTIP: u32 = 0x0000_0020;
const HLINK_FLAG_HAS_TARGET_FRAME: u32 = 0x0000_0080;

// CLSIDs (COM GUIDs) used by hyperlink monikers.
// GUIDs are stored with the first 3 fields little-endian (standard COM GUID layout).
const CLSID_URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B,
    0xA9, 0x0B,
];
const CLSID_FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x46,
];

/// Scan a worksheet BIFF substream for hyperlink records (HLINK, id 0x01B8).
///
/// This is a best-effort parser: malformed records are skipped and surfaced as warnings rather
/// than failing the entire import.
pub(crate) fn parse_biff_sheet_hyperlinks(
    workbook_stream: &[u8],
    start: usize,
    codepage: u16,
) -> Result<SheetHyperlinks, String> {
    let mut out = SheetHyperlinks::default();

    // HLINK records can legally be split across one or more `CONTINUE` records if the hyperlink
    // payload exceeds the BIFF record size limit. Use the logical iterator so we can reassemble
    // those fragments before decoding.
    let allows_continuation = |record_id: u16| record_id == RECORD_HLINK;
    let iter = records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                // Best-effort: stop scanning on malformed record boundaries, but keep any
                // successfully decoded hyperlinks and surface a warning.
                out.warnings.push(format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // BOF indicates the start of a new substream; stop before consuming the next section so we
        // don't attribute later hyperlinks to this worksheet.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_HLINK => match decode_hlink_record(record.data.as_ref(), codepage) {
                Ok(Some(link)) => out.hyperlinks.push(link),
                Ok(None) => {}
                Err(err) => out.warnings.push(format!(
                    "failed to decode HLINK record at offset {}: {err}",
                    record.offset
                )),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn decode_hlink_record(data: &[u8], codepage: u16) -> Result<Option<Hyperlink>, String> {
    // HLINK [MS-XLS 2.4.110]
    // - ref8 (8 bytes): anchor
    // - guid (16 bytes): hyperlink GUID (ignored)
    // - streamVersion (4 bytes): usually 2
    // - linkOpts (4 bytes): flags
    if data.len() < 32 {
        return Err("HLINK record too short".to_string());
    }

    let rw_first = u16::from_le_bytes([data[0], data[1]]) as u32;
    let rw_last = u16::from_le_bytes([data[2], data[3]]) as u32;
    let col_first = u16::from_le_bytes([data[4], data[5]]) as u32;
    let col_last = u16::from_le_bytes([data[6], data[7]]) as u32;

    if rw_first >= EXCEL_MAX_ROWS
        || rw_last >= EXCEL_MAX_ROWS
        || col_first >= EXCEL_MAX_COLS
        || col_last >= EXCEL_MAX_COLS
    {
        // Ignore out-of-bounds anchors to avoid corrupt coordinates.
        return Ok(None);
    }

    let range = Range::new(CellRef::new(rw_first, col_first), CellRef::new(rw_last, col_last));

    // Skip guid (16 bytes).
    let mut pos = 8usize + 16usize;

    let stream_version = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    pos += 4;
    if stream_version != 2 {
        // Non-fatal; continue parsing.
        // Some producers may write a different version, but the layout is usually identical.
    }

    let link_opts = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    pos += 4;

    let mut display: Option<String> = None;
    let mut tooltip: Option<String> = None;
    let mut text_mark: Option<String> = None;
    let mut uri: Option<String> = None;

    // Optional: display string.
    if (link_opts & HLINK_FLAG_HAS_DISPLAY) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        display = (!s.is_empty()).then_some(s);
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: target frame (ignored for now).
    if (link_opts & HLINK_FLAG_HAS_TARGET_FRAME) != 0 {
        let (_s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: moniker (external link target).
    if (link_opts & HLINK_FLAG_HAS_MONIKER) != 0 {
        let (parsed_uri, consumed) = parse_hyperlink_moniker(&data[pos..], codepage)?;
        uri = parsed_uri;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: location / text mark (internal target or sub-address).
    if (link_opts & HLINK_FLAG_HAS_LOCATION) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        text_mark = (!s.is_empty()).then_some(s);
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: tooltip.
    if (link_opts & HLINK_FLAG_HAS_TOOLTIP) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        tooltip = (!s.is_empty()).then_some(s);
        // No further fields depend on the cursor position today, but keep the overflow check so
        // malformed payloads still surface a warning.
        let _ = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    let target = if let Some(uri) = uri {
        if uri.to_ascii_lowercase().starts_with("mailto:") {
            HyperlinkTarget::Email { uri }
        } else {
            HyperlinkTarget::ExternalUrl { uri }
        }
    } else if let Some(mark) = text_mark.as_deref() {
        let (sheet, cell) =
            parse_internal_location(mark).ok_or_else(|| "unsupported internal hyperlink".to_string())?;
        HyperlinkTarget::Internal { sheet, cell }
    } else {
        return Err("HLINK record is missing target information".to_string());
    };

    Ok(Some(Hyperlink {
        range,
        target,
        display,
        tooltip,
        rel_id: None,
    }))
}

fn parse_hyperlink_moniker(input: &[u8], codepage: u16) -> Result<(Option<String>, usize), String> {
    if input.len() < 16 {
        return Err("truncated hyperlink moniker".to_string());
    }
    let clsid: [u8; 16] = input[0..16]
        .try_into()
        .expect("slice length verified");

    // URL moniker: UTF-16LE URL with a 32-bit length prefix.
    if clsid == CLSID_URL_MONIKER {
        if input.len() < 20 {
            return Err("truncated URL moniker".to_string());
        }
        let len = u32::from_le_bytes([input[16], input[17], input[18], input[19]]) as usize;
        let mut consumed = 20usize;

        let (url, url_bytes) = parse_utf16_prefixed_string(&input[20..], len)?;
        consumed = consumed
            .checked_add(url_bytes)
            .ok_or_else(|| "URL moniker length overflow".to_string())?;
        return Ok(((!url.is_empty()).then_some(url), consumed));
    }

    // File moniker: not fully supported yet. Preserve as best-effort file:// URL when we can.
    if clsid == CLSID_FILE_MONIKER {
        // The file moniker payload is more complex (short/long paths, UNC). We attempt a minimal
        // parse that recovers an ANSI/Unicode path string when possible.
        //
        // Best-effort strategy:
        // - The first dword is the length in bytes of the following ANSI path (including NUL).
        // - The ANSI path is followed by optional Unicode extended path data.
        //
        // If this fails, treat as unsupported.
        if input.len() < 20 {
            return Err("truncated file moniker".to_string());
        }
        let ansi_len = u32::from_le_bytes([input[16], input[17], input[18], input[19]]) as usize;
        let mut pos = 20usize;
        if ansi_len > 0 {
            if input.len() < pos + ansi_len {
                return Err("truncated file moniker ANSI path".to_string());
            }
            let bytes = &input[pos..pos + ansi_len];
            pos += ansi_len;
            let mut path = strings::decode_ansi(codepage, bytes);
            path = path.trim_end_matches('\0').to_string();
            if !path.is_empty() {
                // Use a best-effort file URI; Excel will often store DOS paths.
                let uri = format!("file:///{path}");
                return Ok((Some(uri), pos));
            }
        }
        return Err("unsupported file moniker".to_string());
    }

    Err(format!("unsupported hyperlink moniker CLSID {:02X?}", clsid))
}

fn parse_utf16_prefixed_string(input: &[u8], len: usize) -> Result<(String, usize), String> {
    // Heuristic: `len` may be either a byte length or a character count. Prefer byte length when
    // it fits and is even; otherwise treat as chars.
    if len == 0 {
        return Ok((String::new(), 0));
    }
    if len % 2 == 0 && input.len() >= len {
        let bytes = &input[..len];
        let s = decode_utf16le(bytes)?;
        return Ok((trim_trailing_nuls(s), len));
    }

    let byte_len = len
        .checked_mul(2)
        .ok_or_else(|| "string length overflow".to_string())?;
    if input.len() < byte_len {
        return Err("truncated UTF-16 string".to_string());
    }
    let bytes = &input[..byte_len];
    let s = decode_utf16le(bytes)?;
    Ok((trim_trailing_nuls(s), byte_len))
}

fn parse_hyperlink_string(input: &[u8], codepage: u16) -> Result<(String, usize), String> {
    // HyperlinkString [MS-XLS 2.5.??]: cch (u32) + UTF-16LE characters.
    // Some producers may store strings as BIFF8 XLUnicodeString; fall back to that on failure.
    if input.len() >= 4 {
        let cch = u32::from_le_bytes([input[0], input[1], input[2], input[3]]) as usize;
        if cch == 0 {
            return Ok((String::new(), 4));
        }
        if cch <= 1_000_000 {
            if let Some(byte_len) = cch.checked_mul(2) {
                if input.len() >= 4 + byte_len {
                    let bytes = &input[4..4 + byte_len];
                    let s = decode_utf16le(bytes)?;
                    return Ok((trim_trailing_nuls(s), 4 + byte_len));
                }
            }
        }
    }

    // Fallback: BIFF8 XLUnicodeString (u16 length + flags).
    let (s, consumed) = strings::parse_biff8_unicode_string(input, codepage)?;
    Ok((s, consumed))
}

fn decode_utf16le(bytes: &[u8]) -> Result<String, String> {
    if bytes.len() % 2 != 0 {
        return Err("truncated UTF-16 string".to_string());
    }
    let mut u16s = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    Ok(String::from_utf16_lossy(&u16s))
}

fn trim_trailing_nuls(mut s: String) -> String {
    while s.ends_with('\0') {
        s.pop();
    }
    s
}

fn parse_internal_location(location: &str) -> Option<(String, CellRef)> {
    // Mirrors the XLSX hyperlink `location` parsing logic.
    let mut loc = location.trim();
    if let Some(rest) = loc.strip_prefix('#') {
        loc = rest;
    }

    let (sheet, cell) = loc.split_once('!')?;
    let sheet = unquote_sheet_name(sheet.trim());

    let cell_str = cell.trim();
    let cell_str = cell_str
        .split_once(':')
        .map(|(start, _)| start)
        .unwrap_or(cell_str);
    let cell = CellRef::from_a1(cell_str).ok()?;
    Some((sheet, cell))
}

fn unquote_sheet_name(name: &str) -> String {
    // Excel quotes sheet names with single quotes; embedded quotes are doubled.
    let mut s = name.trim();
    if s.starts_with('\'') && s.ends_with('\'') && s.len() >= 2 {
        s = &s[1..s.len() - 1];
        return s.replace("''", "'");
    }
    s.to_string()
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

    #[test]
    fn sheet_row_col_scan_stops_on_truncated_record() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // ROW 1 with explicit height = 20.0 points (400 twips).
        let mut row_payload = [0u8; 16];
        row_payload[0..2].copy_from_slice(&1u16.to_le_bytes());
        row_payload[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_record = record(RECORD_ROW, &row_payload);

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes

        let stream = [sheet_bof, row_record, truncated].concat();
        let props = parse_biff_sheet_row_col_properties(&stream, 0).expect("parse");
        assert_eq!(props.rows.get(&1).and_then(|p| p.height), Some(20.0));
    }

    #[test]
    fn parses_sheet_cell_xf_indices_including_mul_records() {
        // NUMBER cell (A1) with xf=3.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes()); // row
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes()); // col
        number_payload[4..6].copy_from_slice(&3u16.to_le_bytes()); // xf

        // MULBLANK row=1, cols 0..2 with xf {10,11,12}.
        let mut mulblank_payload = Vec::new();
        mulblank_payload.extend_from_slice(&1u16.to_le_bytes()); // row
        mulblank_payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        mulblank_payload.extend_from_slice(&10u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&11u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&12u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&2u16.to_le_bytes()); // colLast

        // MULRK row=2, cols 1..2 with xf {20,21}.
        let mut mulrk_payload = Vec::new();
        mulrk_payload.extend_from_slice(&2u16.to_le_bytes()); // row
        mulrk_payload.extend_from_slice(&1u16.to_le_bytes()); // colFirst
                                                              // cell 1: xf=20 + dummy rk value
        mulrk_payload.extend_from_slice(&20u16.to_le_bytes());
        mulrk_payload.extend_from_slice(&0u32.to_le_bytes());
        // cell 2: xf=21 + dummy rk value
        mulrk_payload.extend_from_slice(&21u16.to_le_bytes());
        mulrk_payload.extend_from_slice(&0u32.to_le_bytes());
        mulrk_payload.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let stream = [
            record(RECORD_NUMBER, &number_payload),
            record(RECORD_MULBLANK, &mulblank_payload),
            record(RECORD_MULRK, &mulrk_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(3));
        assert_eq!(xfs.get(&CellRef::new(1, 0)).copied(), Some(10));
        assert_eq!(xfs.get(&CellRef::new(1, 1)).copied(), Some(11));
        assert_eq!(xfs.get(&CellRef::new(1, 2)).copied(), Some(12));
        assert_eq!(xfs.get(&CellRef::new(2, 1)).copied(), Some(20));
        assert_eq!(xfs.get(&CellRef::new(2, 2)).copied(), Some(21));
    }

    #[test]
    fn parses_mergedcells_records() {
        // First record: A1:B1.
        let mut merged1 = Vec::new();
        merged1.extend_from_slice(&1u16.to_le_bytes()); // cAreas
        merged1.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        merged1.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        merged1.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        merged1.extend_from_slice(&1u16.to_le_bytes()); // colLast

        // Second record: one valid area (C2:D3) and one out-of-bounds (colFirst >= EXCEL_MAX_COLS).
        let mut merged2 = Vec::new();
        merged2.extend_from_slice(&2u16.to_le_bytes()); // cAreas
        // C2:D3 => rows 1..2, cols 2..3 (0-based)
        merged2.extend_from_slice(&1u16.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&2u16.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&2u16.to_le_bytes()); // colFirst
        merged2.extend_from_slice(&3u16.to_le_bytes()); // colLast
        // Out-of-bounds cols.
        merged2.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colFirst (OOB)
        merged2.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colLast (OOB)

        let stream = [
            record(RECORD_MERGEDCELLS, &merged1),
            record(RECORD_MERGEDCELLS, &merged2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let ranges = parse_biff_sheet_merged_cells(&stream, 0).expect("parse");
        assert_eq!(
            ranges,
            vec![
                Range::from_a1("A1:B1").unwrap(),
                Range::from_a1("C2:D3").unwrap(),
            ]
        );
    }

    #[test]
    fn parses_number_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // row
        data.extend_from_slice(&2u16.to_le_bytes()); // col
        data.extend_from_slice(&7u16.to_le_bytes()); // xf
        data.extend_from_slice(&0f64.to_le_bytes()); // value

        let stream = [
            record(RECORD_NUMBER, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(1, 2)).copied(), Some(7));
    }

    #[test]
    fn parses_rk_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&3u16.to_le_bytes()); // row
        data.extend_from_slice(&4u16.to_le_bytes()); // col
        data.extend_from_slice(&9u16.to_le_bytes()); // xf
        data.extend_from_slice(&0u32.to_le_bytes()); // rk

        let stream = [record(RECORD_RK, &data), record(records::RECORD_EOF, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(3, 4)).copied(), Some(9));
    }

    #[test]
    fn parses_blank_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&10u16.to_le_bytes()); // row
        data.extend_from_slice(&3u16.to_le_bytes()); // col
        data.extend_from_slice(&2u16.to_le_bytes()); // xf

        let stream = [
            record(RECORD_BLANK, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(10, 3)).copied(), Some(2));
    }

    #[test]
    fn parses_labelsst_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&0u16.to_le_bytes()); // col
        data.extend_from_slice(&55u16.to_le_bytes()); // xf
        data.extend_from_slice(&123u32.to_le_bytes()); // sst index

        let stream = [
            record(RECORD_LABELSST, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(55));
    }

    #[test]
    fn parses_label_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // row
        data.extend_from_slice(&1u16.to_le_bytes()); // col
        data.extend_from_slice(&77u16.to_le_bytes()); // xf
        data.extend_from_slice(&0u16.to_le_bytes()); // cch (placeholder)

        let stream = [
            record(RECORD_LABEL_BIFF5, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(2, 1)).copied(), Some(77));
    }

    #[test]
    fn parses_boolerr_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&9u16.to_le_bytes()); // row
        data.extend_from_slice(&8u16.to_le_bytes()); // col
        data.extend_from_slice(&5u16.to_le_bytes()); // xf
        data.push(1); // value
        data.push(0); // fErr

        let stream = [
            record(RECORD_BOOLERR, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(9, 8)).copied(), Some(5));
    }

    #[test]
    fn parses_formula_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&4u16.to_le_bytes()); // row
        data.extend_from_slice(&4u16.to_le_bytes()); // col
        data.extend_from_slice(&6u16.to_le_bytes()); // xf
        data.extend_from_slice(&[0u8; 14]); // rest of FORMULA record (dummy)

        let stream = [
            record(RECORD_FORMULA, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(4, 4)).copied(), Some(6));
    }

    #[test]
    fn prefers_last_record_for_duplicate_cells() {
        let blank = {
            let mut data = Vec::new();
            data.extend_from_slice(&0u16.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&1u16.to_le_bytes()); // xf
            record(RECORD_BLANK, &data)
        };

        let number = {
            let mut data = Vec::new();
            data.extend_from_slice(&0u16.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&2u16.to_le_bytes()); // xf
            data.extend_from_slice(&0f64.to_le_bytes());
            record(RECORD_NUMBER, &data)
        };

        let stream = [blank, number, record(records::RECORD_EOF, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(2));
    }

    #[test]
    fn skips_out_of_bounds_cells() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // col (out of bounds)
        data.extend_from_slice(&1u16.to_le_bytes()); // xf

        let stream = [
            record(RECORD_BLANK, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert!(xfs.is_empty());
    }

    #[test]
    fn sheet_row_col_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // ROW 1 with explicit height = 20.0 points (400 twips).
        let mut row_payload = [0u8; 16];
        row_payload[0..2].copy_from_slice(&1u16.to_le_bytes());
        row_payload[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_record = record(RECORD_ROW, &row_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        let stream = [sheet_bof, row_record, next_bof].concat();
        let props = parse_biff_sheet_row_col_properties(&stream, 0).expect("parse");
        assert_eq!(props.rows.get(&1).and_then(|p| p.height), Some(20.0));
    }

    #[test]
    fn sheet_cell_xf_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // NUMBER cell at (0,0) with xf=7.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        number_payload[4..6].copy_from_slice(&7u16.to_le_bytes());
        let number_record = record(RECORD_NUMBER, &number_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        let stream = [sheet_bof, number_record, next_bof].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(7));
    }

    #[test]
    fn sheet_cell_xf_scan_stops_on_truncated_record() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // NUMBER cell at (0,0) with xf=7.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        number_payload[4..6].copy_from_slice(&7u16.to_le_bytes());
        let number_record = record(RECORD_NUMBER, &number_payload);

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes

        let stream = [sheet_bof, number_record, truncated].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(7));
    }
}
