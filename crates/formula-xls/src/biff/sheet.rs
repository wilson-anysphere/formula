use std::collections::{BTreeMap, HashMap};

use formula_model::{CellRef, ColProperties, RowProperties, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

use super::records;

// Record ids used by worksheet parsing.
// See [MS-XLS] sections:
// - ROW: 2.4.184
// - COLINFO: 2.4.48
// - Cell records: 2.5.14
// - MULRK: 2.4.141
// - MULBLANK: 2.4.140
const RECORD_ROW: u16 = 0x0208;
const RECORD_COLINFO: u16 = 0x007D;

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
