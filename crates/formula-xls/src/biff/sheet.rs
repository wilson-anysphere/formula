use std::collections::{BTreeMap, HashMap};

use formula_model::{CellRef, ColProperties, RowProperties, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

use super::records;

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

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;
    while let Some(record) = iter.next() {
        let record = record?;

        // Stop once we reach the BOF record for the next substream. This allows
        // us to recover row/col metadata even if the worksheet EOF record is
        // missing/corrupt.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            // ROW [MS-XLS 2.4.184]
            0x0208 => {
                let data = record.data;
                if data.len() < 16 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let height_options = u16::from_le_bytes([data[6], data[7]]);
                let height_twips = height_options & 0x7FFF;
                let default_height = (height_options & 0x8000) != 0;
                let options = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let hidden = (options & 0x0000_0020) != 0;

                let height = (!default_height && height_twips > 0)
                    .then_some(height_twips as f32 / 20.0);

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
            0x007D => {
                let data = record.data;
                if data.len() < 12 {
                    continue;
                }
                let first_col = u16::from_le_bytes([data[0], data[1]]) as u32;
                let last_col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let width_raw = u16::from_le_bytes([data[4], data[5]]);
                let options = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (options & 0x0001) != 0;

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
            0x000A => break,
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

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;
    while let Some(record) = iter.next() {
        let record = record?;

        // Stop once we reach the BOF record for the next substream. This allows
        // us to recover XF indices even if the worksheet EOF record is
        // missing/corrupt.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data;
        match record.record_id {
            // Cell records with a `Cell` header (rw, col, ixfe) [MS-XLS 2.5.14].
            //
            // We only care about extracting the XF index (`ixfe`) so we can resolve
            // number formats from workbook globals.
            0x0006 // FORMULA
            | 0x0201 // BLANK
            | 0x0203 // NUMBER
            | 0x0204 // LABEL (BIFF5)
            | 0x0205 // BOOLERR
            | 0x027E // RK
            | 0x00D6 // RSTRING
            | 0x00FD => { // LABELSST
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let xf = u16::from_le_bytes([data[4], data[5]]);
                maybe_insert(row, col, xf);
            }
            // MULRK [MS-XLS 2.4.141]
            0x00BD => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last = u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]])
                    as u32;
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
            0x00BE => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last = u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]])
                    as u32;
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
            0x000A => break,
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
            record(0x0203, &number_payload),
            record(0x00BE, &mulblank_payload),
            record(0x00BD, &mulrk_payload),
            record(0x000A, &[]),
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

        let stream = [record(0x0203, &data), record(0x000A, &[])].concat();
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

        let stream = [record(0x027E, &data), record(0x000A, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(3, 4)).copied(), Some(9));
    }

    #[test]
    fn parses_blank_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&10u16.to_le_bytes()); // row
        data.extend_from_slice(&3u16.to_le_bytes()); // col
        data.extend_from_slice(&2u16.to_le_bytes()); // xf

        let stream = [record(0x0201, &data), record(0x000A, &[])].concat();
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

        let stream = [record(0x00FD, &data), record(0x000A, &[])].concat();
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

        let stream = [record(0x0204, &data), record(0x000A, &[])].concat();
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

        let stream = [record(0x0205, &data), record(0x000A, &[])].concat();
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

        let stream = [record(0x0006, &data), record(0x000A, &[])].concat();
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
            record(0x0201, &data)
        };

        let number = {
            let mut data = Vec::new();
            data.extend_from_slice(&0u16.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&2u16.to_le_bytes()); // xf
            data.extend_from_slice(&0f64.to_le_bytes());
            record(0x0203, &data)
        };

        let stream = [blank, number, record(0x000A, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(2));
    }

    #[test]
    fn skips_out_of_bounds_cells() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // col (out of bounds)
        data.extend_from_slice(&1u16.to_le_bytes()); // xf

        let stream = [record(0x0201, &data), record(0x000A, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert!(xfs.is_empty());
    }

    #[test]
    fn sheet_row_col_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(0x0809, &[0u8; 16]);

        // ROW 1 with explicit height = 20.0 points (400 twips).
        let mut row_payload = [0u8; 16];
        row_payload[0..2].copy_from_slice(&1u16.to_le_bytes());
        row_payload[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_record = record(0x0208, &row_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(0x0809, &[0u8; 16]);

        let stream = [sheet_bof, row_record, next_bof].concat();
        let props = parse_biff_sheet_row_col_properties(&stream, 0).expect("parse");
        assert_eq!(props.rows.get(&1).and_then(|p| p.height), Some(20.0));
    }

    #[test]
    fn sheet_cell_xf_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(0x0809, &[0u8; 16]);

        // NUMBER cell at (0,0) with xf=7.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        number_payload[4..6].copy_from_slice(&7u16.to_le_bytes());
        let number_record = record(0x0203, &number_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(0x0809, &[0u8; 16]);

        let stream = [sheet_bof, number_record, next_bof].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(7));
    }
}
