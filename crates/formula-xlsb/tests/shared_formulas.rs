use formula_xlsb::{parse_sheet_bin, CellValue};
use pretty_assertions::assert_eq;
use std::collections::HashMap;
use std::io::Cursor;

fn push_record(out: &mut Vec<u8>, id: u32, data: &[u8]) {
    push_id(out, id);
    push_len(out, data.len() as u32);
    out.extend_from_slice(data);
}

fn push_id(out: &mut Vec<u8>, id: u32) {
    // Mirrors `Biff12Reader::read_id()` (byte-wise little-endian, continuation
    // bit preserved).
    for i in 0..4u32 {
        let byte = ((id >> (8 * i)) & 0xFF) as u8;
        out.push(byte);
        if byte & 0x80 == 0 {
            break;
        }
    }
}

fn push_len(out: &mut Vec<u8>, mut len: u32) {
    // Mirrors `Biff12Reader::read_len()` (LEB128 / 7-bit groups).
    loop {
        let mut byte = (len & 0x7F) as u8;
        len >>= 7;
        if len != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if len == 0 {
            break;
        }
    }
}

#[test]
fn materializes_shared_formulas_via_ptgexp() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0181;
    const WORKSHEET_END: u32 = 0x0182;
    const SHEETDATA_BEGIN: u32 = 0x0191;
    const SHEETDATA_END: u32 = 0x0192;
    const DIMENSION: u32 = 0x0194;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: r1, r2, c1, c2 (all u32).
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=1, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce (best-effort): PtgRefN(row_off=0,col_off=-1) + 1 + +
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&0i32.to_le_bytes());
        v.extend_from_slice(&(-1i16).to_le_bytes());
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // B1 full formula (PtgRef A1 + 1 +)
    let full_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&0u32.to_le_bytes()); // row=0
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col=0, row+col relative
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    let mut b1 = Vec::new();
    b1.extend_from_slice(&1u32.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(full_rgce.len() as u32).to_le_bytes());
    b1.extend_from_slice(&full_rgce);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row 1
    push_record(&mut sheet, ROW, &1u32.to_le_bytes());

    // B2 uses PtgExp to reference base cell B1 (row=0, col=1)
    let ptgexp: [u8; 5] = [0x01, 0x00, 0x00, 0x01, 0x00];
    let mut b2 = Vec::new();
    b2.extend_from_slice(&1u32.to_le_bytes()); // col B
    b2.extend_from_slice(&0u32.to_le_bytes()); // style
    b2.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b2.extend_from_slice(&0u16.to_le_bytes()); // flags
    b2.extend_from_slice(&(ptgexp.len() as u32).to_le_bytes());
    b2.extend_from_slice(&ptgexp);
    push_record(&mut sheet, FMLA_NUM, &b2);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);

    let parsed = parse_sheet_bin(&mut Cursor::new(sheet), &[]).expect("parse synthetic sheet");
    let mut cells: HashMap<(u32, u32), _> = parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(b1.value, CellValue::Number(0.0));
    assert_eq!(b1.formula.as_ref().and_then(|f| f.text.as_deref()), Some("A1+1"));

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(b2.value, CellValue::Number(0.0));
    assert_eq!(b2.formula.as_ref().and_then(|f| f.text.as_deref()), Some("A2+1"));

    // Ensure we stored a materialized rgce (not just PtgExp).
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef
}

