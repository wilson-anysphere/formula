use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::{biff12_varint, parse_sheet_bin, parse_sheet_bin_with_context, CellValue};
use pretty_assertions::assert_eq;
use std::collections::HashMap;
use std::io::Cursor;

fn push_record(out: &mut Vec<u8>, id: u32, data: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    let len = u32::try_from(data.len()).expect("record too large");
    biff12_varint::write_record_len(out, len).expect("write record len");
    out.extend_from_slice(data);
}

#[test]
fn materializes_shared_formulas_with_ptgexp_coordinate_payload_layouts() {
    // `PtgExp` appears in the wild with multiple coordinate payload shapes:
    // - BIFF12-ish: row u32 + col u32 (8 bytes)
    // - BIFF12-ish: row u32 + col u16 (6 bytes)
    // - BIFF8-ish: row u16 + col u16 (4 bytes)
    // Additionally, some producers include extra trailing bytes after the coordinates.
    //
    // This test synthesizes a single worksheet stream with one shared-formula anchor and
    // multiple dependent cells, each using a different `PtgExp` payload layout.

    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B5 (rows 0..4, col 1).
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&4u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B5:
    //   B1: A1+1
    //   B2: A2+1
    //   ...
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=4, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&4u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce: PtgRefN(row_off=0,col_off=-1) + 1 + +
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

    // Helper to emit a BrtFmlaNum cell in column B for a given row.
    let mut push_ptgexp_cell = |row_idx: u32, ptgexp: &[u8]| {
        push_record(&mut sheet, ROW, &row_idx.to_le_bytes());
        let mut cell = Vec::new();
        cell.extend_from_slice(&1u32.to_le_bytes()); // col B
        cell.extend_from_slice(&0u32.to_le_bytes()); // style
        cell.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
        cell.extend_from_slice(&0u16.to_le_bytes()); // flags
        cell.extend_from_slice(&(ptgexp.len() as u32).to_le_bytes());
        cell.extend_from_slice(ptgexp);
        push_record(&mut sheet, FMLA_NUM, &cell);
    };

    // B2: BIFF8-style (row u16 + col u16).
    let ptgexp_u16_u16: [u8; 5] = [0x01, 0x00, 0x00, 0x01, 0x00];
    push_ptgexp_cell(1, &ptgexp_u16_u16);

    // B3: BIFF12-ish (row u32 + col u16).
    let ptgexp_u32_u16: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01);
        v.extend_from_slice(&0u32.to_le_bytes());
        v.extend_from_slice(&1u16.to_le_bytes());
        v
    };
    push_ptgexp_cell(2, &ptgexp_u32_u16);

    // B4: BIFF12-ish (row u32 + col u32).
    let ptgexp_u32_u32: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01);
        v.extend_from_slice(&0u32.to_le_bytes());
        v.extend_from_slice(&1u32.to_le_bytes());
        v
    };
    push_ptgexp_cell(3, &ptgexp_u32_u32);

    // B5: row u32 + col u32, plus extra trailing bytes (seen in the wild).
    let ptgexp_u32_u32_trailing: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01);
        v.extend_from_slice(&0u32.to_le_bytes());
        v.extend_from_slice(&1u32.to_le_bytes());
        v.extend_from_slice(&[0xAA, 0xBB, 0xCC]); // trailing bytes
        v
    };
    push_ptgexp_cell(4, &ptgexp_u32_u32_trailing);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);

    let parsed = parse_sheet_bin(&mut Cursor::new(sheet), &[]).expect("parse synthetic sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A1+1")
    );

    for (row, expected) in &[(1u32, "A2+1"), (2, "A3+1"), (3, "A4+1"), (4, "A5+1")] {
        let cell = cells.remove(&(*row, 1)).expect("dependent cell present");
        assert_eq!(
            cell.formula.as_ref().and_then(|f| f.text.as_deref()),
            Some(*expected),
            "row {row}"
        );
        // Ensure the shared formula was materialized (not left as PtgExp).
        assert_eq!(
            cell.formula.as_ref().unwrap().rgce.first().copied(),
            Some(0x24),
            "row {row}"
        );
    }
}

#[test]
fn materializes_shared_formulas_with_ptgexp_candidate_fallback_to_biff8_layout() {
    // When a `PtgExp` payload is 6 bytes long, it could be interpreted as:
    // - row u32 + col u16 (newer BIFF12-ish layout), or
    // - row u16 + col u16 (legacy BIFF8 layout), with 2 trailing bytes.
    //
    // `parse_ptg_exp_candidates` returns both candidates and the shared-formula materializer
    // should try them in order until one matches an actual `BrtShrFmla` anchor.

    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B2 (rows 0..1, col 1).
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&0u32.to_le_bytes()); // r1
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // r2
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // c1
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // c2

    // Base rgce: PtgRefN(row_off=0,col_off=-1) + 1 + +
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

    // B1 full formula (PtgRef A1 + 1 +) so the anchor cell also has decoded text.
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

    // B2 uses a 6-byte `PtgExp` payload whose *u32/u16* interpretation points at a non-existent
    // anchor (row 65536), while the BIFF8 u16/u16 interpretation points at the real anchor (row 0).
    //
    // Payload bytes (after 0x01):
    //   row u32 = 0x0001_0000 (LE: 00 00 01 00)  -> u16/u16 sees row=0x0000, col=0x0001
    //   col u16 = 0x0001
    let ptgexp_ambiguous: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01); // PtgExp
        v.extend_from_slice(&0x0001_0000u32.to_le_bytes());
        v.extend_from_slice(&1u16.to_le_bytes());
        v
    };
    let mut b2 = Vec::new();
    b2.extend_from_slice(&1u32.to_le_bytes()); // col B
    b2.extend_from_slice(&0u32.to_le_bytes()); // style
    b2.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b2.extend_from_slice(&0u16.to_le_bytes()); // flags
    b2.extend_from_slice(&(ptgexp_ambiguous.len() as u32).to_le_bytes());
    b2.extend_from_slice(&ptgexp_ambiguous);
    push_record(&mut sheet, FMLA_NUM, &b2);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);

    let parsed = parse_sheet_bin(&mut Cursor::new(sheet), &[]).expect("parse synthetic sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );
    assert_eq!(
        b2.formula.as_ref().unwrap().rgce.first().copied(),
        Some(0x24)
    ); // PtgRef
}

#[test]
fn materializes_shared_formulas_with_ptgexp_u32_row_above_u16_max() {
    // Ensure shared-formula materialization works when the `PtgExp` base row exceeds the BIFF8
    // u16 row limit (65535). In those cases, the u16/u16 interpretation should *not* match an
    // anchor, and we rely on the u32-based candidate.

    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let base_row: u32 = 70_000;
    let dep_row: u32 = base_row + 1;
    let base_col: u32 = 1; // column B

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B{base_row+1}:B{dep_row+1}.
    let mut dim = Vec::new();
    dim.extend_from_slice(&base_row.to_le_bytes());
    dim.extend_from_slice(&dep_row.to_le_bytes());
    dim.extend_from_slice(&base_col.to_le_bytes());
    dim.extend_from_slice(&base_col.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Base row.
    push_record(&mut sheet, ROW, &base_row.to_le_bytes());

    // Shared formula over B{base_row+1}:B{dep_row+1}:
    //   B{base_row+1}: A{base_row+1}+1
    //   B{dep_row+1}:  A{dep_row+1}+1
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&base_row.to_le_bytes()); // r1
    shr_fmla.extend_from_slice(&dep_row.to_le_bytes()); // r2
    shr_fmla.extend_from_slice(&base_col.to_le_bytes()); // c1
    shr_fmla.extend_from_slice(&base_col.to_le_bytes()); // c2

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

    // Base cell B{base_row+1} full formula (PtgRef A{base_row+1} + 1 +)
    let full_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&base_row.to_le_bytes()); // row (u32)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col=0 (A), row+col relative
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    let mut b_base = Vec::new();
    b_base.extend_from_slice(&base_col.to_le_bytes());
    b_base.extend_from_slice(&0u32.to_le_bytes()); // style
    b_base.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b_base.extend_from_slice(&0u16.to_le_bytes()); // flags
    b_base.extend_from_slice(&(full_rgce.len() as u32).to_le_bytes());
    b_base.extend_from_slice(&full_rgce);
    push_record(&mut sheet, FMLA_NUM, &b_base);

    // Dependent row.
    push_record(&mut sheet, ROW, &dep_row.to_le_bytes());

    // Dependent cell uses PtgExp with the u32/u16 payload layout: [row u32][col u16]
    let ptgexp_u32_u16: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01); // PtgExp
        v.extend_from_slice(&base_row.to_le_bytes());
        v.extend_from_slice(&(base_col as u16).to_le_bytes());
        v
    };
    let mut b_dep = Vec::new();
    b_dep.extend_from_slice(&base_col.to_le_bytes());
    b_dep.extend_from_slice(&0u32.to_le_bytes()); // style
    b_dep.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b_dep.extend_from_slice(&0u16.to_le_bytes()); // flags
    b_dep.extend_from_slice(&(ptgexp_u32_u16.len() as u32).to_le_bytes());
    b_dep.extend_from_slice(&ptgexp_u32_u16);
    push_record(&mut sheet, FMLA_NUM, &b_dep);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);

    let parsed = parse_sheet_bin(&mut Cursor::new(sheet), &[]).expect("parse synthetic sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b_base = cells.remove(&(base_row, base_col)).expect("base cell present");
    assert_eq!(
        b_base.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A70001+1")
    );

    let b_dep = cells.remove(&(dep_row, base_col)).expect("dependent cell present");
    assert_eq!(
        b_dep.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A70002+1")
    );
    assert_eq!(b_dep.formula.as_ref().unwrap().rgce.first().copied(), Some(0x24)); // PtgRef
}

#[test]
fn materializes_shared_formulas_via_ptgexp() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(b1.value, CellValue::Number(0.0));
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A1+1")
    );

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(b2.value, CellValue::Number(0.0));
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );

    // Ensure we stored a materialized rgce (not just PtgExp).
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef
}

#[test]
fn materializes_shared_formulas_with_ptgexp_trailing_bytes() {
    // Some producers include extra bytes after the PtgExp coordinates. The shared-formula
    // materializer should still find the correct base cell and expand the shared rgce.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

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

    // B2 uses PtgExp to reference base cell B1 (row=0,col=1), but includes 2 extra bytes.
    let ptgexp: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x01); // PtgExp
        v.extend_from_slice(&0u32.to_le_bytes()); // base row (u32)
        v.extend_from_slice(&1u16.to_le_bytes()); // base col (u16)
        v.extend_from_slice(&0x1234u16.to_le_bytes()); // trailing bytes
        v
    };
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );
    assert_eq!(
        b2.formula.as_ref().unwrap().rgce.first().copied(),
        Some(0x24)
    ); // PtgRef
}

#[test]
fn materializes_shared_formulas_with_ptgfuncvar() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

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
    //   B1: SUM(A1,1)
    //   B2: SUM(A2,1)
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=1, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce: PtgRefN(row_off=0,col_off=-1), 1, SUM(argc=2)
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&0i32.to_le_bytes());
        v.extend_from_slice(&(-1i16).to_le_bytes());
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x22); // PtgFuncVar
        v.push(2); // argc
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // SUM
        v
    };
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // B1 full formula (PtgRef A1, 1, SUM(argc=2))
    let full_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&0u32.to_le_bytes()); // row=0
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col=0, row+col relative
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x22); // PtgFuncVar
        v.push(2); // argc
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // SUM
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(b1.value, CellValue::Number(0.0));
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("SUM(A1,1)")
    );

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(b2.value, CellValue::Number(0.0));
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("SUM(A2,1)")
    );

    // Ensure we stored a materialized rgce (not just PtgExp).
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef
}

#[test]
fn materializes_shared_formulas_with_ptgmemarean() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: r1, r2, c1, c2 (all u32). Cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
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

    // Base rgce: PtgRefN(row_off=0,col_off=-1) + PtgMemAreaN + 1 + +
    // The mem token is ignored for printing but must be preserved to keep offsets aligned.
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&0i32.to_le_bytes());
        v.extend_from_slice(&(-1i16).to_le_bytes());
        v.push(0x2E); // PtgMemAreaN
        v.extend_from_slice(&0u16.to_le_bytes()); // cce (unused for printing)
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );

    // Ensure we stored a materialized rgce containing the mem token.
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef
    assert_eq!(b2_rgce.get(7).copied(), Some(0x2E)); // PtgMemAreaN
}

#[test]
fn materializes_shared_formulas_with_ptgmem_subexpression_refs() {
    // Like `materializes_shared_formulas_with_ptgmemarean`, but exercise the BIFF12 layout where
    // the mem token is followed by `cce` bytes of an rgce subexpression. Shared-formula
    // materialization must shift relative refs inside that nested rgce too.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    let base_rgce: Vec<u8> = {
        // Nested rgce stored in the PtgMem* payload:
        // - PtgRefN(row_off=0,col_off=-1) -> points at A{row} for column B.
        let mem_rgce: Vec<u8> = {
            let mut v = Vec::new();
            v.push(0x2C); // PtgRefN
            v.extend_from_slice(&0i32.to_le_bytes());
            v.extend_from_slice(&(-1i16).to_le_bytes());
            v
        };

        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&0i32.to_le_bytes());
        v.extend_from_slice(&(-1i16).to_le_bytes());
        v.push(0x2E); // PtgMemAreaN
        v.extend_from_slice(&(mem_rgce.len() as u16).to_le_bytes()); // cce
        v.extend_from_slice(&mem_rgce);
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );

    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    // Starts with the materialized `A2` reference.
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef

    // Followed by PtgMemAreaN at byte 7 (PtgRef payload is 6 bytes).
    let mem_offset = 7usize;
    assert_eq!(b2_rgce.get(mem_offset).copied(), Some(0x2E)); // PtgMemAreaN

    // PtgMem* layout: [ptg][cce: u16][nested rgce...]
    let cce = u16::from_le_bytes([b2_rgce[mem_offset + 1], b2_rgce[mem_offset + 2]]) as usize;
    assert_eq!(cce, 7);

    // The nested rgce should have been materialized too: `A2` (row 1, col 0).
    let nested_offset = mem_offset + 1 + 2;
    assert_eq!(b2_rgce.get(nested_offset).copied(), Some(0x24)); // PtgRef
    let nested_row = u32::from_le_bytes(
        b2_rgce[nested_offset + 1..nested_offset + 5]
            .try_into()
            .unwrap(),
    );
    let nested_col = u16::from_le_bytes(
        b2_rgce[nested_offset + 5..nested_offset + 7]
            .try_into()
            .unwrap(),
    );
    assert_eq!(nested_row, 1);
    assert_eq!(nested_col, 0xC000); // col=0, row+col relative.

    // Ensure the output rgce stayed aligned after the mem subexpression.
    assert_eq!(
        b2_rgce.get(nested_offset + cce).copied(),
        Some(0x1E) // PtgInt
    );
}

#[test]
fn materializes_shared_formulas_with_ptgmemfunc_subexpression_refs() {
    // Same as `materializes_shared_formulas_with_ptgmem_subexpression_refs`, but use PtgMemFunc
    // (0x29) instead of PtgMemAreaN (0x2E). This covers another common PtgMem* variant.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: A1+1
    //   B2: A2+1
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    let base_rgce: Vec<u8> = {
        let mem_rgce: Vec<u8> = {
            let mut v = Vec::new();
            v.push(0x2C); // PtgRefN
            v.extend_from_slice(&0i32.to_le_bytes());
            v.extend_from_slice(&(-1i16).to_le_bytes());
            v
        };

        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN (main operand)
        v.extend_from_slice(&0i32.to_le_bytes());
        v.extend_from_slice(&(-1i16).to_le_bytes());
        v.push(0x29); // PtgMemFunc
        v.extend_from_slice(&(mem_rgce.len() as u16).to_le_bytes()); // cce
        v.extend_from_slice(&mem_rgce);
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("A2+1")
    );

    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x24)); // PtgRef

    // PtgMemFunc should appear after the first PtgRef token (6-byte payload).
    let mem_offset = 7usize;
    assert_eq!(b2_rgce.get(mem_offset).copied(), Some(0x29)); // PtgMemFunc

    let cce = u16::from_le_bytes([b2_rgce[mem_offset + 1], b2_rgce[mem_offset + 2]]) as usize;
    assert_eq!(cce, 7);

    let nested_offset = mem_offset + 1 + 2;
    assert_eq!(b2_rgce.get(nested_offset).copied(), Some(0x24)); // nested PtgRef
    let nested_row = u32::from_le_bytes(
        b2_rgce[nested_offset + 1..nested_offset + 5]
            .try_into()
            .unwrap(),
    );
    assert_eq!(nested_row, 1);
    assert_eq!(
        b2_rgce.get(nested_offset + cce).copied(),
        Some(0x1E) // PtgInt
    );
}

#[test]
fn materializes_shared_formulas_out_of_bounds_refs_as_ref_error() {
    // When a shared formula is filled near the sheet boundaries, relative references in the base
    // rgce can point outside the valid row/col range. Excel represents those as `#REF!` tokens
    // (`PtgRefErr` / `PtgAreaErr`). The materializer should do the same instead of giving up and
    // leaving the cell with an unresolved `PtgExp`.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    // Excel max row index (0-based).
    const MAX_ROW: u32 = 1_048_575;

    // Shared formula range: B1048575:B1048576 (rows MAX_ROW-1..MAX_ROW, col=1).
    let base_row = MAX_ROW - 1;
    let base_col = 1u32;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    let mut dim = Vec::new();
    dim.extend_from_slice(&base_row.to_le_bytes()); // r1
    dim.extend_from_slice(&MAX_ROW.to_le_bytes()); // r2
    dim.extend_from_slice(&base_col.to_le_bytes()); // c1
    dim.extend_from_slice(&base_col.to_le_bytes()); // c2
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Shared formula definition:
    //   B1048575: A1048576+1
    //   B1048576: #REF!+1 (because A1048577 is out of bounds)
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&base_row.to_le_bytes()); // r1
    shr_fmla.extend_from_slice(&MAX_ROW.to_le_bytes()); // r2
    shr_fmla.extend_from_slice(&base_col.to_le_bytes()); // c1
    shr_fmla.extend_from_slice(&base_col.to_le_bytes()); // c2

    // Base rgce: PtgRefN(row_off=+1,col_off=-1) + 1 + +
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x2C); // PtgRefN
        v.extend_from_slice(&1i32.to_le_bytes()); // row + 1
        v.extend_from_slice(&(-1i16).to_le_bytes()); // col - 1 (B -> A)
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // Row = base_row
    push_record(&mut sheet, ROW, &base_row.to_le_bytes());

    // Base cell full formula for B1048575: A1048576+1
    let full_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x24); // PtgRef
        v.extend_from_slice(&MAX_ROW.to_le_bytes()); // row = MAX_ROW (A1048576)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    let mut b1 = Vec::new();
    b1.extend_from_slice(&base_col.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(full_rgce.len() as u32).to_le_bytes());
    b1.extend_from_slice(&full_rgce);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row = MAX_ROW
    push_record(&mut sheet, ROW, &MAX_ROW.to_le_bytes());

    // Second cell uses PtgExp referencing the base cell (row=u32, col=u16).
    let mut ptgexp = Vec::new();
    ptgexp.push(0x01); // PtgExp
    ptgexp.extend_from_slice(&base_row.to_le_bytes());
    ptgexp.extend_from_slice(&(base_col as u16).to_le_bytes());

    let mut b2 = Vec::new();
    b2.extend_from_slice(&base_col.to_le_bytes()); // col B
    b2.extend_from_slice(&0u32.to_le_bytes()); // style
    b2.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b2.extend_from_slice(&0u16.to_le_bytes()); // flags
    b2.extend_from_slice(&(ptgexp.len() as u32).to_le_bytes());
    b2.extend_from_slice(&ptgexp);
    push_record(&mut sheet, FMLA_NUM, &b2);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);

    let parsed = parse_sheet_bin(&mut Cursor::new(sheet), &[]).expect("parse synthetic sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b2 = cells
        .remove(&(MAX_ROW, base_col))
        .expect("B1048576 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("#REF!+1")
    );

    // Ensure we stored a materialized rgce (not just PtgExp), starting with PtgRefErr.
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x2A)); // PtgRefErr
}

#[test]
fn materializes_shared_formulas_with_ptgref3d() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    // Provide workbook context so `PtgRef3d` can decode to a sheet name.
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: r1, r2, c1, c2 (all u32). Cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: Sheet2!A1+1
    //   B2: Sheet2!A2+1
    //
    // Shared formulas can encode 3D references using `PtgRef3d` with the same row/col relative
    // flags as `PtgRef`. Materialization must shift those refs across the range.
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=1, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce: PtgRef3d(Sheet2!A1, relative row/col) + 1 + +
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x3A); // PtgRef3d
        v.extend_from_slice(&0u16.to_le_bytes()); // ixti (Sheet2)
        v.extend_from_slice(&0u32.to_le_bytes()); // row=0 (A1)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // col=A, row+col relative
        v.push(0x1E); // PtgInt
        v.extend_from_slice(&1u16.to_le_bytes());
        v.push(0x03); // PtgAdd
        v
    };
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // B1 full formula (same as base rgce).
    let mut b1 = Vec::new();
    b1.extend_from_slice(&1u32.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    b1.extend_from_slice(&base_rgce);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row 1
    push_record(&mut sheet, ROW, &1u32.to_le_bytes());

    // B2 uses PtgExp to reference base cell B1 (row=0,col=1).
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

    let parsed =
        parse_sheet_bin_with_context(&mut Cursor::new(sheet), &[], &ctx).expect("parse sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("Sheet2!A1+1")
    );

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("Sheet2!A2+1")
    );

    // Ensure we stored a materialized rgce with an adjusted row field (A2).
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(
        b2_rgce,
        &vec![
            0x3A, // PtgRef3d
            0x00, 0x00, // ixti (Sheet2)
            0x01, 0x00, 0x00, 0x00, // row=1 (A2)
            0x00, 0xC0, // col=A, row+col relative
            0x1E, // PtgInt
            0x01, 0x00, // 1
            0x03, // +
        ]
    );
}

#[test]
fn materializes_shared_formulas_with_ptgarea3d() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    // Provide workbook context so `PtgArea3d` can decode to a sheet name.
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: r1, r2, c1, c2 (all u32). Cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    dim.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2:
    //   B1: SUM(Sheet2!A1:A2)
    //   B2: SUM(Sheet2!A2:A3)
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=1, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce: PtgArea3d(Sheet2!A1:A2, relative endpoints) + SUM(argc=1)
    let base_rgce: Vec<u8> = {
        let mut v = Vec::new();
        v.push(0x3B); // PtgArea3d
        v.extend_from_slice(&0u16.to_le_bytes()); // ixti (Sheet2)
        v.extend_from_slice(&0u32.to_le_bytes()); // r1 (A1)
        v.extend_from_slice(&1u32.to_le_bytes()); // r2 (A2)
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // c1=A, row+col relative
        v.extend_from_slice(&0xC000u16.to_le_bytes()); // c2=A, row+col relative
        v.push(0x22); // PtgFuncVar
        v.push(1); // argc=1
        v.extend_from_slice(&0x0004u16.to_le_bytes()); // SUM
        v
    };
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // B1 full formula (same as base rgce).
    let mut b1 = Vec::new();
    b1.extend_from_slice(&1u32.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    b1.extend_from_slice(&base_rgce);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row 1
    push_record(&mut sheet, ROW, &1u32.to_le_bytes());

    // B2 uses PtgExp to reference base cell B1 (row=0,col=1).
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

    let parsed =
        parse_sheet_bin_with_context(&mut Cursor::new(sheet), &[], &ctx).expect("parse sheet");
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("SUM(Sheet2!A1:A2)")
    );

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("SUM(Sheet2!A2:A3)")
    );

    // Ensure we stored a materialized rgce with adjusted area rows.
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(
        b2_rgce,
        &vec![
            0x3B, // PtgArea3d
            0x00, 0x00, // ixti (Sheet2)
            0x01, 0x00, 0x00, 0x00, // r1=1 (A2)
            0x02, 0x00, 0x00, 0x00, // r2=2 (A3)
            0x00, 0xC0, // c1=A, relative
            0x00, 0xC0, // c2=A, relative
            0x22, // PtgFuncVar
            0x01, // argc
            0x04, 0x00, // SUM
        ]
    );
}

#[test]
fn materializes_shared_formulas_with_ptgarray_rgcb() {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: r1, r2, c1, c2 (all u32). Cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes()); // r1
    dim.extend_from_slice(&1u32.to_le_bytes()); // r2
    dim.extend_from_slice(&1u32.to_le_bytes()); // c1 (B)
    dim.extend_from_slice(&1u32.to_le_bytes()); // c2 (B)
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Shared formula over B1:B2 whose body is the array constant `{1,2;3,4}`.
    let mut shr_fmla = Vec::new();
    // Range: r1=0, r2=1, c1=1, c2=1.
    shr_fmla.extend_from_slice(&0u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());
    shr_fmla.extend_from_slice(&1u32.to_le_bytes());

    // Base rgce: PtgArray + 7 reserved bytes. The array payload lives in rgcb.
    let base_rgce: Vec<u8> = vec![0x20, 0, 0, 0, 0, 0, 0, 0];
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(&base_rgce);

    // rgcb encoding for {1,2;3,4} (2x2, row-major numeric values).
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 rows
    for v in [1.0f64, 2.0, 3.0, 4.0] {
        rgcb.push(0x01); // xltypeNum
        rgcb.extend_from_slice(&v.to_le_bytes());
    }
    shr_fmla.extend_from_slice(&rgcb);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // Row 0 (B1).
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());
    // PtgExp referencing base cell B1 (row=0,col=1).
    let ptgexp: [u8; 5] = [0x01, 0x00, 0x00, 0x01, 0x00];
    let mut b1 = Vec::new();
    b1.extend_from_slice(&1u32.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(ptgexp.len() as u32).to_le_bytes());
    b1.extend_from_slice(&ptgexp);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row 1 (B2).
    push_record(&mut sheet, ROW, &1u32.to_le_bytes());
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
    let mut cells: HashMap<(u32, u32), _> =
        parsed.cells.iter().map(|c| ((c.row, c.col), c)).collect();

    let b1 = cells.remove(&(0, 1)).expect("B1 present");
    assert_eq!(b1.value, CellValue::Number(0.0));
    assert_eq!(
        b1.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("{1,2;3,4}")
    );

    let b2 = cells.remove(&(1, 1)).expect("B2 present");
    assert_eq!(b2.value, CellValue::Number(0.0));
    assert_eq!(
        b2.formula.as_ref().and_then(|f| f.text.as_deref()),
        Some("{1,2;3,4}")
    );

    // Ensure we stored a materialized rgce (not just PtgExp).
    let b2_rgce = &b2.formula.as_ref().unwrap().rgce;
    assert_eq!(b2_rgce.first().copied(), Some(0x20)); // PtgArray
}
