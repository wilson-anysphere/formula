use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use std::io::{Cursor, Write};

fn encode_biff12_id(id: u32) -> Vec<u8> {
    let mut out = Vec::new();
    for i in 0..4 {
        let byte = ((id >> (8 * i)) & 0xFF) as u8;
        out.push(byte);
        if byte & 0x80 == 0 {
            break;
        }
    }
    out
}

fn encode_biff12_len(mut len: u32) -> Vec<u8> {
    let mut out = Vec::new();
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
    out
}

fn biff12_record(id: u32, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&encode_biff12_id(id));
    out.extend_from_slice(&encode_biff12_len(payload.len() as u32));
    out.extend_from_slice(payload);
    out
}

fn encode_utf16_string(s: &str) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut out = Vec::new();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

fn encode_xl_wide_string(
    s: &str,
    flags: u16,
    flags_width: usize,
    rich_runs: Option<&[u8]>,
    phonetic: Option<&[u8]>,
) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut out = Vec::new();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    match flags_width {
        1 => out.push(flags as u8),
        2 => out.extend_from_slice(&flags.to_le_bytes()),
        other => panic!("unexpected flags width {other}"),
    }
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }

    if flags & 0x0001 != 0 {
        let rich = rich_runs.expect("rich flag requires runs");
        assert_eq!(
            rich.len() % 8,
            0,
            "rich run bytes must be a multiple of 8"
        );
        out.extend_from_slice(&((rich.len() / 8) as u32).to_le_bytes());
        out.extend_from_slice(rich);
    }

    if flags & 0x0002 != 0 {
        let pho = phonetic.expect("phonetic flag requires bytes");
        out.extend_from_slice(&(pho.len() as u32).to_le_bytes());
        out.extend_from_slice(pho);
    }

    out
}

fn build_minimal_xlsb(sheet_bin: &[u8]) -> Vec<u8> {
    // workbook.bin contains only a single BrtSheet followed by BrtEndSheets.
    // BrtSheet record data:
    //   [4 bytes unknown][sheetId:u32][relId:XLWideString][name:XLWideString]
    let mut sheet_rec = Vec::new();
    sheet_rec.extend_from_slice(&[0u8; 4]);
    sheet_rec.extend_from_slice(&1u32.to_le_bytes());
    sheet_rec.extend_from_slice(&encode_utf16_string("rId1"));
    sheet_rec.extend_from_slice(&encode_utf16_string("Sheet1"));

    let workbook_bin = [
        biff12_record(0x019C, &sheet_rec),
        biff12_record(0x0190, &[]),
    ]
    .concat();

    let workbook_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>"#;

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let options =
            zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.bin", options).unwrap();
        zip.write_all(&workbook_bin).unwrap();

        zip.start_file("xl/_rels/workbook.bin.rels", options).unwrap();
        zip.write_all(workbook_rels).unwrap();

        zip.start_file("xl/worksheets/sheet1.bin", options).unwrap();
        zip.write_all(sheet_bin).unwrap();

        zip.finish().unwrap();
    }
    cursor.into_inner()
}

#[test]
fn parses_inline_and_formula_strings_with_rich_and_phonetic_extras() {
    // Sheet record ids (subset):
    // - BrtDimension  0x0194
    // - BrtBeginSheetData 0x0191
    // - BrtEndSheetData   0x0192
    // - BrtRow       0x0000
    // - BrtCellSt    0x0006
    // - BrtFmlaString 0x0008
    let mut sheet_bin = Vec::new();

    // Dimension: A1:C1 (r1=0,r2=0,c1=0,c2=2).
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&0u32.to_le_bytes());
    dim.extend_from_slice(&2u32.to_le_bytes());
    sheet_bin.extend_from_slice(&biff12_record(0x0194, &dim));

    sheet_bin.extend_from_slice(&biff12_record(0x0191, &[])); // BrtBeginSheetData

    // Row 0.
    sheet_bin.extend_from_slice(&biff12_record(0x0000, &0u32.to_le_bytes()));

    // A1: inline string containing quotes.
    let inline_text = r#"He said "Hi""#;
    let cell_st = encode_xl_wide_string(inline_text, 0, 1, None, None);
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cell_st);
    sheet_bin.extend_from_slice(&biff12_record(0x0006, &cell_st_payload));

    // B1: formula cached string with rich-text flag set and one dummy run.
    let rich_runs = [0u8; 8];
    let rich_text = "Rich";
    let rich_string = encode_xl_wide_string(rich_text, 0x0001, 2, Some(&rich_runs), None);
    let mut fmla_rich = Vec::new();
    fmla_rich.extend_from_slice(&1u32.to_le_bytes()); // col
    fmla_rich.extend_from_slice(&0u32.to_le_bytes()); // style
    fmla_rich.extend_from_slice(&rich_string);
    fmla_rich.extend_from_slice(&2u32.to_le_bytes()); // cce
    fmla_rich.extend_from_slice(&[0x1D, 0x01]); // rgce: PtgBool TRUE
    sheet_bin.extend_from_slice(&biff12_record(0x0008, &fmla_rich));

    // C1: formula cached string with phonetic flag set and dummy bytes.
    let phonetic_bytes = [1u8, 2, 3, 4, 5];
    let pho_text = "Pho";
    let pho_string = encode_xl_wide_string(pho_text, 0x0002, 2, None, Some(&phonetic_bytes));
    let mut fmla_pho = Vec::new();
    fmla_pho.extend_from_slice(&2u32.to_le_bytes()); // col
    fmla_pho.extend_from_slice(&0u32.to_le_bytes()); // style
    fmla_pho.extend_from_slice(&pho_string);
    fmla_pho.extend_from_slice(&2u32.to_le_bytes()); // cce
    fmla_pho.extend_from_slice(&[0x1D, 0x01]); // rgce: PtgBool TRUE
    sheet_bin.extend_from_slice(&biff12_record(0x0008, &fmla_pho));

    sheet_bin.extend_from_slice(&biff12_record(0x0192, &[])); // BrtEndSheetData

    let xlsb_bytes = build_minimal_xlsb(&sheet_bin);
    let tmp = tempfile::NamedTempFile::new().expect("temp file");
    std::fs::write(tmp.path(), xlsb_bytes).expect("write temp xlsb");

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");

    let mut by_coord = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    assert_eq!(
        by_coord.remove(&(0, 0)).unwrap().value,
        CellValue::Text(inline_text.to_string())
    );

    let rich_cell = by_coord.remove(&(0, 1)).unwrap();
    assert_eq!(rich_cell.value, CellValue::Text(rich_text.to_string()));
    let rich_fmla = rich_cell.formula.as_ref().expect("formula present");
    assert_eq!(rich_fmla.rgce, vec![0x1D, 0x01]);

    let pho_cell = by_coord.remove(&(0, 2)).unwrap();
    assert_eq!(pho_cell.value, CellValue::Text(pho_text.to_string()));
    let pho_fmla = pho_cell.formula.as_ref().expect("formula present");
    assert_eq!(pho_fmla.rgce, vec![0x1D, 0x01]);
}
