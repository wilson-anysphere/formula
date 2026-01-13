use std::io::{Cursor, Write};

use formula_xlsb::biff12_varint;
use formula_xlsb::rgce::{decode_rgce_with_context, CellCoord};
#[cfg(feature = "write")]
use formula_xlsb::rgce::encode_rgce_with_context_ast;
#[cfg(not(feature = "write"))]
use formula_xlsb::rgce::encode_rgce_with_context;
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;
use tempfile::NamedTempFile;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

// BIFF record IDs used by this minimal fixture. (Not re-exported from the crate.)
mod biff12 {
    pub const SHEET: u32 = 0x009C;

    // External-link records (BIFF8-era ids, still observed in XLSB).
    pub const SUPBOOK: u32 = 0x00AE;
    pub const SUPBOOK_END: u32 = 0x00AF;
    pub const EXTERN_SHEET: u32 = 0x0017;
    pub const EXTERN_NAME: u32 = 0x0023;
}

fn write_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    let len = u32::try_from(payload.len()).expect("record too large");
    biff12_varint::write_record_len(out, len).expect("write record len");
    out.extend_from_slice(payload);
}

fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    let len = u32::try_from(units.len()).expect("string too large");
    out.extend_from_slice(&len.to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
}

fn build_workbook_bin() -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    // Sheet metadata (3 sheets).
    for (idx, sheet_name) in ["Sheet1", "Sheet2", "Sheet3"].iter().enumerate() {
        let mut sheet = Vec::<u8>::new();
        sheet.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet.extend_from_slice(&(idx as u32 + 1).to_le_bytes()); // sheet id

        let rid = format!("rId{}", idx + 1);
        write_utf16_string(&mut sheet, &rid);
        write_utf16_string(&mut sheet, sheet_name);

        write_record(&mut out, biff12::SHEET, &sheet);
    }

    // Internal SupBook (raw_name="") + a single ExternName for "MyName".
    //
    // This combination allows encoding a 3D name reference (`Sheet1:Sheet3!MyName`) via
    // PtgNameX(ixti, nameIndex).
    let mut supbook = Vec::<u8>::new();
    supbook.extend_from_slice(&0u16.to_le_bytes()); // ctab=0 (no sheet name list)
    write_utf16_string(&mut supbook, ""); // raw_name (empty => internal)
    write_record(&mut out, biff12::SUPBOOK, &supbook);

    let mut extern_name = Vec::<u8>::new();
    extern_name.extend_from_slice(&0u16.to_le_bytes()); // flags
    extern_name.extend_from_slice(&0xFFFFu16.to_le_bytes()); // scope = none
    write_utf16_string(&mut extern_name, "MyName");
    write_record(&mut out, biff12::EXTERN_NAME, &extern_name);

    write_record(&mut out, biff12::SUPBOOK_END, &[]);

    // ExternSheet table with a single internal sheet span: ixti=0 -> Sheet1:Sheet3.
    let mut extern_sheet = Vec::<u8>::new();
    extern_sheet.extend_from_slice(&1u16.to_le_bytes()); // cxti
    extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // supbook_index=0 (internal)
    extern_sheet.extend_from_slice(&0u16.to_le_bytes()); // sheet_first=0 (Sheet1)
    extern_sheet.extend_from_slice(&2u16.to_le_bytes()); // sheet_last=2 (Sheet3)
    write_record(&mut out, biff12::EXTERN_SHEET, &extern_sheet);

    out
}

fn build_workbook_rels_xml() -> String {
    // `parse_relationships` only needs Id + Target. Type is optional.
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#
    .to_string();

    for idx in 1..=3 {
        xml.push_str(&format!(
            "  <Relationship Id=\"rId{idx}\" Target=\"worksheets/sheet{idx}.bin\"/>\n"
        ));
    }

    xml.push_str("</Relationships>\n");
    xml
}

fn write_temp_xlsb(bytes: &[u8]) -> NamedTempFile {
    let mut file = tempfile::Builder::new()
        .prefix("formula_xlsb_fixture_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    file.write_all(bytes).expect("write temp xlsb");
    file.flush().expect("flush temp xlsb");
    file
}

fn build_fixture_bytes() -> Vec<u8> {
    let workbook_bin = build_workbook_bin();
    let workbook_rels_xml = build_workbook_rels_xml();

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip_out
        .start_file("xl/workbook.bin", options.clone())
        .expect("start xl/workbook.bin");
    zip_out
        .write_all(&workbook_bin)
        .expect("write xl/workbook.bin");

    zip_out
        .start_file("xl/_rels/workbook.bin.rels", options)
        .expect("start xl/_rels/workbook.bin.rels");
    zip_out
        .write_all(workbook_rels_xml.as_bytes())
        .expect("write xl/_rels/workbook.bin.rels");

    zip_out.finish().expect("finish zip").into_inner()
}

#[test]
fn encoder_roundtrips_sheet_range_scoped_defined_name_via_namex() {
    let bytes = build_fixture_bytes();
    let tmp = write_temp_xlsb(&bytes);

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded_unquoted = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast("=Sheet1:Sheet3!MyName", ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context("=Sheet1:Sheet3!MyName", ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };
    assert_eq!(
        encoded_unquoted.rgce,
        vec![
            0x39, // PtgNameX (ref class)
            0x00, 0x00, // ixti=0 (Sheet1:Sheet3)
            0x01, 0x00, // nameIndex=1 ("MyName")
        ]
    );

    let encoded_quoted = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast("='Sheet1:Sheet3'!MyName", ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context("='Sheet1:Sheet3'!MyName", ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, ctx).expect("decode");
    assert_eq!(decoded, "'Sheet1:Sheet3'!MyName");
}
