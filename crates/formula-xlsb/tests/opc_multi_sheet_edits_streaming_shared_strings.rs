use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

// Record IDs copied from MS-XLSB / our parser constants. (Not re-exported from the crate.)
mod biff12 {
    pub const SHEET: u32 = 0x009C;

    pub const WORKSHEET: u32 = 0x0081;
    pub const WORKSHEET_END: u32 = 0x0082;
    pub const SHEETDATA: u32 = 0x0091;
    pub const SHEETDATA_END: u32 = 0x0092;
    pub const DIMENSION: u32 = 0x0094;

    pub const ROW: u32 = 0x0000;
    pub const FLOAT: u32 = 0x0005;
    pub const STRING: u32 = 0x0007;

    pub const SST: u32 = 0x009F;
    pub const SST_END: u32 = 0x00A0;
    pub const SI: u32 = 0x0013;
}

fn format_report(report: &xlsx_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

fn assert_no_unexpected_extra_parts(report: &xlsx_diff::DiffReport) {
    let extra_parts: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "extra_part")
        .map(|d| d.part.clone())
        .collect();
    assert!(
        extra_parts.is_empty(),
        "unexpected extra parts in diff: {extra_parts:?}\n{}",
        format_report(report)
    );
}

fn read_zip_part(path: &Path, part_path: &str) -> Vec<u8> {
    let file = std::fs::File::open(path).expect("open xlsb");
    let mut zip = ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => return None,
        };
        let mut payload = vec![0u8; len];
        cursor.read_exact(&mut payload).ok()?;

        match id {
            0x0091 => in_sheet_data = true,  // BrtSheetData
            0x0092 => in_sheet_data = false, // BrtSheetDataEnd
            0x0000 if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() < 8 {
                    continue;
                }
                let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                if current_row == target_row && col == target_col {
                    return Some((id, payload));
                }
            }
            _ => {}
        }
    }
    None
}

struct SharedStringsStats {
    total_count: Option<u32>,
    unique_count: Option<u32>,
    si_count: u32,
    strings: Vec<String>,
}

fn read_shared_strings_stats(shared_strings_bin: &[u8]) -> SharedStringsStats {
    let mut cursor = Cursor::new(shared_strings_bin);
    let mut total_count = None;
    let mut unique_count = None;
    let mut si_count = 0u32;
    let mut strings = Vec::new();

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        let mut payload = vec![0u8; len];
        cursor
            .read_exact(&mut payload)
            .expect("read record payload");

        match id {
            0x009F if payload.len() >= 8 => {
                // BrtSST: [totalCount][uniqueCount]
                total_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique_count = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
            }
            0x0013 if payload.len() >= 5 => {
                // BrtSI payload: [flags: u8][text: XLWideString]
                si_count = si_count.saturating_add(1);
                let flags = payload[0];
                let cch = u32::from_le_bytes(payload[1..5].try_into().unwrap()) as usize;
                let byte_len = cch.saturating_mul(2);
                let raw = payload.get(5..5 + byte_len).unwrap_or(&[]);
                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                if flags == 0 {
                    strings.push(String::from_utf16_lossy(&units));
                }
            }
            0x00A0 => break, // BrtSSTEnd
            _ => {}
        }
    }

    SharedStringsStats {
        total_count,
        unique_count,
        si_count,
        strings,
    }
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

fn build_workbook_bin(sheet_names: &[&str]) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    for (idx, sheet_name) in sheet_names.iter().enumerate() {
        let mut sheet = Vec::<u8>::new();
        sheet.extend_from_slice(&0u32.to_le_bytes()); // flags/state (unused by our parser)
        sheet.extend_from_slice(&(idx as u32 + 1).to_le_bytes()); // sheet id

        let rid = format!("rId{}", idx + 1);
        write_utf16_string(&mut sheet, &rid);
        write_utf16_string(&mut sheet, sheet_name);

        write_record(&mut out, biff12::SHEET, &sheet);
    }
    out
}

fn build_sheet_bin_single_cell_float(row: u32, col: u32, value: f64) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    write_record(&mut out, biff12::WORKSHEET, &[]);

    // BrtWsDim: [r1: u32][r2: u32][c1: u32][c2: u32]
    let mut dim = Vec::<u8>::new();
    dim.extend_from_slice(&row.to_le_bytes());
    dim.extend_from_slice(&row.to_le_bytes());
    dim.extend_from_slice(&col.to_le_bytes());
    dim.extend_from_slice(&col.to_le_bytes());
    write_record(&mut out, biff12::DIMENSION, &dim);

    write_record(&mut out, biff12::SHEETDATA, &[]);
    write_record(&mut out, biff12::ROW, &row.to_le_bytes());

    // BrtCellReal: [col: u32][style: u32][value: f64]
    let mut cell = Vec::<u8>::new();
    cell.extend_from_slice(&col.to_le_bytes());
    cell.extend_from_slice(&0u32.to_le_bytes()); // style
    cell.extend_from_slice(&value.to_le_bytes());
    write_record(&mut out, biff12::FLOAT, &cell);

    write_record(&mut out, biff12::SHEETDATA_END, &[]);
    write_record(&mut out, biff12::WORKSHEET_END, &[]);
    out
}

fn build_sheet_bin_single_cell_sst(row: u32, col: u32, isst: u32) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    write_record(&mut out, biff12::WORKSHEET, &[]);

    // BrtWsDim: [r1: u32][r2: u32][c1: u32][c2: u32]
    let mut dim = Vec::<u8>::new();
    dim.extend_from_slice(&row.to_le_bytes());
    dim.extend_from_slice(&row.to_le_bytes());
    dim.extend_from_slice(&col.to_le_bytes());
    dim.extend_from_slice(&col.to_le_bytes());
    write_record(&mut out, biff12::DIMENSION, &dim);

    write_record(&mut out, biff12::SHEETDATA, &[]);
    write_record(&mut out, biff12::ROW, &row.to_le_bytes());

    // BrtCellIsst: [col: u32][style: u32][isst: u32]
    let mut cell = Vec::<u8>::new();
    cell.extend_from_slice(&col.to_le_bytes());
    cell.extend_from_slice(&0u32.to_le_bytes()); // style
    cell.extend_from_slice(&isst.to_le_bytes());
    write_record(&mut out, biff12::STRING, &cell);

    write_record(&mut out, biff12::SHEETDATA_END, &[]);
    write_record(&mut out, biff12::WORKSHEET_END, &[]);
    out
}

fn build_shared_strings_bin(strings: &[&str], total_count: u32) -> Vec<u8> {
    let mut out = Vec::<u8>::new();

    let unique_count = strings.len() as u32;
    let mut sst = Vec::<u8>::new();
    sst.extend_from_slice(&total_count.to_le_bytes());
    sst.extend_from_slice(&unique_count.to_le_bytes());
    write_record(&mut out, biff12::SST, &sst);

    for s in strings {
        let mut si = Vec::<u8>::new();
        si.push(0u8); // flags (plain)
        write_utf16_string(&mut si, s);
        write_record(&mut out, biff12::SI, &si);
    }

    write_record(&mut out, biff12::SST_END, &[]);
    out
}

fn build_content_types_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
  <Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
  <Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet"/>
  <Override PartName="/xl/worksheets/sheet2.bin" ContentType="application/vnd.ms-excel.worksheet"/>
  <Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>
</Types>
"#
    .to_string()
}

fn build_root_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>
"#
    .to_string()
}

fn build_workbook_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.bin"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>
</Relationships>
"#
    .to_string()
}

fn build_fixture_bytes() -> Vec<u8> {
    let workbook_bin = build_workbook_bin(&["Sheet1", "Sheet2"]);

    // Sheet1 has a numeric A1; Sheet2 has a shared-string A1 ("Hello").
    let sheet1_bin = build_sheet_bin_single_cell_float(0, 0, 42.0);
    let sheet2_bin = build_sheet_bin_single_cell_sst(0, 0, 0);

    // sharedStrings contains two items, but only one cell references the table initially.
    let shared_strings_bin = build_shared_strings_bin(&["Hello", "World"], 1);

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip_out
        .start_file("[Content_Types].xml", options.clone())
        .expect("start [Content_Types].xml");
    zip_out
        .write_all(build_content_types_xml().as_bytes())
        .expect("write [Content_Types].xml");

    zip_out
        .start_file("_rels/.rels", options.clone())
        .expect("start _rels/.rels");
    zip_out
        .write_all(build_root_rels_xml().as_bytes())
        .expect("write _rels/.rels");

    zip_out
        .start_file("xl/workbook.bin", options.clone())
        .expect("start xl/workbook.bin");
    zip_out
        .write_all(&workbook_bin)
        .expect("write xl/workbook.bin");

    zip_out
        .start_file("xl/_rels/workbook.bin.rels", options.clone())
        .expect("start xl/_rels/workbook.bin.rels");
    zip_out
        .write_all(build_workbook_rels_xml().as_bytes())
        .expect("write xl/_rels/workbook.bin.rels");

    zip_out
        .start_file("xl/worksheets/sheet1.bin", options.clone())
        .expect("start xl/worksheets/sheet1.bin");
    zip_out
        .write_all(&sheet1_bin)
        .expect("write xl/worksheets/sheet1.bin");

    zip_out
        .start_file("xl/worksheets/sheet2.bin", options.clone())
        .expect("start xl/worksheets/sheet2.bin");
    zip_out
        .write_all(&sheet2_bin)
        .expect("write xl/worksheets/sheet2.bin");

    zip_out
        .start_file("xl/sharedStrings.bin", options.clone())
        .expect("start xl/sharedStrings.bin");
    zip_out
        .write_all(&shared_strings_bin)
        .expect("write xl/sharedStrings.bin");

    zip_out.finish().expect("finish zip").into_inner()
}

#[test]
fn save_with_cell_edits_streaming_multi_shared_strings_updates_sheets_and_sst() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let out_path = tmpdir.path().join("out.xlsb");
    std::fs::write(&input_path, build_fixture_bytes()).expect("write fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open fixture");
    assert_eq!(wb.sheet_metas().len(), 2);

    let mut edits_by_sheet: BTreeMap<usize, Vec<CellEdit>> = BTreeMap::new();
    edits_by_sheet.insert(
        0,
        vec![CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    );
    edits_by_sheet.insert(
        1,
        vec![CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    );

    wb.save_with_cell_edits_streaming_multi_shared_strings(&out_path, &edits_by_sheet)
        .expect("save_with_cell_edits_streaming_multi_shared_strings");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet1 = patched.read_sheet(0).expect("read sheet1");
    let a1 = sheet1
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("Sheet1!A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hello".to_string()));

    let sheet2 = patched.read_sheet(1).expect("read sheet2");
    let a1 = sheet2
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("Sheet2!A1 exists");
    assert_eq!(a1.value, CellValue::Text("New".to_string()));

    let sheet1_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet1_bin, 0, 0).expect("find Sheet1!A1 record");
    assert_eq!(id, 0x0007, "expected Sheet1!A1 to be BrtCellIsst");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected Sheet1!A1 to reference isst=0 ('Hello')"
    );

    let sheet2_bin = read_zip_part(&out_path, "xl/worksheets/sheet2.bin");
    let (id, payload) = find_cell_record(&sheet2_bin, 0, 0).expect("find Sheet2!A1 record");
    assert_eq!(id, 0x0007, "expected Sheet2!A1 to be BrtCellIsst");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        2,
        "expected Sheet2!A1 to reference appended isst=2 ('New')"
    );

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(2));
    assert_eq!(stats.unique_count, Some(3));
    assert_eq!(stats.si_count, 3);
    assert!(
        stats.strings.contains(&"New".to_string()),
        "expected sharedStrings.bin to contain 'New', got {:?}",
        stats.strings
    );

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");
    assert_no_unexpected_extra_parts(&report);
    let report_text = format_report(&report);

    let allowed_parts = BTreeSet::from([
        "xl/worksheets/sheet1.bin".to_string(),
        "xl/worksheets/sheet2.bin".to_string(),
        "xl/sharedStrings.bin".to_string(),
    ]);
    let diff_parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = diff_parts.difference(&allowed_parts).cloned().collect();
    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{report_text}"
    );
}
