use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};
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

fn zip_has_part(path: &Path, part: &str) -> bool {
    let file = std::fs::File::open(path).expect("open zip");
    let zip = ZipArchive::new(file).expect("read zip");
    let has = zip.file_names().any(|name| name == part);
    has
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

fn build_sheet_bin_single_float(row: u32, col: u32, value: f64) -> Vec<u8> {
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

fn build_content_types_xml(sheet_count: usize) -> String {
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
  <Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.main"/>
"#
    .to_string();

    for idx in 1..=sheet_count {
        xml.push_str(&format!(
            "  <Override PartName=\"/xl/worksheets/sheet{idx}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>\n"
        ));
    }

    xml.push_str(
        "  <Override PartName=\"/xl/calcChain.bin\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n",
    );
    xml.push_str("</Types>\n");
    xml
}

fn build_root_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
</Relationships>
"#
    .to_string()
}

fn build_workbook_rels_xml(sheet_count: usize) -> String {
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#
    .to_string();

    for idx in 1..=sheet_count {
        xml.push_str(&format!(
            "  <Relationship Id=\"rId{idx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{idx}.bin\"/>\n"
        ));
    }

    let calc_chain_rid = sheet_count + 1;
    xml.push_str(&format!(
        "  <Relationship Id=\"rId{calc_chain_rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.bin\"/>\n"
    ));

    xml.push_str("</Relationships>\n");
    xml
}

fn build_three_sheet_fixture_bytes() -> Vec<u8> {
    let workbook_bin = build_workbook_bin(&["Sheet1", "Sheet2", "Sheet3"]);
    let sheet1_bin = build_sheet_bin_single_float(0, 1, 10.0);
    let sheet2_bin = build_sheet_bin_single_float(0, 1, 20.0);
    let sheet3_bin = build_sheet_bin_single_float(0, 1, 30.0);

    let content_types_xml = build_content_types_xml(3);
    let rels_xml = build_root_rels_xml();
    let workbook_rels_xml = build_workbook_rels_xml(3);

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip_out
        .start_file("[Content_Types].xml", options.clone())
        .expect("start [Content_Types].xml");
    zip_out
        .write_all(content_types_xml.as_bytes())
        .expect("write [Content_Types].xml");

    zip_out
        .start_file("_rels/.rels", options.clone())
        .expect("start _rels/.rels");
    zip_out
        .write_all(rels_xml.as_bytes())
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
        .write_all(workbook_rels_xml.as_bytes())
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
        .start_file("xl/worksheets/sheet3.bin", options.clone())
        .expect("start xl/worksheets/sheet3.bin");
    zip_out
        .write_all(&sheet3_bin)
        .expect("write xl/worksheets/sheet3.bin");

    zip_out
        .start_file("xl/calcChain.bin", options.clone())
        .expect("start xl/calcChain.bin");
    zip_out.write_all(b"dummy").expect("write calcChain");

    zip_out.finish().expect("finish zip").into_inner()
}

#[test]
fn save_with_cell_edits_multi_changes_only_edited_sheet_parts() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("multi_sheet.xlsb");
    let out_path = tmpdir.path().join("patched.xlsb");
    std::fs::write(&input_path, build_three_sheet_fixture_bytes()).expect("write fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open fixture");
    assert_eq!(wb.sheet_metas().len(), 3);
    assert!(zip_has_part(&input_path, "xl/calcChain.bin"));

    let mut edits_by_sheet: BTreeMap<usize, Vec<CellEdit>> = BTreeMap::new();
    edits_by_sheet.insert(
        0,
        vec![CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(123.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    );
    edits_by_sheet.insert(
        2,
        vec![CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(456.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    );

    wb.save_with_cell_edits_multi(&out_path, &edits_by_sheet)
        .expect("save_with_cell_edits_multi");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet1 = patched.read_sheet(0).expect("read patched sheet1");
    let b1 = sheet1
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("Sheet1!B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));

    let sheet2 = patched.read_sheet(1).expect("read patched sheet2");
    let b1 = sheet2
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("Sheet2!B1 exists");
    assert_eq!(b1.value, CellValue::Number(20.0));

    let sheet3 = patched.read_sheet(2).expect("read patched sheet3");
    let b1 = sheet3
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("Sheet3!B1 exists");
    assert_eq!(b1.value, CellValue::Number(456.0));

    // Any sheet edit must invalidate calcChain (remove the part and clean up its plumbing).
    assert!(
        !zip_has_part(&out_path, "xl/calcChain.bin"),
        "patched workbook should not retain xl/calcChain.bin"
    );

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");
    assert_no_unexpected_extra_parts(&report);
    let report_text = format_report(&report);

    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.bin"),
        "expected worksheet part sheet1.bin to change, got:\n{report_text}"
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet3.bin"),
        "expected worksheet part sheet3.bin to change, got:\n{report_text}"
    );
    assert!(
        !report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet2.bin"),
        "did not expect unedited worksheet part sheet2.bin to change, got:\n{report_text}"
    );

    let missing_parts: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part")
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        missing_parts,
        vec!["xl/calcChain.bin".to_string()],
        "expected only calcChain.bin to be missing; got {missing_parts:?}\n{report_text}"
    );

    let allowed_parts = BTreeSet::from([
        "xl/worksheets/sheet1.bin".to_string(),
        "xl/worksheets/sheet3.bin".to_string(),
        "xl/calcChain.bin".to_string(),
        "[Content_Types].xml".to_string(),
        "xl/_rels/workbook.bin.rels".to_string(),
    ]);
    let diff_parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = diff_parts.difference(&allowed_parts).cloned().collect();

    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{report_text}"
    );
}
