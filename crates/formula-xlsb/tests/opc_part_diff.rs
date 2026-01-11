use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn fixture_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb")
}

fn insert_before_closing_tag(mut xml: String, closing_tag: &str, insert: &str) -> String {
    let idx = xml
        .rfind(closing_tag)
        .unwrap_or_else(|| panic!("missing closing tag {closing_tag}"));
    xml.insert_str(idx, insert);
    xml
}

fn build_fixture_with_calc_chain_and_styles(base_bytes: &[u8]) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(base_bytes)).expect("open base xlsb zip");
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("read zip entry");
        if !file.is_file() {
            continue;
        }
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read part bytes");
        parts.insert(file.name().to_string(), buf);
    }

    // Add an arbitrary, non-empty styles.bin payload. `formula-xlsb` treats this as an opaque
    // preserved part; the tests ensure we never churn it accidentally.
    parts.insert("xl/styles.bin".to_string(), b"styles\0".to_vec());

    // Add a dummy calcChain part + references, so edited saves can validate the "remove calcChain"
    // behavior without checking in an additional binary fixture.
    parts.insert("xl/calcChain.bin".to_string(), b"dummy".to_vec());

    let content_types =
        String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
    let content_types = insert_before_closing_tag(
        content_types,
        "</Types>",
        "  <Override PartName=\"/xl/calcChain.bin\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n",
    );
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    let workbook_rels =
        String::from_utf8(parts["xl/_rels/workbook.bin.rels"].clone()).expect("utf8 workbook rels");
    let workbook_rels = insert_before_closing_tag(
        workbook_rels,
        "</Relationships>",
        "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.bin\"/>\n  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.bin\"/>\n",
    );
    parts.insert("xl/_rels/workbook.bin.rels".to_string(), workbook_rels.into_bytes());

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::default().compression_method(CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip_out.start_file(name, options).expect("start zip file");
        zip_out.write_all(&bytes).expect("write zip bytes");
    }

    zip_out.finish().expect("finish zip").into_inner()
}

fn format_report(report: &xlsx_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn save_as_is_lossless_at_opc_part_level() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("roundtrip.xlsb");
    wb.save_as(&out_path).expect("save_as");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got:\n{}",
        format_report(&report)
    );
}

#[test]
fn save_as_is_lossless_when_styles_and_calc_chain_parts_exist() {
    let base_path = fixture_path();
    let base_bytes = std::fs::read(&base_path).expect("read base fixture");
    let variant_bytes = build_fixture_with_calc_chain_and_styles(&base_bytes);

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("with_calc_chain.xlsb");
    let out_path = tmpdir.path().join("roundtrip.xlsb");
    std::fs::write(&input_path, variant_bytes).expect("write variant fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open variant xlsb");
    wb.save_as(&out_path).expect("save_as");

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got:\n{}",
        format_report(&report)
    );
}

#[test]
fn patch_writer_changes_only_target_sheet_part() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched.xlsb");
    wb.save_with_edits(&out_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet = patched.read_sheet(0).expect("read patched sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.bin"),
        "expected worksheet part to change, got:\n{}",
        format_report(&report)
    );

    let unexpected_missing: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part" && !is_calc_chain_part(&d.part))
        .map(|d| d.part.clone())
        .collect();
    assert!(
        unexpected_missing.is_empty(),
        "unexpected missing parts: {unexpected_missing:?}\n{}",
        format_report(&report)
    );

    let parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = parts
        .iter()
        .filter(|part| !is_allowed_patch_diff_part(part))
        .cloned()
        .collect();

    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{}",
        format_report(&report)
    );
}

#[test]
fn patch_writer_allows_only_expected_calc_chain_side_effects() {
    let base_path = fixture_path();
    let base_bytes = std::fs::read(&base_path).expect("read base fixture");
    let variant_bytes = build_fixture_with_calc_chain_and_styles(&base_bytes);

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("with_calc_chain.xlsb");
    let out_path = tmpdir.path().join("patched.xlsb");
    std::fs::write(&input_path, variant_bytes).expect("write variant fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open variant xlsb");
    wb.save_with_edits(&out_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");

    // Edits must change the sheet part.
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.bin"),
        "expected worksheet part to change, got:\n{}",
        format_report(&report)
    );

    // Edits should remove calcChain.bin; allow the necessary rels/CT rewrites.
    let allowed_parts = BTreeSet::from([
        "xl/worksheets/sheet1.bin".to_string(),
        "xl/calcChain.bin".to_string(),
        "[Content_Types].xml".to_string(),
        "xl/_rels/workbook.bin.rels".to_string(),
    ]);

    let missing_parts: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part")
        .map(|d| d.part.clone())
        .collect();
    assert_eq!(
        missing_parts,
        vec!["xl/calcChain.bin".to_string()],
        "expected only calcChain.bin to be missing; report:\n{}",
        format_report(&report)
    );

    for diff in report
        .differences
        .iter()
        .filter(|d| d.part == "[Content_Types].xml")
    {
        let mentions_calc_chain = diff.path.contains("calcChain")
            || diff
                .expected
                .as_deref()
                .map_or(false, |value| value.contains("calcChain"))
            || diff
                .actual
                .as_deref()
                .map_or(false, |value| value.contains("calcChain"));
        assert!(
            mentions_calc_chain,
            "unexpected diff in [Content_Types].xml:\n{diff}\nfull report:\n{}",
            format_report(&report)
        );
    }
    for diff in report
        .differences
        .iter()
        .filter(|d| d.part == "xl/_rels/workbook.bin.rels")
    {
        let mentions_calc_chain = diff.path.contains("rId4")
            || diff.path.contains("calcChain")
            || diff
                .expected
                .as_deref()
                .map_or(false, |value| value.contains("calcChain"))
            || diff
                .actual
                .as_deref()
                .map_or(false, |value| value.contains("calcChain"));
        assert!(
            mentions_calc_chain,
            "unexpected diff in xl/_rels/workbook.bin.rels:\n{diff}\nfull report:\n{}",
            format_report(&report)
        );
    }

    let diff_parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = diff_parts
        .difference(&allowed_parts)
        .cloned()
        .collect();
    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{}",
        format_report(&report)
    );

    // Ensure we did not accidentally drop unrelated parts.
    let out_zip = ZipArchive::new(std::fs::File::open(&out_path).expect("open patched zip"))
        .expect("read patched zip");
    assert!(
        out_zip
            .file_names()
            .any(|name| name == "xl/styles.bin"),
        "xl/styles.bin must be preserved in patched workbook"
    );
}

fn is_allowed_patch_diff_part(part: &str) -> bool {
    part == "xl/worksheets/sheet1.bin" || is_calc_chain_part(part)
}

fn is_calc_chain_part(part: &str) -> bool {
    part.starts_with("xl/calcChain.")
}
