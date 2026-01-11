use formula_xlsb::{biff12_varint, CellValue, OpenOptions, XlsbWorkbook};
use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use tempfile::tempdir;
use zip::ZipArchive;
use xlsx_diff::DiffReport;

fn insert_before_closing_tag(mut xml: String, closing_tag: &str, insert: &str) -> String {
    let idx = xml
        .rfind(closing_tag)
        .unwrap_or_else(|| panic!("missing closing tag {closing_tag}"));
    xml.insert_str(idx, insert);
    xml
}

fn build_fixture_with_calc_chain_custom(base_bytes: &[u8], calc_chain_part: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(base_bytes)).expect("open base zip");
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

    let calc_chain_part = calc_chain_part.trim_start_matches('/');
    let content_types_part_name = format!("/{calc_chain_part}");
    let workbook_rels_target = calc_chain_part
        .strip_prefix("xl/")
        .unwrap_or(calc_chain_part);

    parts.insert(calc_chain_part.to_string(), b"dummy".to_vec());

    let content_types = String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 ct");
    let content_types_insert = format!(
        "  <Override PartName=\"{content_types_part_name}\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n"
    );
    let content_types = insert_before_closing_tag(
        content_types,
        "</Types>",
        &content_types_insert,
    );
    parts.insert(
        "[Content_Types].xml".to_string(),
        content_types.into_bytes(),
    );

    let workbook_rels =
        String::from_utf8(parts["xl/_rels/workbook.bin.rels"].clone()).expect("utf8 rels");
    let workbook_rels_insert = format!(
        "  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"{workbook_rels_target}\"/>\n"
    );
    let workbook_rels = insert_before_closing_tag(
        workbook_rels,
        "</Relationships>",
        &workbook_rels_insert,
    );
    parts.insert(
        "xl/_rels/workbook.bin.rels".to_string(),
        workbook_rels.into_bytes(),
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip_out
            .start_file(name, options)
            .expect("write part header");
        zip_out.write_all(&bytes).expect("write part bytes");
    }

    zip_out.finish().expect("finish zip").into_inner()
}

fn build_fixture_with_calc_chain(base_bytes: &[u8]) -> Vec<u8> {
    build_fixture_with_calc_chain_custom(base_bytes, "xl/calcChain.bin")
}

fn read_record_id(buf: &[u8]) -> Option<(u32, usize)> {
    let mut cursor = Cursor::new(buf);
    let id = biff12_varint::read_record_id(&mut cursor).ok()??;
    Some((id, cursor.position() as usize))
}

fn read_record_len(buf: &[u8]) -> Option<(u32, usize)> {
    let mut cursor = Cursor::new(buf);
    let len = biff12_varint::read_record_len(&mut cursor).ok()??;
    Some((len, cursor.position() as usize))
}

fn tweak_first_float_cell(sheet_bytes: &[u8]) -> Vec<u8> {
    // This fixture's B1 cell is encoded as a BrtCellReal / FLOAT record. We keep the
    // record structure identical and just nudge the f64 value so the worksheet bytes
    // differ from the original.
    let mut out = sheet_bytes.to_vec();
    let mut offset = 0usize;

    while offset < out.len() {
        let (id, id_len) =
            read_record_id(&out[offset..]).unwrap_or_else(|| panic!("bad record id at {offset}"));
        offset += id_len;

        let (len, len_len) =
            read_record_len(&out[offset..]).unwrap_or_else(|| panic!("bad record len at {offset}"));
        offset += len_len;

        let data_start = offset;
        let data_end = offset
            .checked_add(len as usize)
            .unwrap_or_else(|| panic!("overflow reading record data at {offset}"));
        assert!(data_end <= out.len(), "record overruns buffer");

        if id == 0x0005 && len as usize >= 16 {
            // [col: u32][style: u32][value: f64]
            let col = u32::from_le_bytes(out[data_start..data_start + 4].try_into().unwrap());
            if col == 1 {
                let value_off = data_start + 8;
                let mut bytes = [0u8; 8];
                bytes.copy_from_slice(&out[value_off..value_off + 8]);
                let v = f64::from_le_bytes(bytes);
                let updated = v + 1.0;
                out[value_off..value_off + 8].copy_from_slice(&updated.to_le_bytes());
                return out;
            }
        }

        offset = data_end;
    }

    panic!("did not find FLOAT record to tweak");
}

fn format_report(report: &DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn save_as_preserves_calc_chain_when_unedited() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    let with_calc_chain = build_fixture_with_calc_chain(&base_bytes);

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("roundtrip.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");
    wb.save_as(&output_path).expect("save_as");

    let mut zip = ZipArchive::new(File::open(&output_path).expect("open output zip"))
        .expect("read output zip");

    zip.by_name("xl/calcChain.bin")
        .expect("calcChain.bin should be preserved");
}

#[test]
fn edited_save_removes_calc_chain_and_references() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    let with_calc_chain = build_fixture_with_calc_chain(&base_bytes);

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("edited.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");

    // Override the worksheet part with a tiny, structure-preserving edit.
    let mut zip_in =
        ZipArchive::new(File::open(&input_path).expect("open input zip")).expect("read input zip");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let mut sheet_bytes = Vec::new();
    zip_in
        .by_name(&sheet_part)
        .expect("read sheet part")
        .read_to_end(&mut sheet_bytes)
        .expect("read sheet bytes");
    let edited_sheet = tweak_first_float_cell(&sheet_bytes);

    let mut overrides = HashMap::new();
    overrides.insert(sheet_part, edited_sheet);

    wb.save_with_part_overrides(&output_path, &overrides)
        .expect("save with part overrides");

    let mut zip_out = ZipArchive::new(File::open(&output_path).expect("open output zip"))
        .expect("read output zip");

    assert!(
        zip_out.by_name("xl/calcChain.bin").is_err(),
        "calcChain.bin should be removed on edited save"
    );

    let mut ct_bytes = Vec::new();
    zip_out
        .by_name("[Content_Types].xml")
        .expect("read content types")
        .read_to_end(&mut ct_bytes)
        .expect("read ct bytes");
    let ct = String::from_utf8(ct_bytes).expect("utf8 ct");
    assert!(
        !ct.contains("calcChain"),
        "[Content_Types].xml should not reference calcChain after edit"
    );

    let mut rels_bytes = Vec::new();
    zip_out
        .by_name("xl/_rels/workbook.bin.rels")
        .expect("read workbook rels")
        .read_to_end(&mut rels_bytes)
        .expect("read rels bytes");
    let rels = String::from_utf8(rels_bytes).expect("utf8 rels");
    assert!(
        !rels.contains("calcChain"),
        "workbook.bin.rels should not reference calcChain after edit"
    );

    // Sanity-check that the resulting workbook is still readable.
    let reopened =
        XlsbWorkbook::open_with_options(&output_path, OpenOptions::default()).expect("reopen");
    let sheet = reopened.read_sheet(0).expect("read sheet after save");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 1))
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(43.5));
}

#[test]
fn edited_save_removes_calc_chain_with_weird_casing() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    let with_calc_chain = build_fixture_with_calc_chain_custom(&base_bytes, "xl/CalcChain.bin");

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("edited.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");

    // Override the worksheet part with a tiny, structure-preserving edit.
    let mut zip_in =
        ZipArchive::new(File::open(&input_path).expect("open input zip")).expect("read input zip");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let mut sheet_bytes = Vec::new();
    zip_in
        .by_name(&sheet_part)
        .expect("read sheet part")
        .read_to_end(&mut sheet_bytes)
        .expect("read sheet bytes");
    let edited_sheet = tweak_first_float_cell(&sheet_bytes);

    let mut overrides = HashMap::new();
    overrides.insert(sheet_part, edited_sheet);

    wb.save_with_part_overrides(&output_path, &overrides)
        .expect("save with part overrides");

    let mut zip_out = ZipArchive::new(File::open(&output_path).expect("open output zip"))
        .expect("read output zip");

    let has_calc_chain = zip_out
        .file_names()
        .any(|name| name.to_ascii_lowercase() == "xl/calcchain.bin");
    assert!(
        !has_calc_chain,
        "calcChain part should be removed regardless of casing"
    );

    let mut ct_bytes = Vec::new();
    zip_out
        .by_name("[Content_Types].xml")
        .expect("read content types")
        .read_to_end(&mut ct_bytes)
        .expect("read ct bytes");
    let ct = String::from_utf8(ct_bytes).expect("utf8 ct");
    assert!(
        !ct.to_ascii_lowercase().contains("calcchain"),
        "[Content_Types].xml should not reference calcChain after edit (case-insensitive)"
    );

    let mut rels_bytes = Vec::new();
    zip_out
        .by_name("xl/_rels/workbook.bin.rels")
        .expect("read workbook rels")
        .read_to_end(&mut rels_bytes)
        .expect("read rels bytes");
    let rels = String::from_utf8(rels_bytes).expect("utf8 rels");
    assert!(
        !rels.to_ascii_lowercase().contains("calcchain"),
        "workbook.bin.rels should not reference calcChain after edit (case-insensitive)"
    );
}

#[test]
fn edited_save_ignores_calc_chain_override_key_case() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    // Package contains `xl/CalcChain.bin` (upper-case C).
    let with_calc_chain = build_fixture_with_calc_chain_custom(&base_bytes, "xl/CalcChain.bin");

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("edited.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");

    // Apply a worksheet edit (to trigger calcChain invalidation) and also provide an
    // override for calcChain using different casing than the ZIP entry. The writer should
    // ignore the calcChain override rather than error.
    let mut zip_in =
        ZipArchive::new(File::open(&input_path).expect("open input zip")).expect("read input zip");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let mut sheet_bytes = Vec::new();
    zip_in
        .by_name(&sheet_part)
        .expect("read sheet part")
        .read_to_end(&mut sheet_bytes)
        .expect("read sheet bytes");
    let edited_sheet = tweak_first_float_cell(&sheet_bytes);

    let mut overrides = HashMap::new();
    overrides.insert(sheet_part, edited_sheet);
    overrides.insert("xl/calcChain.bin".to_string(), b"ignored".to_vec());

    wb.save_with_part_overrides(&output_path, &overrides)
        .expect("save with part overrides");

    let zip_out = ZipArchive::new(File::open(&output_path).expect("open output zip"))
        .expect("read output zip");
    let has_calc_chain = zip_out
        .file_names()
        .any(|name| name.to_ascii_lowercase() == "xl/calcchain.bin");
    assert!(
        !has_calc_chain,
        "calcChain should be removed even if caller provides an override with different casing"
    );
}

#[test]
fn edited_save_changes_only_expected_parts() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    let with_calc_chain = build_fixture_with_calc_chain(&base_bytes);

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("edited.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");

    // Override the worksheet part with a tiny, structure-preserving edit.
    let mut zip_in =
        ZipArchive::new(File::open(&input_path).expect("open input zip")).expect("read input zip");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let mut sheet_bytes = Vec::new();
    zip_in
        .by_name(&sheet_part)
        .expect("read sheet part")
        .read_to_end(&mut sheet_bytes)
        .expect("read sheet bytes");
    let edited_sheet = tweak_first_float_cell(&sheet_bytes);

    let mut overrides = HashMap::new();
    overrides.insert(sheet_part.clone(), edited_sheet);

    wb.save_with_part_overrides(&output_path, &overrides)
        .expect("save with part overrides");

    let report = xlsx_diff::diff_workbooks(&input_path, &output_path).expect("diff workbooks");

    let expected_parts: std::collections::BTreeSet<String> = [
        sheet_part,
        "xl/calcChain.bin".to_string(),
        "[Content_Types].xml".to_string(),
        "xl/_rels/workbook.bin.rels".to_string(),
    ]
    .into_iter()
    .collect();

    let actual_parts: std::collections::BTreeSet<String> =
        report.differences.iter().map(|d| d.part.clone()).collect();

    let unexpected: Vec<_> = actual_parts
        .difference(&expected_parts)
        .cloned()
        .collect();

    assert!(
        unexpected.is_empty(),
        "unexpected part diffs: {unexpected:?}\n{}",
        format_report(&report)
    );
}

#[test]
fn noop_worksheet_override_preserves_calc_chain() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let base_bytes = std::fs::read(fixture_path).expect("read fixture");
    let with_calc_chain = build_fixture_with_calc_chain(&base_bytes);

    let dir = tempdir().expect("tempdir");
    let input_path = dir.path().join("with_calc_chain.xlsb");
    let output_path = dir.path().join("noop_override.xlsb");
    std::fs::write(&input_path, with_calc_chain).expect("write input");

    let wb = XlsbWorkbook::open_with_options(&input_path, OpenOptions::default()).expect("open");

    // Provide an override for the worksheet part that is byte-identical to the original.
    let mut zip_in =
        ZipArchive::new(File::open(&input_path).expect("open input zip")).expect("read input zip");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let mut sheet_bytes = Vec::new();
    zip_in
        .by_name(&sheet_part)
        .expect("read sheet part")
        .read_to_end(&mut sheet_bytes)
        .expect("read sheet bytes");

    let mut overrides = HashMap::new();
    overrides.insert(sheet_part, sheet_bytes);

    wb.save_with_part_overrides(&output_path, &overrides)
        .expect("save with part overrides");

    let mut zip_out = ZipArchive::new(File::open(&output_path).expect("open output zip"))
        .expect("read output zip");

    zip_out
        .by_name("xl/calcChain.bin")
        .expect("calcChain.bin should be preserved");

    let mut ct_bytes = Vec::new();
    zip_out
        .by_name("[Content_Types].xml")
        .expect("read content types")
        .read_to_end(&mut ct_bytes)
        .expect("read ct bytes");
    let ct = String::from_utf8(ct_bytes).expect("utf8 ct");
    assert!(
        ct.contains("calcChain"),
        "[Content_Types].xml should still reference calcChain when worksheet override is a no-op"
    );

    let mut rels_bytes = Vec::new();
    zip_out
        .by_name("xl/_rels/workbook.bin.rels")
        .expect("read workbook rels")
        .read_to_end(&mut rels_bytes)
        .expect("read rels bytes");
    let rels = String::from_utf8(rels_bytes).expect("utf8 rels");
    assert!(
        rels.contains("calcChain"),
        "workbook.bin.rels should still reference calcChain when worksheet override is a no-op"
    );
}
