use formula_xlsb::{OpenOptions, XlsbWorkbook};
use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use tempfile::tempdir;
use zip::ZipArchive;

fn insert_before_closing_tag(mut xml: String, closing_tag: &str, insert: &str) -> String {
    let idx = xml
        .rfind(closing_tag)
        .unwrap_or_else(|| panic!("missing closing tag {closing_tag}"));
    xml.insert_str(idx, insert);
    xml
}

fn build_fixture_with_calc_chain(base_bytes: &[u8]) -> Vec<u8> {
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

    parts.insert("xl/calcChain.bin".to_string(), b"dummy".to_vec());

    let content_types = String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 ct");
    let content_types = insert_before_closing_tag(
        content_types,
        "</Types>",
        "  <Override PartName=\"/xl/calcChain.bin\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n",
    );
    parts.insert(
        "[Content_Types].xml".to_string(),
        content_types.into_bytes(),
    );

    let workbook_rels =
        String::from_utf8(parts["xl/_rels/workbook.bin.rels"].clone()).expect("utf8 rels");
    let workbook_rels = insert_before_closing_tag(
        workbook_rels,
        "</Relationships>",
        "  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.bin\"/>\n",
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

fn read_record_id(buf: &[u8]) -> Option<(u32, usize)> {
    let mut v: u32 = 0;
    for i in 0..4 {
        let byte = *buf.get(i)?;
        v |= (byte as u32) << (8 * i);
        if byte & 0x80 == 0 {
            return Some((v, i + 1));
        }
    }
    None
}

fn read_record_len(buf: &[u8]) -> Option<(u32, usize)> {
    let mut v: u32 = 0;
    for i in 0..4 {
        let byte = *buf.get(i)?;
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Some((v, i + 1));
        }
    }
    None
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
}
