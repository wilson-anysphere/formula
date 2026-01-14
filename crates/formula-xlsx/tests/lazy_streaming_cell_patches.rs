use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, RecalcPolicy, WorkbookCellPatches, XlsxLazyPackage};
use zip::{write::FileOptions, ZipArchive, ZipWriter};

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = archive.by_name(name).unwrap();
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

/// Return the raw local file record bytes for `name` (local header + compressed data + optional data descriptor).
///
/// This is useful for asserting that an entry was preserved by `ZipWriter::raw_copy_file`.
fn local_file_record_bytes(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let file = archive.by_name(name).unwrap();
    let data_start = file.data_start() as usize;
    let compressed_size = file.compressed_size() as usize;

    // Find the local file header signature (`PK\x03\x04`) just before the data start.
    let sig = [0x50, 0x4b, 0x03, 0x04];
    let search_start = data_start.saturating_sub(4096);
    let header_start = zip_bytes[search_start..data_start]
        .windows(sig.len())
        .rposition(|w| w == sig)
        .map(|idx| search_start + idx)
        .expect("local header signature should be found before data_start");

    // Parse the general purpose bit flag to detect data descriptors.
    let gp_flags = u16::from_le_bytes([
        zip_bytes[header_start + 6],
        zip_bytes[header_start + 7],
    ]);
    let has_data_descriptor = (gp_flags & 0x0008) != 0;

    let mut record_end = data_start + compressed_size;
    if has_data_descriptor {
        // Data descriptor can be 12 bytes (no signature) or 16 bytes (with signature).
        let dd_sig = [0x50, 0x4b, 0x07, 0x08];
        if zip_bytes
            .get(record_end..record_end + 4)
            .is_some_and(|b| b == dd_sig)
        {
            record_end += 16;
        } else {
            record_end += 12;
        }
    }

    zip_bytes[header_start..record_end].to_vec()
}

#[test]
fn lazy_package_patch_cells_raw_copies_unmodified_entries() -> Result<(), Box<dyn std::error::Error>>
{
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/octet-stream"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // A large binary part that should be preserved byte-for-byte via `raw_copy_file`.
    let big_bin: Vec<u8> = (0..2_000_000u32)
        .map(|i| ((i.wrapping_mul(31) ^ (i >> 3)) % 251) as u8)
        .collect();

    let input_bytes = build_package(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", worksheet_xml),
        ("xl/media/big.bin", big_bin.as_slice()),
    ]);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value(CellValue::Number(42.0)),
    );

    let pkg = XlsxLazyPackage::from_bytes(&input_bytes)?;
    let mut out = Cursor::new(Vec::new());
    pkg.patch_cells_to_writer(&mut out, &patches, RecalcPolicy::default(), None)?;
    let out_bytes = out.into_inner();

    // Ensure the binary entry is preserved byte-for-byte (raw ZIP record, not just decompressed bytes).
    assert_eq!(
        local_file_record_bytes(&input_bytes, "xl/media/big.bin"),
        local_file_record_bytes(&out_bytes, "xl/media/big.bin"),
        "expected xl/media/big.bin local file record to be preserved via raw_copy_file"
    );
    assert_eq!(
        read_zip_part(&input_bytes, "xl/media/big.bin"),
        read_zip_part(&out_bytes, "xl/media/big.bin"),
        "expected xl/media/big.bin payload bytes to be preserved"
    );

    // Ensure the worksheet cell was patched.
    let sheet_xml = String::from_utf8(read_zip_part(&out_bytes, "xl/worksheets/sheet1.xml"))?;
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell to exist");
    let v_text = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v_text.trim(), "42");

    Ok(())
}
