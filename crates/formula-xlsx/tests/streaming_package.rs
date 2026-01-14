use std::io::{Cursor, Read, Seek, SeekFrom, Write};

use formula_xlsx::{PartOverride, StreamingXlsxPackage, WorkbookKind};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, DateTime, ZipArchive, ZipWriter};

fn build_zip(entries: &[(&str, CompressionMethod, &[u8])]) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut cursor);
        for (name, method, bytes) in entries {
            let options = FileOptions::<()>::default().compression_method(*method);
            zip.start_file(name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    cursor.into_inner()
}

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = zip.by_name(name).unwrap();
    let mut out = Vec::new();
    file.read_to_end(&mut out).unwrap();
    out
}

fn zip_part_last_modified(zip_bytes: &[u8], name: &str) -> DateTime {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let ts = zip
        .by_name(name)
        .unwrap()
        .last_modified()
        .expect("zip entry missing last_modified time");
    ts
}

fn read_zip_part_compressed_bytes(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let file = zip.by_name(name).unwrap();
    let start = file.data_start();
    let len = file.compressed_size();
    drop(file);

    let mut reader = zip.into_inner();
    reader.seek(SeekFrom::Start(start)).unwrap();
    let mut out = vec![0u8; len as usize];
    reader.read_exact(&mut out).unwrap();
    out
}

fn content_types_override_map(xml: &str) -> std::collections::BTreeMap<String, String> {
    let doc = Document::parse(xml).expect("parse [Content_Types].xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| {
            let part = n.attribute("PartName")?.to_string();
            let ct = n.attribute("ContentType")?.to_string();
            Some((part, ct))
        })
        .collect()
}

fn streaming_pkg_fixture_with_prefixed_content_types() -> (Vec<u8>, String) {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <ct:Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <ct:Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</ct:Types>"#;

    let input = build_zip(&[
        (
            "[Content_Types].xml",
            CompressionMethod::Deflated,
            content_types.as_bytes(),
        ),
        ("xl/workbook.xml", CompressionMethod::Deflated, b"<workbook/>"),
        ("xl/styles.xml", CompressionMethod::Deflated, b"<styleSheet/>"),
        (
            "xl/worksheets/sheet1.xml",
            CompressionMethod::Deflated,
            b"<worksheet/>",
        ),
    ]);

    (input, content_types.to_string())
}

#[test]
fn streaming_package_write_to_raw_copies_large_part() {
    // Highly compressible payload so that "raw-copy" vs "store uncompressed" is obvious.
    let big = vec![0u8; 5 * 1024 * 1024];

    // Also set a non-default timestamp on the entry so we can detect whether it was
    // raw-copied (timestamp preserved) vs rewritten (timestamp would change).
    let big_ts = DateTime::from_date_and_time(2001, 2, 3, 4, 5, 6).unwrap();
    let input = {
        let mut cursor = Cursor::new(Vec::new());
        {
            let mut zip = ZipWriter::new(&mut cursor);
            let options = FileOptions::<()>::default()
                .compression_method(CompressionMethod::Deflated)
                .last_modified_time(big_ts);
            zip.start_file("xl/big.bin", options).unwrap();
            zip.write_all(&big).unwrap();

            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
            zip.start_file("xl/other.txt", options).unwrap();
            zip.write_all(b"hello world").unwrap();

            zip.finish().unwrap();
        }
        cursor.into_inner()
    };

    let input_compressed = read_zip_part_compressed_bytes(&input, "xl/big.bin");
    let input_ts = zip_part_last_modified(&input, "xl/big.bin");

    let pkg = StreamingXlsxPackage::from_reader(Cursor::new(input.clone())).unwrap();
    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    // Uncompressed bytes must match.
    let output_big = read_zip_part(&output, "xl/big.bin");
    assert_eq!(output_big, big);

    // The compressed bytes should be identical when the entry is raw-copied.
    let output_compressed = read_zip_part_compressed_bytes(&output, "xl/big.bin");
    assert_eq!(output_compressed, input_compressed);

    // Raw-copy should also preserve entry metadata like timestamps.
    let output_ts = zip_part_last_modified(&output, "xl/big.bin");
    assert_eq!(output_ts, input_ts);

    // And the overall file size should not balloon toward the uncompressed payload size.
    assert!(
        output.len() < 200_000,
        "output ZIP unexpectedly large: {} bytes",
        output.len()
    );
}

#[test]
fn streaming_package_set_part_and_remove_part() {
    let input = build_zip(&[
        ("xl/workbook.xml", CompressionMethod::Deflated, b"old"),
        ("xl/to_remove.bin", CompressionMethod::Deflated, b"bye"),
    ]);

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();
    pkg.set_part("xl/workbook.xml", b"new".to_vec());
    pkg.remove_part("xl/to_remove.bin");

    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    assert_eq!(read_zip_part(&output, "xl/workbook.xml"), b"new");

    let mut zip = ZipArchive::new(Cursor::new(output)).unwrap();
    assert!(zip.by_name("xl/to_remove.bin").is_err());
}

#[test]
fn streaming_package_normalizes_backslashes_and_leading_slash() {
    let input = build_zip(&[
        // Non-canonical ZIP entry name (`\\` separator) seen in some broken producers.
        ("xl\\workbook.xml", CompressionMethod::Deflated, b"old"),
        // Also exercise leading `/` mismatch.
        ("/xl/keep.txt", CompressionMethod::Deflated, b"keep"),
    ]);

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();

    assert_eq!(
        pkg.read_part("xl/workbook.xml").unwrap().as_deref(),
        Some(b"old".as_slice())
    );
    // Canonical part names should be surfaced through part_names().
    let names: Vec<String> = pkg.part_names().collect();
    assert!(names.iter().any(|n| n == "xl/workbook.xml"));
    assert!(names.iter().any(|n| n == "xl/keep.txt"));

    pkg.set_part("/xl/workbook.xml", b"new".to_vec());
    pkg.remove_part("xl/keep.txt");

    let mut out = Cursor::new(Vec::new());
    pkg.write_to(&mut out).unwrap();
    let output = out.into_inner();

    // Replaced entry should still be found under its original ZIP name.
    assert_eq!(read_zip_part(&output, "xl\\workbook.xml"), b"new");

    // Removed entry should be absent (regardless of leading `/`).
    let mut zip = ZipArchive::new(Cursor::new(output)).unwrap();
    assert!(zip.by_name("/xl/keep.txt").is_err());
    assert!(zip.by_name("xl/keep.txt").is_err());
}

#[test]
fn streaming_package_enforce_workbook_kind_template_updates_only_workbook_override() {
    let (input, content_types) = streaming_pkg_fixture_with_prefixed_content_types();

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();
    pkg.enforce_workbook_kind(WorkbookKind::Template).unwrap();

    let updated_bytes = pkg.read_part("[Content_Types].xml").unwrap().unwrap();
    let updated = std::str::from_utf8(&updated_bytes).unwrap();

    // Ensure we recorded the rewrite as a `Replace` override.
    assert!(matches!(
        pkg.part_overrides().get("[Content_Types].xml"),
        Some(PartOverride::Replace(_))
    ));

    let original_overrides = content_types_override_map(&content_types);
    let mut expected_overrides = original_overrides.clone();
    expected_overrides.insert(
        "/xl/workbook.xml".to_string(),
        WorkbookKind::Template.workbook_content_type().to_string(),
    );

    let actual_overrides = content_types_override_map(updated);
    assert_eq!(actual_overrides, expected_overrides);

    // Prefix behavior: preserve the `ct:` prefix from the root for the workbook override.
    assert!(
        updated.contains("<ct:Override"),
        "expected output to preserve `ct:` prefix, got:\n{updated}"
    );
    assert!(
        !updated.contains("<Override"),
        "should not introduce unprefixed Override tags, got:\n{updated}"
    );
}

#[test]
fn streaming_package_enforce_workbook_kind_addin_updates_only_workbook_override() {
    let (input, content_types) = streaming_pkg_fixture_with_prefixed_content_types();

    let mut pkg = StreamingXlsxPackage::from_reader(Cursor::new(input)).unwrap();
    pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledAddIn)
        .unwrap();

    let updated_bytes = pkg.read_part("[Content_Types].xml").unwrap().unwrap();
    let updated = std::str::from_utf8(&updated_bytes).unwrap();

    let original_overrides = content_types_override_map(&content_types);
    let mut expected_overrides = original_overrides.clone();
    expected_overrides.insert(
        "/xl/workbook.xml".to_string(),
        WorkbookKind::MacroEnabledAddIn
            .workbook_content_type()
            .to_string(),
    );

    let actual_overrides = content_types_override_map(updated);
    assert_eq!(actual_overrides, expected_overrides);
}
