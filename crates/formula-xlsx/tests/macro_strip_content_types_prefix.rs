use std::io::{Cursor, Read, Write};

use formula_xlsx::strip_vba_project_streaming;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[test]
fn strip_vba_project_streaming_preserves_content_types_prefix() {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <ct:Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</ct:Types>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    // Minimal workbook payload; not strictly required for stripping, but keeps the package shape
    // closer to a real XLSM.
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><workbook/>"#)
        .unwrap();

    // Macro payload to be deleted.
    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(b"dummy").unwrap();

    let input_bytes = zip.finish().unwrap().into_inner();

    let mut input = Cursor::new(input_bytes);
    let mut output = Cursor::new(Vec::new());
    strip_vba_project_streaming(&mut input, &mut output).unwrap();

    let out_bytes = output.into_inner();
    let mut archive = ZipArchive::new(Cursor::new(out_bytes)).unwrap();
    let mut file = archive.by_name("[Content_Types].xml").unwrap();
    let mut out_ct = String::new();
    file.read_to_string(&mut out_ct).unwrap();

    roxmltree::Document::parse(&out_ct).expect("valid xml");

    assert!(out_ct.contains("<ct:Override"));
    assert!(!out_ct.contains("<Override"));
    assert!(out_ct
        .contains("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
    assert!(!out_ct.contains("macroEnabled.main+xml"));
}
