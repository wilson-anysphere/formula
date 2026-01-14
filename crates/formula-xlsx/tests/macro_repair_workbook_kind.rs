use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

    for (name, bytes) in files {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn macro_repair_upgrades_xltx_workbook_content_type_when_vba_is_present() {
    // Regression test for workbook-kind preservation/upgrade behavior:
    //
    // If a template workbook advertises the `.xltx` content type but contains `xl/vbaProject.bin`,
    // Excel expects the workbook main content type to be upgraded to `.xltm` (macro-enabled
    // template), not forced to `.xlsm`.
    let input_bytes = build_zip(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"/>
</Types>"#,
        ),
        (
            "xl/workbook.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ),
        ("xl/vbaProject.bin", b"fake-vba"),
    ]);

    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    let out_bytes = pkg.write_to_bytes().expect("write package");
    let out_pkg = XlsxPackage::from_bytes(&out_bytes).expect("re-open written package");

    let content_types =
        std::str::from_utf8(out_pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains("application/vnd.ms-excel.template.macroEnabled.main+xml"),
        "expected workbook main content type to be upgraded to macro-enabled template"
    );
    assert!(
        !content_types.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"),
        "expected macro-free template workbook content type to be removed"
    );
    assert!(
        !content_types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected workbook main content type to not be forced to `.xlsm`"
    );
}

#[test]
fn macro_repair_upgrades_xltx_workbook_content_type_when_vba_is_present_with_noncanonical_part_name() {
    // Some XLSM producers (or intermediaries) store ZIP entry names non-canonically (case
    // differences, `\\` separators, etc). Macro repair should still detect the VBA project payload
    // and upgrade the workbook main content type accordingly.
    let input_bytes = build_zip(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"/>
</Types>"#,
        ),
        (
            "xl/workbook.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ),
        // Intentionally non-canonical ZIP entry name for the VBA project payload.
        ("XL\\VBAPROJECT.BIN", b"fake-vba"),
    ]);

    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    let out_bytes = pkg.write_to_bytes().expect("write package");
    let out_pkg = XlsxPackage::from_bytes(&out_bytes).expect("re-open written package");

    let content_types =
        std::str::from_utf8(out_pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains("application/vnd.ms-excel.template.macroEnabled.main+xml"),
        "expected workbook main content type to be upgraded to macro-enabled template"
    );
    assert!(
        !content_types.contains("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"),
        "expected macro-free template workbook content type to be removed"
    );
    assert!(
        !content_types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected workbook main content type to not be forced to `.xlsm`"
    );
}
