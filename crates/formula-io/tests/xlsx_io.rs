use std::io::{Cursor, Write};
use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook, Workbook};
use formula_xlsx::XlsxPackage;
use xlsx_diff::Severity;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

#[test]
fn opens_basic_xlsx_fixture() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&path).expect("open workbook");

    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(pkg.part("xl/workbook.xml").is_some());
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn opens_xlsm_fixture_and_preserves_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(
                pkg.vba_project_bin().is_some(),
                "expected xl/vbaProject.bin to be present"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn roundtrips_xlsx_package_without_critical_diffs() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("roundtrip.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let report = xlsx_diff::diff_workbooks(&path, &out_path).expect("diff workbooks");
    let critical = report.count(Severity::Critical);
    assert_eq!(
        critical, 0,
        "expected no critical diffs, got {critical}\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<String>()
    );
}

#[test]
fn saving_xlsm_as_xlsx_strips_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        pkg.vba_project_bin().is_none(),
        "expected VBA project to be stripped when saving `.xlsm` as `.xlsx`"
    );
    assert!(
        pkg.part("xl/vbaProject.bin").is_none(),
        "expected xl/vbaProject.bin to be removed"
    );
    assert!(
        pkg.part("xl/vbaProjectSignature.bin").is_none(),
        "expected xl/vbaProjectSignature.bin to be removed"
    );

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        !content_types.contains("macroEnabled"),
        "expected `[Content_Types].xml` to no longer advertise a macro-enabled workbook"
    );
    assert!(
        !content_types.contains("vbaProject.bin"),
        "expected `[Content_Types].xml` to no longer reference vbaProject.bin"
    );

    let wb_rels =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").expect("workbook rels"))
            .unwrap();
    assert!(
        !wb_rels.contains("vbaProject"),
        "expected workbook relationships to no longer reference vbaProject"
    );
}

#[test]
fn saving_xlsm_as_xlsm_preserves_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlsm");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        pkg.vba_project_bin().is_some(),
        "expected VBA project to be preserved when saving `.xlsm` as `.xlsm`"
    );
}

#[test]
fn opens_basic_xltx_fixture() {
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let tmp = tempfile::tempdir().expect("temp dir");
    let dst = tmp.path().join("basic.xltx");
    std::fs::copy(&src, &dst).expect("copy fixture to .xltx");

    let wb = open_workbook(&dst).expect("open workbook");
    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(pkg.part("xl/workbook.xml").is_some());
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn opens_xltm_and_xlam_as_macro_capable_packages() {
    let src = fixture_path("xlsx/macros/basic.xlsm");
    let tmp = tempfile::tempdir().expect("temp dir");

    for ext in ["xltm", "xlam"] {
        let dst = tmp.path().join(format!("basic.{ext}"));
        std::fs::copy(&src, &dst).expect("copy macro fixture");
        let wb = open_workbook(&dst).expect("open workbook");
        match wb {
            Workbook::Xlsx(pkg) => {
                assert!(
                    pkg.vba_project_bin().is_some(),
                    "expected xl/vbaProject.bin to be present for {ext}"
                );
            }
            other => panic!("expected Workbook::Xlsx, got {other:?}"),
        }
    }
}

#[test]
fn saving_xltm_preserves_vba_and_saving_xltx_strips_vba() {
    let src = fixture_path("xlsx/macros/basic.xlsm");
    let tmp = tempfile::tempdir().expect("temp dir");

    let xltm_path = tmp.path().join("basic.xltm");
    std::fs::copy(&src, &xltm_path).expect("copy macro fixture");

    let original_bytes = std::fs::read(&xltm_path).expect("read xltm bytes");
    let original_pkg = XlsxPackage::from_bytes(&original_bytes).expect("parse original pkg");
    let original_vba = original_pkg
        .vba_project_bin()
        .expect("fixture should have vbaProject.bin")
        .to_vec();

    let wb = open_workbook(&xltm_path).expect("open workbook");

    // Saving as `.xltm` should preserve VBA.
    let saved_xltm = tmp.path().join("saved.xltm");
    save_workbook(&wb, &saved_xltm).expect("save as xltm");
    let saved_bytes = std::fs::read(&saved_xltm).expect("read saved xltm");
    let saved_pkg = XlsxPackage::from_bytes(&saved_bytes).expect("parse saved pkg");
    assert_eq!(
        saved_pkg
            .vba_project_bin()
            .expect("saved xltm should contain vbaProject.bin"),
        original_vba.as_slice()
    );

    // Saving as `.xlam` should also preserve VBA.
    let saved_xlam = tmp.path().join("saved.xlam");
    save_workbook(&wb, &saved_xlam).expect("save as xlam");
    let saved_bytes = std::fs::read(&saved_xlam).expect("read saved xlam");
    let saved_pkg = XlsxPackage::from_bytes(&saved_bytes).expect("parse saved xlam");
    assert_eq!(
        saved_pkg
            .vba_project_bin()
            .expect("saved xlam should contain vbaProject.bin"),
        original_vba.as_slice()
    );

    // Saving as `.xltx` should strip VBA.
    let saved_xltx = tmp.path().join("saved.xltx");
    save_workbook(&wb, &saved_xltx).expect("save as xltx");
    let saved_bytes = std::fs::read(&saved_xltx).expect("read saved xltx");
    let saved_pkg = XlsxPackage::from_bytes(&saved_bytes).expect("parse saved xltx");
    assert!(
        saved_pkg.vba_project_bin().is_none(),
        "expected vbaProject.bin to be removed when saving as .xltx"
    );
}

fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Stored);

    for (name, bytes) in files {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn macro_capable_xlm_package_bytes() -> Vec<u8> {
    build_zip(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
  <Override PartName="/xl/dialogsheets/sheet1.xml" ContentType="application/vnd.ms-excel.dialogsheet+xml"/>
</Types>"#,
        ),
        (
            "xl/workbook.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
</workbook>"#,
        ),
        ("xl/macrosheets/sheet1.xml", br#"<macrosheet/>"#),
        ("xl/dialogsheets/sheet1.xml", br#"<dialogsheet/>"#),
    ])
}

#[test]
fn saving_xlsx_strips_xlm_macrosheets_and_dialogsheets_without_vba_project() {
    // Build a tiny "XLSX-in-disguise" package:
    // - Has XLM macrosheets + dialog sheets
    // - Does *not* have `xl/vbaProject.bin`
    // - Advertises a macro-enabled workbook content type
    //
    // When saving to `.xlsx`, formula-io must strip *all* macro-capable content, not just VBA.
    let input_bytes = macro_capable_xlm_package_bytes();

    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    assert!(
        pkg.vba_project_bin().is_none(),
        "test package should not contain xl/vbaProject.bin"
    );
    assert!(
        pkg.macro_presence().any(),
        "test package should be detected as macro-capable"
    );

    let wb = Workbook::Xlsx(pkg);
    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let out_pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        !out_pkg
            .part_names()
            .any(|name| name.starts_with("xl/macrosheets/")),
        "expected XLM macrosheet parts to be stripped when saving `.xlsx`"
    );
    assert!(
        !out_pkg
            .part_names()
            .any(|name| name.starts_with("xl/dialogsheets/")),
        "expected legacy dialog sheet parts to be stripped when saving `.xlsx`"
    );

    let content_types = std::str::from_utf8(
        out_pkg
            .part("[Content_Types].xml")
            .expect("content types should be preserved/updated"),
    )
    .unwrap();
    assert!(
        !content_types.contains("macroEnabled"),
        "expected `[Content_Types].xml` to no longer advertise a macro-enabled workbook"
    );
}

#[test]
fn saving_xltx_strips_macro_capable_content() {
    // Regression test: `.xltx` (macro-free template) must never contain macro-capable parts, even
    // for packages that only contain XLM macro sheets / dialog sheets (and no vbaProject.bin).
    let input_bytes = macro_capable_xlm_package_bytes();

    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    assert!(
        pkg.vba_project_bin().is_none(),
        "test package should not contain xl/vbaProject.bin"
    );
    assert!(
        pkg.macro_presence().any(),
        "test package should be detected as macro-capable"
    );

    let wb = Workbook::Xlsx(pkg);
    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xltx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let out_pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        !out_pkg
            .part_names()
            .any(|name| name.starts_with("xl/macrosheets/")),
        "expected XLM macrosheet parts to be stripped when saving `.xltx`"
    );
    assert!(
        !out_pkg
            .part_names()
            .any(|name| name.starts_with("xl/dialogsheets/")),
        "expected legacy dialog sheet parts to be stripped when saving `.xltx`"
    );

    let content_types = std::str::from_utf8(
        out_pkg
            .part("[Content_Types].xml")
            .expect("content types should be preserved/updated"),
    )
    .unwrap();
    assert!(
        content_types.contains(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        ),
        "expected `[Content_Types].xml` to advertise a template workbook main content type"
    );
    assert!(
        !content_types.contains("macroEnabled"),
        "expected `[Content_Types].xml` to not advertise a macro-enabled workbook"
    );
}

#[test]
fn saving_xltm_preserves_macro_capable_content() {
    let input_bytes = macro_capable_xlm_package_bytes();
    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    assert!(
        pkg.macro_presence().any(),
        "test package should be detected as macro-capable"
    );

    let wb = Workbook::Xlsx(pkg);
    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xltm");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let out_pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        out_pkg.part("xl/macrosheets/sheet1.xml").is_some(),
        "expected XLM macrosheet parts to be preserved when saving `.xltm`"
    );
    assert!(
        out_pkg.part("xl/dialogsheets/sheet1.xml").is_some(),
        "expected legacy dialog sheet parts to be preserved when saving `.xltm`"
    );
}

#[test]
fn saving_xlam_preserves_macro_capable_content() {
    let input_bytes = macro_capable_xlm_package_bytes();
    let pkg = XlsxPackage::from_bytes(&input_bytes).expect("parse test package");
    assert!(
        pkg.macro_presence().any(),
        "test package should be detected as macro-capable"
    );

    let wb = Workbook::Xlsx(pkg);
    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlam");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let out_pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        out_pkg.part("xl/macrosheets/sheet1.xml").is_some(),
        "expected XLM macrosheet parts to be preserved when saving `.xlam`"
    );
    assert!(
        out_pkg.part("xl/dialogsheets/sheet1.xml").is_some(),
        "expected legacy dialog sheet parts to be preserved when saving `.xlam`"
    );
}
