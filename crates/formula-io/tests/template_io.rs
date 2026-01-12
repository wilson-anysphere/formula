use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook};
use formula_xlsx::XlsxPackage;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn reopen_pkg(path: &std::path::Path) -> XlsxPackage {
    let bytes = std::fs::read(path).expect("read saved workbook");
    XlsxPackage::from_bytes(&bytes).expect("re-open saved package")
}

#[test]
fn saving_xlsm_as_xltx_strips_vba_and_sets_template_content_type() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xltx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let pkg = reopen_pkg(&out_path);

    assert!(
        pkg.vba_project_bin().is_none(),
        "expected VBA project to be stripped when saving `.xlsm` as `.xltx`"
    );
    assert!(
        pkg.part("xl/vbaProject.bin").is_none(),
        "expected xl/vbaProject.bin to be removed"
    );

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        ),
        "expected `[Content_Types].xml` to advertise a template workbook main content type"
    );
}

#[test]
fn saving_xlsm_as_xltm_preserves_vba_and_sets_template_macro_content_type() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xltm");
    save_workbook(&wb, &out_path).expect("save workbook");

    let pkg = reopen_pkg(&out_path);

    assert!(
        pkg.vba_project_bin().is_some(),
        "expected VBA project to be preserved when saving `.xlsm` as `.xltm`"
    );

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains("application/vnd.ms-excel.template.macroEnabled.main+xml"),
        "expected `[Content_Types].xml` to advertise a macro-enabled template workbook main content type"
    );
    assert!(
        !content_types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected `[Content_Types].xml` workbook main content type to match `.xltm` (not `.xlsm`)"
    );
}

#[test]
fn saving_xlsm_as_xlam_preserves_vba_and_sets_addin_macro_content_type() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlam");
    save_workbook(&wb, &out_path).expect("save workbook");

    let pkg = reopen_pkg(&out_path);

    assert!(
        pkg.vba_project_bin().is_some(),
        "expected VBA project to be preserved when saving `.xlsm` as `.xlam`"
    );

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains("application/vnd.ms-excel.addin.macroEnabled.main+xml"),
        "expected `[Content_Types].xml` to advertise an add-in workbook main content type"
    );
    assert!(
        !content_types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected `[Content_Types].xml` workbook main content type to match `.xlam` (not `.xlsm`)"
    );
}
