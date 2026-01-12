use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook, Workbook};
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
fn saving_xlsx_with_xlm_macrosheets_as_xltx_strips_macrosheets() {
    // Start from a macro-free XLSX fixture and inject an XLM macro sheet part. This simulates an
    // XLSX package containing macro-capable content without a VBA project.
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&path).expect("open workbook");
    let Workbook::Xlsx(mut pkg) = wb else {
        panic!("expected Workbook::Xlsx");
    };

    // Any content is fine for this test; macro stripping is based on the part name.
    pkg.set_part("xl/macrosheets/sheet1.xml", b"<worksheet/>".to_vec());
    assert!(
        pkg.macro_presence().has_xlm_macrosheets,
        "expected test package to contain XLM macro sheet parts"
    );
    assert!(
        pkg.vba_project_bin().is_none(),
        "expected test package to contain no VBA project"
    );

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xltx");
    save_workbook(&Workbook::Xlsx(pkg), &out_path).expect("save workbook");

    let saved = reopen_pkg(&out_path);
    assert!(
        saved.part("xl/macrosheets/sheet1.xml").is_none(),
        "expected XLM macro sheet parts to be stripped when saving to `.xltx`"
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

#[test]
fn saving_xltx_as_xlsx_sets_workbook_content_type() {
    // Start from an XLSM fixture so we can generate a template via the save path (which also strips VBA).
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");

    // First, save as a macro-free template.
    let xltx_path = dir.path().join("out.xltx");
    save_workbook(&wb, &xltx_path).expect("save as xltx");

    // Then, re-save that template as a normal workbook.
    let xltx_pkg = reopen_pkg(&xltx_path);
    let xlsx_path = dir.path().join("out.xlsx");
    save_workbook(&Workbook::Xlsx(xltx_pkg), &xlsx_path).expect("save as xlsx");

    let pkg = reopen_pkg(&xlsx_path);
    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        ),
        "expected workbook main content type to match `.xlsx`"
    );
    assert!(
        !content_types.contains(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        ),
        "expected workbook main content type to not be `.xltx` when saving as `.xlsx`"
    );
}

#[test]
fn saving_xltm_as_xlsm_sets_workbook_content_type() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");

    // First, save as a macro-enabled template.
    let xltm_path = dir.path().join("out.xltm");
    save_workbook(&wb, &xltm_path).expect("save as xltm");

    // Then, re-save that template as a macro-enabled workbook.
    let xltm_pkg = reopen_pkg(&xltm_path);
    let xlsm_path = dir.path().join("out.xlsm");
    save_workbook(&Workbook::Xlsx(xltm_pkg), &xlsm_path).expect("save as xlsm");

    let pkg = reopen_pkg(&xlsm_path);
    assert!(pkg.vba_project_bin().is_some(), "expected vbaProject.bin to be preserved");

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        content_types.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        "expected workbook main content type to match `.xlsm`"
    );
    assert!(
        !content_types.contains("application/vnd.ms-excel.template.macroEnabled.main+xml"),
        "expected workbook main content type to not be `.xltm` when saving as `.xlsm`"
    );
}
