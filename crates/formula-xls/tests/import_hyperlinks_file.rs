use std::io::Write;

use formula_model::{HyperlinkTarget, Range};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff_hyperlinks_file_drive_path() {
    let bytes = xls_fixture_builder::build_file_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("File").expect("File missing");
    assert_eq!(sheet.hyperlinks.len(), 1, "hyperlinks={:?}", sheet.hyperlinks);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::ExternalUrl {
            uri: "file:///C:/foo/bar.txt".to_string()
        }
    );
}

#[test]
fn imports_biff_hyperlinks_file_unc_path() {
    let bytes = xls_fixture_builder::build_unc_file_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("UNC").expect("UNC missing");
    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::ExternalUrl {
            uri: "file://server/share/file.xlsx".to_string()
        }
    );
}

#[test]
fn imports_biff_hyperlinks_file_unicode_path_prefers_unicode_extension() {
    let bytes = xls_fixture_builder::build_unicode_file_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Unicode")
        .expect("Unicode missing");
    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::ExternalUrl {
            uri: "file:///C:/foo/%E6%97%A5%E6%9C%AC.txt".to_string()
        }
    );
}
