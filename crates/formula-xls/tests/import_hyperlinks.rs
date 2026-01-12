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
fn imports_biff_hyperlinks() {
    let bytes = xls_fixture_builder::build_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Links")
        .expect("Links missing");

    assert_eq!(sheet.hyperlinks.len(), 1, "hyperlinks={:?}", sheet.hyperlinks);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::ExternalUrl {
            uri: "https://example.com".to_string()
        }
    );
    assert_eq!(link.display.as_deref(), Some("Example"));
    assert_eq!(link.tooltip.as_deref(), Some("Example tooltip"));
}

