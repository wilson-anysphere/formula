use std::io::Write;

use formula_model::{CellRef, HyperlinkTarget, Range};

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

#[test]
fn imports_biff_hyperlinks_internal() {
    let bytes = xls_fixture_builder::build_internal_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Internal")
        .expect("Internal missing");

    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];
    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::Internal {
            sheet: "Internal".to_string(),
            cell: CellRef::from_a1("B2").unwrap()
        }
    );
    assert_eq!(link.display.as_deref(), Some("Go to B2"));
    assert_eq!(link.tooltip.as_deref(), Some("Internal tooltip"));
}

#[test]
fn imports_biff_hyperlinks_mailto() {
    let bytes = xls_fixture_builder::build_mailto_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("Mail").expect("Mail missing");
    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::Email {
            uri: "mailto:test@example.com".to_string()
        }
    );
    assert_eq!(link.display.as_deref(), Some("Email"));
    assert_eq!(link.tooltip.as_deref(), Some("Email tooltip"));
}

#[test]
fn imports_biff_hyperlinks_continued_record() {
    let bytes = xls_fixture_builder::build_continued_hyperlink_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Continued")
        .expect("Continued missing");

    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    match &link.target {
        HyperlinkTarget::ExternalUrl { uri } => {
            assert!(
                uri.starts_with("https://example.com/"),
                "unexpected uri {uri}"
            );
            assert!(
                uri.len() > "https://example.com/".len() + 50,
                "expected long continued uri, got {uri}"
            );
        }
        other => panic!("expected ExternalUrl hyperlink target, got {other:?}"),
    }
    assert_eq!(link.display.as_deref(), Some("Example"));
    assert_eq!(link.tooltip.as_deref(), Some("Example tooltip"));
}
