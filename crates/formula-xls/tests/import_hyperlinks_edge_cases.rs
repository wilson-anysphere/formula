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
fn imports_continued_hyperlinks_with_utf16_splits_and_embedded_nuls() {
    let bytes = xls_fixture_builder::build_hyperlink_edge_cases_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("EdgeCases")
        .expect("EdgeCases missing");

    // The fixture contains:
    // - one malformed hyperlink (B1) that should be skipped
    // - one valid continued internal hyperlink (A1)
    assert_eq!(
        sheet.hyperlinks.len(),
        1,
        "expected only the valid continued hyperlink to be imported; hyperlinks={:?}",
        sheet.hyperlinks
    );

    let link = &sheet.hyperlinks[0];
    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::Internal {
            sheet: "EdgeCases".to_string(),
            cell: CellRef::from_a1("B2").unwrap(),
        }
    );

    // Embedded NULs in the BIFF payload should be truncated for best-effort compatibility.
    assert_eq!(link.display.as_deref(), Some("Display"));
    assert_eq!(link.tooltip.as_deref(), Some("Tooltip"));
}

#[test]
fn warns_and_skips_malformed_hyperlink_records() {
    let bytes = xls_fixture_builder::build_hyperlink_edge_cases_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("EdgeCases")
        .expect("EdgeCases missing");

    // The malformed hyperlink is anchored at B1; it should not be imported.
    assert!(
        sheet.hyperlink_at(CellRef::from_a1("B1").unwrap()).is_none(),
        "expected malformed hyperlink to be skipped"
    );

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("failed to decode HLINK record")),
        "expected warning about malformed HLINK record; warnings={:?}",
        result.warnings
    );
}

#[test]
fn trims_embedded_nuls_in_url_moniker_strings() {
    let bytes = xls_fixture_builder::build_url_hyperlink_embedded_nul_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result.workbook.sheet_by_name("UrlNul").expect("UrlNul missing");
    assert_eq!(sheet.hyperlinks.len(), 1);
    let link = &sheet.hyperlinks[0];

    assert_eq!(link.range, Range::from_a1("A1").unwrap());
    assert_eq!(
        link.target,
        HyperlinkTarget::ExternalUrl {
            uri: "https://example.com".to_string()
        }
    );
    assert_eq!(link.display.as_deref(), Some("Example"));
    // Tooltip doesn't contain NULs in this fixture.
    assert_eq!(link.tooltip.as_deref(), Some("Tooltip"));
    assert!(result.warnings.is_empty(), "warnings={:?}", result.warnings);
}
