use std::io::Write;

use formula_model::Range;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff8_autofilter12_future_record_best_effort() {
    let bytes = xls_fixture_builder::build_autofilter12_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Filter12")
        .expect("Filter12 missing");
    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());

    // Best-effort: prefer a decoded filter column, but accept a deterministic warning if the
    // payload layout is unsupported.
    if !auto_filter.filter_columns.is_empty() {
        let col0 = auto_filter
            .filter_columns
            .iter()
            .find(|c| c.col_id == 0)
            .expect("expected colId=0 filter column");
        assert!(
            col0.values.iter().any(|v| v == "Alice"),
            "expected 'Alice' in filter values; col0={col0:?}",
        );
        assert!(
            col0.values.iter().any(|v| v == "Bob"),
            "expected 'Bob' in filter values; col0={col0:?}",
        );
    } else {
        assert!(
            result
                .warnings
                .iter()
                .any(|w| w.message.contains("unsupported AutoFilter12")),
            "expected unsupported-AutoFilter12 warning; warnings={:?}",
            result.warnings
        );
    }
}

