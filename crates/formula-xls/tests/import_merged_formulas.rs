use std::io::Write;

use formula_model::{CellRef, Range};

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn anchors_formulas_inside_merged_regions_to_top_left_cell() {
    let bytes = xls_fixture_builder::build_merged_non_anchor_formula_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("MergedFormula")
        .expect("MergedFormula missing");

    let merge_range = Range::from_a1("A1:B1").unwrap();
    assert!(
        sheet.merged_regions.iter().any(|r| r.range == merge_range),
        "missing expected merged range {merge_range:?}"
    );

    let a1 = CellRef::from_a1("A1").unwrap();
    let b1 = CellRef::from_a1("B1").unwrap();
    assert_eq!(sheet.merged_regions.resolve_cell(b1), a1);

    // The fixture stores the FORMULA record only at B1, but merged-region semantics dictate that
    // the top-left anchor owns the formula.
    let formula = sheet
        .formula(a1)
        .expect("expected formula on merged-cell anchor");
    assert_eq!(formula, "1+1");
    assert_parseable_formula(formula);

    // Ensure we did not store a separate formula entry on the non-anchor cell.
    let b1_stored_formula = sheet
        .iter_cells()
        .find_map(|(cell_ref, cell)| {
            (cell_ref == b1).then(|| cell.formula.as_deref().map(|s| s.to_string()))
        })
        .flatten();
    assert!(
        b1_stored_formula.is_none(),
        "expected no formula stored for non-anchor B1; got {b1_stored_formula:?}"
    );
}
