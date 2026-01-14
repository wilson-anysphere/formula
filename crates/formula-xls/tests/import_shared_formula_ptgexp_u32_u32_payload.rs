use std::io::Write;

use formula_model::CellRef;

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_shared_formula_with_u32_u32_ptgexp_payload() {
    // The fixture uses a shared formula in B2:B3, where the follower cell B3 stores `PtgExp`
    // coordinates as row u32 + col u32 (8 bytes). The importer should still resolve it against the
    // SHRFMLA definition and decode the correct shifted formula.
    let bytes = xls_fixture_builder::build_shared_formula_ptgexp_u32_row_u32_col_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Shared")
        .expect("Shared sheet missing");

    let b3 = CellRef::from_a1("B3").expect("valid ref");
    let formula = sheet.formula(b3).expect("expected formula in Shared!B3");
    assert_eq!(formula, "A3*2");
    assert_parseable_formula(formula);
}

