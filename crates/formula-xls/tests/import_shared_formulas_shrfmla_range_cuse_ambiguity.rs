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
fn resolves_shrfmla_by_range_without_confusing_cuse_for_ref8_header_bytes() {
    let bytes = xls_fixture_builder::build_shared_formula_shrfmla_range_cuse_ambiguity_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("ShrfmlaCuseAmbiguity")
        .expect("ShrfmlaCuseAmbiguity sheet missing");

    let b2 = CellRef::from_a1("B2").unwrap();
    let formula = sheet.formula(b2).expect("expected formula in B2");

    assert_eq!(formula, "C2+1");
    assert_parseable_formula(formula);
}

