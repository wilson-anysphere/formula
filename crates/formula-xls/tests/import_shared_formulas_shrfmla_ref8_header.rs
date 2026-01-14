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
fn imports_shared_formulas_when_shrfmla_range_header_uses_ref8_starting_at_col_a() {
    let bytes = xls_fixture_builder::build_shared_formula_shrfmla_ref8_header_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("SharedRef8")
        .expect("SharedRef8 sheet missing");

    for a1 in ["A1", "A2", "B1", "B2"] {
        let cell = CellRef::from_a1(a1).expect("valid ref");
        let formula = sheet.formula(cell).expect("expected formula");
        assert_eq!(formula, "\"X\"", "unexpected formula in {a1}");
        assert_parseable_formula(formula);
    }
}

