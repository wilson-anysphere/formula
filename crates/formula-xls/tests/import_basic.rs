use std::path::PathBuf;

use formula_engine::{parse_formula, ParseOptions};
use formula_model::{CellRef, CellValue};

fn assert_parseable(formula_body: &str) {
    let formula = format!("={formula_body}");
    parse_formula(&formula, ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected formula to be parseable, formula={formula:?}, err={e:?}"));
}

#[test]
fn imports_basic_xls() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("basic.xls");

    let result = formula_xls::import_xls_path(&fixture_path).expect("import xls");

    assert_eq!(result.source.format, formula_xls::SourceFormat::Xls);
    assert_eq!(result.source.default_save_extension(), "xlsx");

    let sheet1 = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    assert_eq!(
        sheet1.value_a1("A1").unwrap(),
        CellValue::String("Hello".to_owned())
    );
    assert_eq!(sheet1.value_a1("B2").unwrap(), CellValue::Number(123.0));

    let c3 = CellRef::from_a1("C3").unwrap();
    assert_eq!(sheet1.formula(c3), Some("B2*2"));
    assert_parseable("B2*2");

    let sheet2 = result
        .workbook
        .sheet_by_name("Second")
        .expect("Second missing");
    assert_eq!(
        sheet2.value_a1("A1").unwrap(),
        CellValue::String("Second sheet".to_owned())
    );
}
