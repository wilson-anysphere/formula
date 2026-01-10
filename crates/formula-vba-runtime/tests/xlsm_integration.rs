use formula_vba_runtime::{parse_program, InMemoryWorkbook, VbaRuntime};

#[test]
fn executes_macro_from_xlsm_fixture() {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/macros/basic.xlsm"
    );
    let bytes = std::fs::read(fixture_path).expect("fixture xlsm exists");
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("valid xlsm package");
    let vba_bin = pkg.vba_project_bin().expect("vbaProject.bin present");
    let project = formula_vba::VBAProject::parse(vba_bin).expect("parse VBA project");
    assert!(!project.modules.is_empty());

    let combined = project
        .modules
        .iter()
        .map(|m| m.code.as_str())
        .collect::<Vec<_>>()
        .join("\n\n");
    let program = parse_program(&combined).expect("parse VBA code");
    let runtime = VbaRuntime::new(program);

    let mut workbook = InMemoryWorkbook::new();
    runtime
        .execute(&mut workbook, "WriteCells", &[])
        .expect("run macro");

    assert!(
        matches!(
            workbook.get_value_a1("Sheet1", "A1"),
            Ok(formula_vba_runtime::VbaValue::String(ref s)) if s == "Written"
        ),
        "expected A1 to be written, got: {:?}",
        workbook.get_value_a1("Sheet1", "A1")
    );
    assert!(
        matches!(
            workbook.get_value_a1("Sheet1", "B2"),
            Ok(formula_vba_runtime::VbaValue::Double(v)) if (v - 42.0).abs() < f64::EPSILON
        ),
        "expected B2 to be 42, got: {:?}",
        workbook.get_value_a1("Sheet1", "B2")
    );
}
