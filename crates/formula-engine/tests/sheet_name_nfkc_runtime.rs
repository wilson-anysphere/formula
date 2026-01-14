use formula_engine::{Engine, Value};

#[test]
fn runtime_sheet_name_lookup_uses_nfkc_case_insensitive_matching() {
    let mut engine = Engine::new();

    // U+212A KELVIN SIGN (K) is compatibility-equivalent (NFKC) to ASCII 'K'.
    //
    // Excel treats sheet names as case-insensitive across Unicode and applies compatibility
    // normalization, so runtime-resolved sheet names (e.g. via INDIRECT) should match even when
    // the query uses a compatibility-equivalent form.
    engine.set_cell_value("Kelvin", "A1", 42.0).unwrap();

    // INDIRECT parses the reference text at runtime and resolves the sheet name via
    // `ValueResolver::sheet_id`.
    engine
        .set_cell_formula("Sheet1", "A1", "=INDIRECT(\"KELVIN!A1\")")
        .unwrap();

    // SHEET("name") also resolves sheet names at runtime.
    engine
        .set_cell_formula("Sheet1", "A2", "=SHEET(\"Kelvin\")")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
    // "Kelvin" was created first, so it is sheet number 1.
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
}
