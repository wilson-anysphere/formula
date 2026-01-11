use formula_engine::{eval::parse_a1, locale, Engine, Value};

#[test]
fn canonicalize_and_localize_round_trip_for_de_de() {
    let localized = "=SUMME(1,5;2,5)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(1.5,2.5)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_supports_thousands_and_leading_decimal_in_de_de() {
    let canonical = locale::canonicalize_formula("=SUMME(1.234,56;,5)", &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(1234.56,.5)");
}

#[test]
fn canonicalize_and_localize_array_literals_for_de_de() {
    let localized = "=SUMME({1\\2;3\\4})";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM({1,2;3,4})");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_and_localize_unions_for_de_de() {
    let localized = "(A1;B1)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "(A1,B1)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn translates_xlfn_prefixed_functions() {
    let localized = "=_xlfn.SEQUENZ(1;2)";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=_xlfn.SEQUENCE(1,2)");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_and_localize_round_trip_for_fr_fr_and_es_es() {
    let fr = "=SOMME(1,5;2,5)";
    let fr_canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(fr_canon, "=SUM(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&fr_canon, &locale::FR_FR).unwrap(),
        fr
    );

    let es = "=SUMA(1,5;2,5)";
    let es_canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(es_canon, "=SUM(1.5,2.5)");
    assert_eq!(
        locale::localize_formula(&es_canon, &locale::ES_ES).unwrap(),
        es
    );
}

#[test]
fn structured_reference_separators_are_not_translated() {
    let canonical = "=SUM(Table1[[#Headers],[Qty]],1)";
    let localized = locale::localize_formula(canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized, "=SUMME(Table1[[#Headers],[Qty]];1)");

    let canonical_roundtrip = locale::canonicalize_formula(&localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical_roundtrip, canonical);
}

#[test]
fn canonicalize_and_localize_boolean_literals() {
    let de = "=WENN(WAHR;1;0)";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(locale::localize_formula(&canon, &locale::DE_DE).unwrap(), de);

    let fr = "=SI(VRAI;1;0)";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), fr);

    let es = "=SI(VERDADERO;1;0)";
    let canon = locale::canonicalize_formula(es, &locale::ES_ES).unwrap();
    assert_eq!(canon, "=IF(TRUE,1,0)");
    assert_eq!(locale::localize_formula(&canon, &locale::ES_ES).unwrap(), es);
}

#[test]
fn localized_boolean_keywords_are_not_translated_inside_structured_refs() {
    // `WAHR` is the de-DE TRUE keyword, but table names can still be identifiers; separators
    // inside structured refs should never be touched by translation.
    let localized = "=SUMME(WAHR[Col])";
    let canonical = locale::canonicalize_formula(localized, &locale::DE_DE).unwrap();
    assert_eq!(canonical, "=SUM(WAHR[Col])");

    let localized_roundtrip = locale::localize_formula(&canonical, &locale::DE_DE).unwrap();
    assert_eq!(localized_roundtrip, localized);
}

#[test]
fn canonicalize_and_localize_error_literals() {
    let de = "=#WERT!";
    let canon = locale::canonicalize_formula(de, &locale::DE_DE).unwrap();
    assert_eq!(canon, "=#VALUE!");
    assert_eq!(locale::localize_formula(&canon, &locale::DE_DE).unwrap(), de);

    let fr = "=#VALEUR!";
    let canon = locale::canonicalize_formula(fr, &locale::FR_FR).unwrap();
    assert_eq!(canon, "=#VALUE!");
    assert_eq!(locale::localize_formula(&canon, &locale::FR_FR).unwrap(), fr);
}

#[test]
fn engine_accepts_localized_formulas_and_persists_canonical() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula_localized("Sheet1", "A1", "=SUMME(1,5;2,5)", &locale::DE_DE)
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=SUM(1.5,2.5)"));

    let displayed = locale::localize_formula(engine.get_cell_formula("Sheet1", "A1").unwrap(), &locale::DE_DE).unwrap();
    assert_eq!(displayed, "=SUMME(1,5;2,5)");
}

#[test]
fn engine_accepts_localized_r1c1_formulas_and_persists_canonical_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();

    engine
        .set_cell_formula_localized_r1c1(
            "Sheet1",
            "C5",
            "=SUMME(R[-4]C[-2];R[-4]C[-1])",
            &locale::DE_DE,
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(3.0));
    assert_eq!(engine.get_cell_formula("Sheet1", "C5"), Some("=SUM(A1,B1)"));
}

#[test]
fn engine_accepts_localized_spilling_formulas() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula_localized("Sheet1", "A1", "=SEQUENZ(2;2)", &locale::DE_DE)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=SEQUENCE(2,2)"));
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));

    assert_eq!(
        engine.spill_range("Sheet1", "A1"),
        Some((parse_a1("A1").unwrap(), parse_a1("B2").unwrap()))
    );

    let localized = locale::localize_formula(engine.get_cell_formula("Sheet1", "A1").unwrap(), &locale::DE_DE).unwrap();
    assert_eq!(localized, "=SEQUENZ(2;2)");
}
