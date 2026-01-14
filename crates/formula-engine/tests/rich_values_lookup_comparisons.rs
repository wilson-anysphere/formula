use formula_engine::{
    value::{EntityValue, RecordValue},
    Engine, Value,
};
use std::collections::HashMap;

#[test]
fn match_compares_entity_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=MATCH(B1, A1:A1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
}

#[test]
fn match_wildcards_compare_against_entity_display_string() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "app*").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=MATCH(B1, A1:A1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
}

#[test]
fn match_wildcards_compare_against_record_display_field() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", r#"=MATCH("App*", A1:A1, 0)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected MATCH formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    }
}

#[test]
fn match_compares_against_record_display_field_for_exact_match() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", r#"=MATCH("Apple", A1:A1, 0)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected MATCH formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    }
}

#[test]
fn match_wildcards_compare_record_display_fields_for_pattern_and_candidate() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "CandidateFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_value(
                "Sheet1",
                "B1",
                Value::Record(RecordValue {
                    display: "PatternFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("App*"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "C1", r#"=MATCH(B1, A1:A1, 0)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected MATCH formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    }
}

#[test]
fn xlookup_compares_entity_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "C1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=XLOOKUP(B1, A1:A1, C1:C1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
}

#[test]
fn vlookup_compares_record_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Record(RecordValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", 42.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=VLOOKUP(\"apple\", A1:B1, 2, FALSE)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
}

#[test]
fn vlookup_compares_against_record_display_field_for_exact_match() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "C1", r#"=VLOOKUP("Apple", A1:B1, 2, FALSE)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected VLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
    }
}

#[test]
fn vlookup_wildcards_compare_record_display_fields_for_pattern_and_candidate() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "CandidateFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        engine
            .set_cell_value(
                "Sheet1",
                "C1",
                Value::Record(RecordValue {
                    display: "PatternFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("App*"))]),
                }),
            )
            .unwrap();

        engine
            .set_cell_formula("Sheet1", "D1", r#"=VLOOKUP(C1, A1:B1, 2, FALSE)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected VLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(42.0));
    }
}

#[test]
fn xmatch_compares_record_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Record(RecordValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=XMATCH(B1, A1:A1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
}

#[test]
fn xmatch_compares_against_record_display_field_for_exact_match() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", r#"=XMATCH("Apple", A1:A1, 0)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected XMATCH formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    }
}

#[test]
fn hlookup_wildcards_compare_record_display_fields_for_pattern_and_candidate() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "CandidateFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "A2", 42.0).unwrap();
        engine
            .set_cell_value(
                "Sheet1",
                "B1",
                Value::Record(RecordValue {
                    display: "PatternFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("App*"))]),
                }),
            )
            .unwrap();

        engine
            .set_cell_formula("Sheet1", "C1", r#"=HLOOKUP(B1, A1:A2, 2, FALSE)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected HLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
    }
}

#[test]
fn hlookup_compares_against_record_display_field_for_exact_match() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "A2", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", r#"=HLOOKUP("Apple", A1:A2, 2, FALSE)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected HLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(42.0));
    }
}

#[test]
fn xmatch_wildcards_compare_against_record_display_field() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", r#"=XMATCH("App*", A1:A1, 2)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    }
}

#[test]
fn xlookup_wildcards_compare_record_display_fields_for_pattern_and_candidate() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "CandidateFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine
            .set_cell_value(
                "Sheet1",
                "B1",
                Value::Record(RecordValue {
                    display: "PatternFallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("App*"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "C1", 42.0).unwrap();

        engine
            .set_cell_formula(
                "Sheet1",
                "D1",
                r#"=XLOOKUP(B1, A1:A1, C1:C1, "not found", 2)"#,
            )
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected XLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(42.0));
    }
}

#[test]
fn xlookup_compares_against_record_display_field_for_exact_match() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Record(RecordValue {
                    display: "Fallback".to_string(),
                    display_field: Some("Name".to_string()),
                    fields: HashMap::from([("Name".to_string(), Value::from("Apple"))]),
                }),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "B1", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "C1", r#"=XLOOKUP("Apple", A1:A1, B1:B1)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        if bytecode_enabled {
            assert!(
                engine.bytecode_program_count() > 0,
                "expected XLOOKUP formula to compile to bytecode for this test"
            );
        }

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
    }
}
