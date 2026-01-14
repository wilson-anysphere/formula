use formula_engine::{Engine, Value};

fn eval_cp932(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set cell formula");
    // Use the single-threaded recalc path in tests to avoid initializing a global Rayon pool.
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "A1")
}

#[test]
fn leftb_rightb_truncate_at_dbcs_character_boundaries() {
    let mut engine = Engine::new();
    engine.set_text_codepage(932);

    // Under cp932, "漢" is a 2-byte character. When asked to return a range that would split a
    // DBCS code unit, the engine truncates to avoid returning partial characters.
    assert_eq!(
        eval_cp932(&mut engine, r#"=LEFTB("漢",1)"#),
        Value::Text(String::new())
    );
    assert_eq!(
        eval_cp932(&mut engine, r#"=LEFTB("漢",2)"#),
        Value::Text("漢".to_string())
    );

    assert_eq!(
        eval_cp932(&mut engine, r#"=RIGHTB("漢",1)"#),
        Value::Text(String::new())
    );
    assert_eq!(
        eval_cp932(&mut engine, r#"=RIGHTB("漢",2)"#),
        Value::Text("漢".to_string())
    );
}

#[test]
fn midb_rounds_start_forward_and_truncates_length_to_character_boundaries() {
    let mut engine = Engine::new();
    engine.set_text_codepage(932);

    // "A漢B" encoded in cp932:
    // - "A" = 1 byte
    // - "漢" = 2 bytes
    // - "B" = 1 byte
    //
    // MIDB uses a 1-indexed byte start offset. When the start/length would split a DBCS character,
    // the engine rounds the slice to character boundaries, potentially producing an empty string.
    assert_eq!(
        eval_cp932(&mut engine, r#"=MIDB("A漢B",2,1)"#),
        Value::Text(String::new())
    );
    assert_eq!(
        eval_cp932(&mut engine, r#"=MIDB("A漢B",3,1)"#),
        Value::Text(String::new())
    );
    assert_eq!(
        eval_cp932(&mut engine, r#"=MIDB("A漢B",3,2)"#),
        Value::Text("B".to_string())
    );
}

#[test]
fn replaceb_inserts_at_aligned_boundaries_when_length_would_split_dbcs_character() {
    let mut engine = Engine::new();
    engine.set_text_codepage(932);

    // Like MIDB, REPLACEB operates on byte offsets under DBCS codepages. When the byte length would
    // split a character, the engine collapses the replacement range to the nearest character
    // boundary (potentially to an empty range, which behaves like insertion).
    assert_eq!(
        eval_cp932(&mut engine, r#"=REPLACEB("A漢B",2,1,"Z")"#),
        Value::Text("AZ漢B".to_string())
    );
    assert_eq!(
        eval_cp932(&mut engine, r#"=REPLACEB("A漢B",3,1,"Z")"#),
        Value::Text("A漢ZB".to_string())
    );
}

#[test]
fn findb_rounds_start_offset_forward_when_it_lands_mid_dbcs_character() {
    let mut engine = Engine::new();
    engine.set_text_codepage(932);

    // Byte offsets are 1-indexed. A start position of 3 lands inside the 2-byte "漢" (bytes 2-3),
    // so the engine rounds the start forward to the next character boundary and begins searching
    // from "B".
    assert_eq!(
        eval_cp932(&mut engine, r#"=FINDB("B","A漢B",3)"#),
        Value::Number(4.0)
    );
}
