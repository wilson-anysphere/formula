use formula_engine::locale::ValueLocaleConfig;
use formula_engine::{EditOp, ErrorKind, Value};
use formula_model::{CellRef, Range};

use super::harness::TestSheet;

#[test]
fn byte_count_text_functions_match_non_b_variants_in_en_us() {
    let mut sheet = TestSheet::new();

    // In single-byte locales (en-US), the `*B` functions are expected to behave
    // identically to their non-`B` equivalents.
    assert_eq!(sheet.eval(r#"=LENB("abc")"#), sheet.eval(r#"=LEN("abc")"#));
    assert_eq!(
        sheet.eval(r#"=LEFTB("abc")"#),
        sheet.eval(r#"=LEFT("abc")"#)
    );
    assert_eq!(
        sheet.eval(r#"=LEFTB("abc",2)"#),
        sheet.eval(r#"=LEFT("abc",2)"#)
    );
    assert_eq!(
        sheet.eval(r#"=RIGHTB("abc")"#),
        sheet.eval(r#"=RIGHT("abc")"#)
    );
    assert_eq!(
        sheet.eval(r#"=RIGHTB("abc",2)"#),
        sheet.eval(r#"=RIGHT("abc",2)"#)
    );
    assert_eq!(
        sheet.eval(r#"=MIDB("abc",2,2)"#),
        sheet.eval(r#"=MID("abc",2,2)"#)
    );

    // Smoke test non-ASCII text. Note: real Excel byte-count semantics are locale/codepage
    // dependent; for now the engine treats these identically to the non-`B` versions.
    assert_eq!(sheet.eval(r#"=LENB("é")"#), sheet.eval(r#"=LEN("é")"#));
    assert_eq!(sheet.eval(r#"=LENB("漢")"#), sheet.eval(r#"=LEN("漢")"#));
}

#[test]
fn findb_searchb_replaceb_match_find_search_replace_in_en_us() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval(r#"=FINDB("b","abc")"#),
        sheet.eval(r#"=FIND("b","abc")"#)
    );
    assert_eq!(
        sheet.eval(r#"=FINDB("B","abc")"#),
        sheet.eval(r#"=FIND("B","abc")"#)
    );

    assert_eq!(
        sheet.eval(r#"=SEARCHB("B","abc")"#),
        sheet.eval(r#"=SEARCH("B","abc")"#)
    );
    assert_eq!(
        sheet.eval(r#"=SEARCHB("b","abc",3)"#),
        sheet.eval(r#"=SEARCH("b","abc",3)"#)
    );

    assert_eq!(
        sheet.eval(r#"=REPLACEB("abcdef",2,3,"X")"#),
        sheet.eval(r#"=REPLACE("abcdef",2,3,"X")"#)
    );
}

#[test]
fn asc_and_dbcs_are_identity_transforms_by_default_codepage() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval(r#"=ASC("ABC")"#), Value::Text("ABC".to_string()));
    assert_eq!(
        sheet.eval(r#"=DBCS("ABC")"#),
        Value::Text("ABC".to_string())
    );

    // Non-ASCII smoke test: the engine currently does not implement the locale-specific
    // half-width/full-width conversions that Excel performs in some DBCS locales unless
    // the workbook text codepage is explicitly set to the appropriate DBCS codepage.
    assert_eq!(sheet.eval(r#"=ASC("漢")"#), Value::Text("漢".to_string()));
    assert_eq!(sheet.eval(r#"=DBCS("漢")"#), Value::Text("漢".to_string()));

    // Fullwidth/halfwidth transforms should be gated on the active text codepage.
    assert_eq!(
        sheet.eval(r#"=DBCS("ABC 123")"#),
        Value::Text("ABC 123".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=ASC("ＡＢＣ　１２３")"#),
        Value::Text("ＡＢＣ　１２３".to_string())
    );
    assert_eq!(sheet.eval(r#"=ASC("ガ")"#), Value::Text("ガ".to_string()));
}

#[test]
fn asc_and_dbcs_convert_under_cp932() {
    let mut sheet = TestSheet::new();
    sheet.set_text_codepage(932);

    assert_eq!(
        sheet.eval(r#"=DBCS("ABC 123")"#),
        Value::Text("ＡＢＣ　１２３".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=ASC("ＡＢＣ　１２３")"#),
        Value::Text("ABC 123".to_string())
    );

    // Katakana + dakuten/handakuten composition.
    assert_eq!(
        sheet.eval(r#"=DBCS("ｶﾞｯﾁｮｲ")"#),
        Value::Text("ガッチョイ".to_string())
    );
    assert_eq!(sheet.eval(r#"=ASC("ガ")"#), Value::Text("ｶﾞ".to_string()));
    assert_eq!(sheet.eval(r#"=ASC("パ")"#), Value::Text("ﾊﾟ".to_string()));
    assert_eq!(sheet.eval(r#"=DBCS("ｳﾞ")"#), Value::Text("ヴ".to_string()));

    // Katakana punctuation + small kana should round-trip.
    assert_eq!(
        sheet.eval(r#"=ASC("。「」、・")"#),
        Value::Text("｡｢｣､･".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=DBCS("｡｢｣､･")"#),
        Value::Text("。「」、・".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=ASC("ァィゥェォャュョッー")"#),
        Value::Text("ｧｨｩｪｫｬｭｮｯｰ".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=DBCS("ｧｨｩｪｫｬｭｮｯｰ")"#),
        Value::Text("ァィゥェォャュョッー".to_string())
    );

    // Less common voiced katakana.
    assert_eq!(sheet.eval(r#"=ASC("ヷ")"#), Value::Text("ﾜﾞ".to_string()));
    assert_eq!(sheet.eval(r#"=DBCS("ﾜﾞ")"#), Value::Text("ヷ".to_string()));

    // Fullwidth compatibility symbols (U+FFE0..U+FFE6).
    assert_eq!(
        sheet.eval(r#"=DBCS("¢£¬¯¦¥₩")"#),
        Value::Text("￠￡￢￣￤￥￦".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=ASC("￠￡￢￣￤￥￦")"#),
        Value::Text("¢£¬¯¦¥₩".to_string())
    );

    // Decomposed voiced/semi-voiced katakana (base + combining mark) should match precomposed
    // behavior.
    assert_eq!(
        sheet.eval("=ASC(\"カ\u{3099}\")"),
        Value::Text("ｶﾞ".to_string())
    );
    assert_eq!(
        sheet.eval("=ASC(\"ハ\u{309A}\")"),
        Value::Text("ﾊﾟ".to_string())
    );
}

#[test]
fn byte_count_text_functions_use_dbcs_byte_semantics_under_cp932() {
    let mut sheet = TestSheet::new();
    sheet.set_text_codepage(932);

    // Byte-count functions should treat Hiragana/Kanji as 2 bytes under Shift_JIS.
    assert_eq!(sheet.eval(r#"=LENB("あ")"#), Value::Number(2.0));
    assert_eq!(sheet.eval(r#"=LENB("Aあ")"#), Value::Number(3.0));

    // LEFTB/RIGHTB/MIDB operate on byte counts, truncating at character boundaries.
    assert_eq!(
        sheet.eval(r#"=LEFTB("A漢B",2)"#),
        Value::Text("A".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=LEFTB("A漢B",3)"#),
        Value::Text("A漢".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=RIGHTB("A漢B",2)"#),
        Value::Text("B".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=RIGHTB("A漢B",3)"#),
        Value::Text("漢B".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=MIDB("A漢B",2,2)"#),
        Value::Text("漢".to_string())
    );

    // FINDB/SEARCHB return 1-indexed byte positions.
    assert_eq!(sheet.eval(r#"=FINDB("漢","A漢B")"#), Value::Number(2.0));
    assert_eq!(sheet.eval(r#"=SEARCHB("b","A漢B")"#), Value::Number(4.0));
    assert_eq!(sheet.eval(r#"=SEARCHB("漢*","A漢B")"#), Value::Number(2.0));

    // REPLACEB uses byte-based start/length.
    assert_eq!(
        sheet.eval(r#"=REPLACEB("A漢B",2,2,"Z")"#),
        Value::Text("AZB".to_string())
    );
}

#[test]
fn lenb_counts_bytes_for_chinese_korean_and_big5_dbcs_codepages() {
    let mut sheet = TestSheet::new();

    // Simplified Chinese (GBK / cp936).
    sheet.set_text_codepage(936);
    assert_eq!(sheet.eval(r#"=LENB("汉")"#), Value::Number(2.0));
    assert_eq!(sheet.eval(r#"=LENB("A汉")"#), Value::Number(3.0));

    // Korean (EUC-KR / cp949).
    sheet.set_text_codepage(949);
    assert_eq!(sheet.eval(r#"=LENB("가")"#), Value::Number(2.0));
    assert_eq!(sheet.eval(r#"=LENB("A가")"#), Value::Number(3.0));

    // Traditional Chinese (Big5 / cp950).
    sheet.set_text_codepage(950);
    assert_eq!(sheet.eval(r#"=LENB("漢")"#), Value::Number(2.0));
    assert_eq!(sheet.eval(r#"=LENB("A漢")"#), Value::Number(3.0));
}

#[test]
fn byte_count_text_functions_use_dbcs_semantics_for_other_dbcs_codepages() {
    let mut sheet = TestSheet::new();

    let cases = [
        (936, "汉"), // GBK (Simplified Chinese)
        (949, "가"), // EUC-KR (Korean)
        (950, "漢"), // Big5 (Traditional Chinese)
    ];

    for (cp, ch) in cases {
        sheet.set_text_codepage(cp);

        assert_eq!(sheet.eval(&format!(r#"=LENB("{ch}")"#)), Value::Number(2.0));
        assert_eq!(
            sheet.eval(&format!(r#"=LENB("A{ch}B")"#)),
            Value::Number(4.0)
        );

        assert_eq!(
            sheet.eval(&format!(r#"=LEFTB("A{ch}B",2)"#)),
            Value::Text("A".to_string()),
            "LEFTB should truncate partial DBCS characters (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=LEFTB("A{ch}B",3)"#)),
            Value::Text(format!("A{ch}")),
            "LEFTB should include full DBCS characters when length permits (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=RIGHTB("A{ch}B",2)"#)),
            Value::Text("B".to_string()),
            "RIGHTB should truncate partial DBCS characters (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=RIGHTB("A{ch}B",3)"#)),
            Value::Text(format!("{ch}B")),
            "RIGHTB should include full DBCS characters when length permits (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=MIDB("A{ch}B",2,2)"#)),
            Value::Text(ch.to_string()),
            "MIDB should slice on byte indices in DBCS locales (cp={cp})"
        );

        // FINDB/SEARCHB return 1-indexed byte positions.
        assert_eq!(
            sheet.eval(&format!(r#"=FINDB("{ch}","A{ch}B")"#)),
            Value::Number(2.0),
            "FINDB should return byte position (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=SEARCHB("b","A{ch}B")"#)),
            Value::Number(4.0),
            "SEARCHB should return byte position (cp={cp})"
        );
        assert_eq!(
            sheet.eval(&format!(r#"=SEARCHB("{ch}*","A{ch}B")"#)),
            Value::Number(2.0),
            "SEARCHB should support wildcards with byte positions (cp={cp})"
        );

        // REPLACEB uses byte-based start/length.
        assert_eq!(
            sheet.eval(&format!(r#"=REPLACEB("A{ch}B",2,2,"Z")"#)),
            Value::Text("AZB".to_string()),
            "REPLACEB should replace by bytes in DBCS locales (cp={cp})"
        );
    }
}

#[test]
fn phonetic_reads_cell_metadata_or_falls_back_to_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "abc");

    // Excel returns the cell's displayed text when no phonetic guide metadata is present.
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text("abc".to_string()));

    // When phonetic guides are present, return them.
    sheet.set_phonetic("A1", Some("あびし"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("あびし".to_string())
    );

    // Clearing phonetic guides should fall back again.
    sheet.set_phonetic("A1", None);
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text("abc".to_string()));

    sheet.set("A1", Value::Blank); // Blank input should remain blank.
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_spills_over_range_references() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "abc");
    sheet.set("A2", "def");
    sheet.set_phonetic("A1", Some("あびし"));

    assert_eq!(
        sheet.eval("=PHONETIC(A1:A2)"),
        Value::Text("あびし".to_string())
    );
    assert_eq!(sheet.get("Z2"), Value::Text("def".to_string()));
}

#[test]
fn phonetic_metadata_is_cleared_when_cell_input_changes() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    // Editing the cell value should clear stored phonetic metadata so PHONETIC never returns stale
    // furigana.
    sheet.set("A1", "東京");
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("東京".to_string())
    );

    // Editing the cell formula should also clear stored phonetic metadata.
    sheet.set("A1", "日本");
    sheet.set_phonetic("A1", Some("にほん"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("にほん".to_string())
    );

    sheet.set_formula("A1", "=\"大阪\"");
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("大阪".to_string())
    );
}

#[test]
fn phonetic_metadata_is_cleared_when_set_range_values_overwrites_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    let values = vec![vec![Value::Text("東京".to_string())]];
    sheet.set_range_values("A1", &values);
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("東京".to_string())
    );
}

#[test]
fn phonetic_metadata_is_cleared_when_copy_range_overwrites_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    sheet.set("B1", "東京");

    sheet.apply_operation(EditOp::CopyRange {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst_top_left: CellRef::from_a1("A1").expect("cell"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("東京".to_string())
    );
}

#[test]
fn phonetic_metadata_is_cleared_when_fill_overwrites_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    sheet.set("B1", "東京");

    sheet.apply_operation(EditOp::Fill {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst: Range::from_a1("A1").expect("range"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("東京".to_string())
    );
}

#[test]
fn phonetic_metadata_is_cleared_when_move_range_overwrites_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    sheet.set("B1", "東京");

    sheet.apply_operation(EditOp::MoveRange {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst_top_left: CellRef::from_a1("A1").expect("cell"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("東京".to_string())
    );
}

#[test]
fn phonetic_metadata_is_not_copied_by_copy_range() {
    let mut sheet = TestSheet::new();
    sheet.set("B1", "漢字");
    sheet.set_phonetic("B1", Some("かんじ"));

    sheet.apply_operation(EditOp::CopyRange {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst_top_left: CellRef::from_a1("A1").expect("cell"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("漢字".to_string())
    );
}

#[test]
fn phonetic_metadata_is_not_copied_by_fill() {
    let mut sheet = TestSheet::new();
    sheet.set("B1", "漢字");
    sheet.set_phonetic("B1", Some("かんじ"));

    sheet.apply_operation(EditOp::Fill {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst: Range::from_a1("A1").expect("range"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("漢字".to_string())
    );
}

#[test]
fn phonetic_metadata_moves_with_cell_on_move_range() {
    let mut sheet = TestSheet::new();
    sheet.set("B1", "漢字");
    sheet.set_phonetic("B1", Some("かんじ"));

    sheet.apply_operation(EditOp::MoveRange {
        sheet: "Sheet1".to_string(),
        src: Range::from_a1("B1").expect("range"),
        dst_top_left: CellRef::from_a1("A1").expect("cell"),
    });

    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );
    assert_eq!(sheet.eval("=PHONETIC(B1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_metadata_is_removed_when_clear_cell_deletes_record() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    sheet.clear_cell("A1");
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_metadata_is_removed_when_clear_range_deletes_record() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    sheet.clear_range("A1:A1");
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_metadata_is_removed_when_set_range_values_clears_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    let values = vec![vec![Value::Blank]];
    sheet.set_range_values("A1", &values);
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_metadata_is_removed_when_set_cell_value_clears_cell() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    sheet.set("A1", Value::Blank);
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_metadata_is_removed_when_set_cell_value_clears_contents_but_preserves_style() {
    use formula_engine::Engine;
    use formula_model::Style;

    let mut engine = Engine::new();
    let style_id = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();
    engine.set_cell_value("Sheet1", "A1", "漢字").unwrap();
    engine
        .set_cell_phonetic("Sheet1", "A1", Some("かんじ".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=PHONETIC(A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("かんじ".to_string())
    );

    // Clearing contents should preserve style, but must clear phonetic metadata.
    engine.set_cell_value("Sheet1", "A1", Value::Blank).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text(String::new()));
    assert_eq!(engine.get_cell_style_id("Sheet1", "A1").unwrap(), Some(style_id));
    assert_eq!(engine.get_cell_phonetic("Sheet1", "A1"), None);
}

#[test]
fn phonetic_metadata_is_retained_on_style_only_edits() {
    use formula_model::Style;

    let mut sheet = TestSheet::new();
    sheet.set("A1", "漢字");
    sheet.set_phonetic("A1", Some("かんじ"));
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );

    // Style-only edits must not clear phonetic metadata.
    let style_id = sheet.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    sheet.set_cell_style_id("A1", style_id);
    assert_eq!(
        sheet.eval("=PHONETIC(A1)"),
        Value::Text("かんじ".to_string())
    );
}

#[test]
fn phonetic_propagates_errors() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=1/0");
    sheet.recalculate();

    // Error values are propagated even if phonetic metadata is present.
    sheet.set_phonetic("A1", Some("should-not-win"));
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn phonetic_fallback_coerces_numbers_using_value_locale() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());
    sheet.set("A1", 1.5);
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text("1,5".to_string()));
}

#[test]
fn changing_text_codepage_marks_formulas_dirty() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", r#"=DBCS("ABC 123")"#);
    sheet.recalc();
    assert_eq!(sheet.get("A1"), Value::Text("ABC 123".to_string()));

    sheet.set_text_codepage(932);
    sheet.recalc();
    assert_eq!(
        sheet.get("A1"),
        Value::Text("ＡＢＣ　１２３".to_string())
    );
}
