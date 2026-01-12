use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn byte_count_text_functions_match_non_b_variants_in_en_us() {
    let mut sheet = TestSheet::new();

    // In single-byte locales (en-US), the `*B` functions are expected to behave
    // identically to their non-`B` equivalents.
    assert_eq!(sheet.eval(r#"=LENB("abc")"#), sheet.eval(r#"=LEN("abc")"#));
    assert_eq!(sheet.eval(r#"=LEFTB("abc")"#), sheet.eval(r#"=LEFT("abc")"#));
    assert_eq!(
        sheet.eval(r#"=LEFTB("abc",2)"#),
        sheet.eval(r#"=LEFT("abc",2)"#)
    );
    assert_eq!(sheet.eval(r#"=RIGHTB("abc")"#), sheet.eval(r#"=RIGHT("abc")"#));
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
fn asc_and_dbcs_are_identity_transforms_for_now() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval(r#"=ASC("ABC")"#), Value::Text("ABC".to_string()));
    assert_eq!(sheet.eval(r#"=DBCS("ABC")"#), Value::Text("ABC".to_string()));

    // Non-ASCII smoke test: the engine currently does not implement the locale-specific
    // half-width/full-width conversions that Excel performs in some DBCS locales.
    assert_eq!(sheet.eval(r#"=ASC("漢")"#), Value::Text("漢".to_string()));
    assert_eq!(sheet.eval(r#"=DBCS("漢")"#), Value::Text("漢".to_string()));
}

#[test]
fn phonetic_is_a_deterministic_placeholder() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "abc");

    // Real Excel reads a phonetic guide from cell metadata. We don't model that yet,
    // so we currently return the referenced cell's text.
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text("abc".to_string()));

    // Blank input should remain blank.
    sheet.set("A1", Value::Blank);
    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Text(String::new()));
}

#[test]
fn phonetic_propagates_errors() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=1/0");
    sheet.recalculate();

    assert_eq!(sheet.eval("=PHONETIC(A1)"), Value::Error(ErrorKind::Div0));
}

