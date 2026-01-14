#[path = "functions/harness.rs"]
mod harness;

use formula_engine::Value;

use harness::TestSheet;

fn assert_dbcs_byte_semantics_for_codepage(codepage: u16, dbcs_char: &str) {
    let mut sheet = TestSheet::new();
    sheet.set_text_codepage(codepage);

    let text = format!("A{dbcs_char}B");

    // Byte-length accounting.
    assert_eq!(
        sheet.eval(&format!("=LENB(\"{dbcs_char}\")")),
        Value::Number(2.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LENB(\"A{dbcs_char}\")")),
        Value::Number(3.0)
    );

    // Byte-based slicing should align to character boundaries (never return partial DBCS chars).
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{text}\",2)")),
        Value::Text("A".to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{text}\",3)")),
        Value::Text(format!("A{dbcs_char}"))
    );
    assert_eq!(
        sheet.eval(&format!("=RIGHTB(\"{text}\",2)")),
        Value::Text("B".to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=RIGHTB(\"{text}\",3)")),
        Value::Text(format!("{dbcs_char}B"))
    );
    assert_eq!(
        sheet.eval(&format!("=MIDB(\"{text}\",2,2)")),
        Value::Text(dbcs_char.to_string())
    );

    // FINDB/SEARCHB return 1-indexed byte positions.
    assert_eq!(
        sheet.eval(&format!("=FINDB(\"{dbcs_char}\",\"{text}\")")),
        Value::Number(2.0)
    );
    assert_eq!(
        sheet.eval(&format!("=SEARCHB(\"b\",\"{text}\")")),
        Value::Number(4.0)
    );

    // REPLACEB uses byte-based start/length.
    assert_eq!(
        sheet.eval(&format!("=REPLACEB(\"{text}\",2,2,\"Z\")")),
        Value::Text("AZB".to_string())
    );
}

#[test]
fn byte_count_text_functions_use_dbcs_byte_semantics_under_cp936() {
    // Simplified Chinese (GBK). Choose a character that is representable as a 2-byte sequence.
    assert_dbcs_byte_semantics_for_codepage(936, "汉");
}

#[test]
fn byte_count_text_functions_use_dbcs_byte_semantics_under_cp949() {
    // Korean (EUC-KR / codepage 949). Choose a Hangul syllable (2 bytes).
    assert_dbcs_byte_semantics_for_codepage(949, "가");
}

#[test]
fn byte_count_text_functions_use_dbcs_byte_semantics_under_cp950() {
    // Traditional Chinese (Big5). Choose a character that is representable as a 2-byte sequence.
    assert_dbcs_byte_semantics_for_codepage(950, "漢");
}

