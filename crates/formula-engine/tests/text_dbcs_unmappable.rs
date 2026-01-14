use formula_engine::{Engine, Value};

struct TestSheet {
    engine: Engine,
    sheet: &'static str,
    cell: &'static str,
}

impl TestSheet {
    fn new(codepage: u16) -> Self {
        let mut engine = Engine::new();
        engine.set_text_codepage(codepage);
        Self {
            engine,
            sheet: "Sheet1",
            cell: "A1",
        }
    }

    fn eval(&mut self, formula: &str) -> Value {
        self.engine
            .set_cell_formula(self.sheet, self.cell, formula)
            .unwrap();
        // Use the single-threaded recalc path in tests to avoid initializing a global Rayon pool.
        self.engine.recalculate_single_threaded();
        self.engine.get_cell_value(self.sheet, self.cell)
    }
}

#[test]
fn lenb_counts_unmappable_as_single_byte_cp932() {
    let mut sheet = TestSheet::new(932);
    let emoji = "\u{1F600}"; // 😀
    let hiragana_a = "\u{3042}"; // あ

    assert_eq!(
        sheet.eval(&format!("=LENB(\"{emoji}\")")),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LENB(\"A{emoji}{hiragana_a}\")")),
        Value::Number(4.0)
    );
}

#[test]
fn leftb_respects_unmappable_byte_count_cp932() {
    let mut sheet = TestSheet::new(932);
    let emoji = "\u{1F600}"; // 😀
    let hiragana_a = "\u{3042}"; // あ

    let text = format!("{emoji}{hiragana_a}");

    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{text}\",1)")),
        Value::Text(emoji.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{text}\",2)")),
        Value::Text(emoji.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{text}\",3)")),
        Value::Text(text)
    );
}

#[test]
fn findb_positions_include_unmappable_bytes_cp932() {
    let mut sheet = TestSheet::new(932);
    let emoji = "\u{1F600}"; // 😀
    let hiragana_a = "\u{3042}"; // あ

    let text = format!("{emoji}{hiragana_a}");

    assert_eq!(
        sheet.eval(&format!("=FINDB(\"{hiragana_a}\",\"{text}\")")),
        Value::Number(2.0)
    );
    assert_eq!(
        sheet.eval(&format!("=FINDB(\"{emoji}\",\"{text}\")")),
        Value::Number(1.0)
    );
}

#[test]
fn midb_rightb_searchb_replaceb_respect_unmappable_byte_count_cp932() {
    let mut sheet = TestSheet::new(932);
    let emoji = "\u{1F600}"; // 😀
    let hiragana_a = "\u{3042}"; // あ

    // Byte counts under cp932 for this test data:
    // 😀 = 1 byte (unmappable -> replacement byte)
    // あ = 2 bytes
    // Total for "😀あ" = 3 bytes.
    let text = format!("{emoji}{hiragana_a}");

    // RIGHTB is byte-based and truncates at character boundaries.
    assert_eq!(
        sheet.eval(&format!("=RIGHTB(\"{text}\",1)")),
        Value::Text(String::new())
    );
    assert_eq!(
        sheet.eval(&format!("=RIGHTB(\"{text}\",2)")),
        Value::Text(hiragana_a.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=RIGHTB(\"{text}\",3)")),
        Value::Text(text.clone())
    );

    // MIDB uses byte-based start/length (1-indexed bytes), truncating at character boundaries.
    let longer = format!("A{emoji}{hiragana_a}");
    assert_eq!(
        sheet.eval(&format!("=MIDB(\"{longer}\",2,1)")),
        Value::Text(emoji.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=MIDB(\"{longer}\",3,1)")),
        Value::Text(String::new())
    );
    assert_eq!(
        sheet.eval(&format!("=MIDB(\"{longer}\",3,2)")),
        Value::Text(hiragana_a.to_string())
    );

    // SEARCHB byte positions include the unmappable replacement byte.
    assert_eq!(
        sheet.eval(&format!("=SEARCHB(\"{hiragana_a}\",\"{text}\")")),
        Value::Number(2.0)
    );
    assert_eq!(
        sheet.eval(&format!("=SEARCHB(\"{emoji}\",\"{text}\")")),
        Value::Number(1.0)
    );

    // REPLACEB start/len are byte-based.
    assert_eq!(
        sheet.eval(&format!("=REPLACEB(\"{text}\",1,1,\"X\")")),
        Value::Text(format!("X{hiragana_a}"))
    );
    assert_eq!(
        sheet.eval(&format!("=REPLACEB(\"{text}\",2,2,\"X\")")),
        Value::Text(format!("{emoji}X"))
    );
}

#[test]
fn unmappable_semantics_are_consistent_across_dbcs_codepages_cp936() {
    let mut sheet = TestSheet::new(936);
    let emoji = "\u{1F600}"; // 😀
    let han = "\u{4E2D}"; // 中

    assert_eq!(
        sheet.eval(&format!("=LENB(\"{emoji}\")")),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LENB(\"A{emoji}{han}\")")),
        Value::Number(4.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{emoji}{han}\",2)")),
        Value::Text(emoji.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=FINDB(\"{han}\",\"{emoji}{han}\")")),
        Value::Number(2.0)
    );
}

#[test]
fn unmappable_semantics_are_consistent_across_dbcs_codepages_cp949() {
    let mut sheet = TestSheet::new(949);
    let emoji = "\u{1F600}"; // 😀
    let hangul = "\u{D55C}"; // 한

    assert_eq!(
        sheet.eval(&format!("=LENB(\"{emoji}\")")),
        Value::Number(1.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LENB(\"A{emoji}{hangul}\")")),
        Value::Number(4.0)
    );
    assert_eq!(
        sheet.eval(&format!("=LEFTB(\"{emoji}{hangul}\",2)")),
        Value::Text(emoji.to_string())
    );
    assert_eq!(
        sheet.eval(&format!("=FINDB(\"{hangul}\",\"{emoji}{hangul}\")")),
        Value::Number(2.0)
    );
}
