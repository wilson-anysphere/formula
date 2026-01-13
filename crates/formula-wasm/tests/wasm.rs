#![cfg(target_arch = "wasm32")]

use serde_json::json;
use serde_json::Value as JsonValue;
use js_sys::{Object, Reflect};
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_model::CellValue as ModelCellValue;
use formula_wasm::{
    canonicalize_formula, lex_formula, localize_formula, parse_formula_partial,
    rewrite_formulas_for_copy_delta, WasmWorkbook, DEFAULT_SHEET,
};

#[derive(Debug, serde::Deserialize, PartialEq, Eq)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct LexToken {
    kind: String,
    span: Span,
    #[serde(default)]
    value: Option<JsonValue>,
}

#[derive(Debug, serde::Deserialize)]
struct TokenKindOnly {
    kind: String,
}

#[derive(Debug, serde::Deserialize)]
struct LexError {
    message: String,
    span: Span,
}

#[derive(Debug, serde::Deserialize)]
struct PartialLexResult {
    tokens: Vec<LexToken>,
    error: Option<LexError>,
}

fn assert_json_number(value: &JsonValue, expected: f64) {
    let actual = value
        .as_f64()
        .unwrap_or_else(|| panic!("expected JSON number, got {value:?}"));
    assert_eq!(actual, expected);
}

fn to_js_value<T: serde::Serialize>(value: &T) -> JsValue {
    value
        .serialize(&serde_wasm_bindgen::Serializer::json_compatible())
        .expect("failed to serialize JS value")
}

#[derive(Debug, serde::Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct PartialParseResult {
    error: Option<PartialParseError>,
    context: PartialParseContext,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct PartialParseError {
    message: String,
    span: Span,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct PartialParseContext {
    function: Option<PartialParseFunctionContext>,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct PartialParseFunctionContext {
    name: String,
    arg_index: usize,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct CellChange {
    sheet: String,
    address: String,
    value: JsonValue,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct CellData {
    sheet: String,
    address: String,
    input: JsonValue,
    value: JsonValue,
}

#[wasm_bindgen_test]
fn debug_function_registry_contains_builtins() {
    // Ensure the wasm module invoked Rust global constructors before touching the
    // function registry (otherwise it can be cached as empty).
    let _ = WasmWorkbook::new();
    assert!(formula_engine::functions::lookup_function("SUM").is_some());
    assert!(formula_engine::functions::lookup_function("SEQUENCE").is_some());
}

#[wasm_bindgen_test]
fn parse_formula_partial_fallback_context_in_unterminated_string() {
    let formula = r#"=SUM("hello"#.to_string();
    let cursor = formula.encode_utf16().count() as u32;

    let parsed_js = parse_formula_partial(formula, Some(cursor), None).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    assert!(parsed.error.is_some());
    assert_eq!(
        parsed.error.unwrap().message,
        "Unterminated string literal".to_string()
    );

    let ctx = parsed.context.function.unwrap();
    assert_eq!(ctx.name, "SUM".to_string());
    assert_eq!(ctx.arg_index, 0);
}

#[wasm_bindgen_test]
fn parse_formula_partial_uses_utf16_cursor_and_spans_for_emoji() {
    let formula = r#"=SUM("😀"#.to_string();
    let cursor = formula.encode_utf16().count() as u32;

    let parsed_js = parse_formula_partial(formula, Some(cursor), None).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    let err = parsed.error.expect("expected error");
    assert_eq!(err.message, "Unterminated string literal".to_string());
    // Span offsets must be expressed as UTF-16 code units (JS string indices).
    assert_eq!(err.span.start, 5);
    assert_eq!(err.span.end, cursor as usize);

    let ctx = parsed.context.function.expect("expected function context");
    assert_eq!(ctx.name, "SUM".to_string());
    assert_eq!(ctx.arg_index, 0);
}

#[wasm_bindgen_test]
fn parse_formula_partial_fallback_context_in_unterminated_sheet_quote() {
    let formula = "=SUM('My Sheet".to_string();
    let cursor = formula.encode_utf16().count() as u32;

    let parsed_js = parse_formula_partial(formula, Some(cursor), None).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    assert!(parsed.error.is_some());
    assert_eq!(
        parsed.error.unwrap().message,
        "Unterminated quoted identifier".to_string()
    );

    let ctx = parsed.context.function.unwrap();
    assert_eq!(ctx.name, "SUM".to_string());
    assert_eq!(ctx.arg_index, 0);
}

#[wasm_bindgen_test]
fn lex_formula_emits_utf16_spans_for_emoji() {
    let formula = "=\"😀\"".to_string();
    let expected_end = formula.encode_utf16().count();

    let tokens_js = lex_formula(&formula, None).unwrap();
    let tokens: Vec<LexToken> = serde_wasm_bindgen::from_value(tokens_js).unwrap();

    let string_token = tokens
        .iter()
        .find(|t| t.kind == "String")
        .expect("expected string token");
    assert_eq!(string_token.span.start, 1);
    assert_eq!(string_token.span.end, expected_end);
}

#[wasm_bindgen_test]
fn lex_formula_honors_locale_id_option_for_arg_separator() {
    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("localeId"), &JsValue::from_str("de-DE")).unwrap();

    // In the default en-US locale, `;` is not accepted as a function argument separator.
    // In de-DE, `;` *is* the argument separator.
    let tokens_js = lex_formula("=SUMME(1;2)", Some(opts.into())).unwrap();
    let tokens: Vec<TokenKindOnly> = serde_wasm_bindgen::from_value(tokens_js).unwrap();
    assert!(
        tokens.iter().any(|t| t.kind == "ArgSep"),
        "expected de-DE lexing to emit ArgSep tokens, got {tokens:?}"
    );

    let default_err = lex_formula("=SUMME(1;2)", None).unwrap_err();
    let message = default_err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {default_err:?}"));
    assert!(
        message.contains("Unexpected character `;`"),
        "expected default en-US lexing to reject ';' as a function arg separator, got {message:?}"
    );
}

#[wasm_bindgen_test]
fn lex_formula_rejects_unknown_locale_id_option() {
    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("localeId"), &JsValue::from_str("xx-XX")).unwrap();

    let err = lex_formula("=1+2", Some(opts.into())).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "unknown localeId: xx-XX");
}

#[wasm_bindgen_test]
fn lex_formula_rejects_non_object_options() {
    let err = lex_formula("=1+2", Some(JsValue::from_f64(1.0))).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "options must be an object");
}

#[wasm_bindgen_test]
fn lex_formula_rejects_unrecognized_options_object() {
    let opts = Object::new();
    // Common mistake: wrong casing on localeId.
    Reflect::set(&opts, &JsValue::from_str("localeID"), &JsValue::from_str("de-DE")).unwrap();

    let err = lex_formula("=1+2", Some(opts.into())).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(
        message,
        "options must be { localeId?: string, referenceStyle?: \"A1\" | \"R1C1\" } or a ParseOptions object"
    );
}

#[wasm_bindgen_test]
fn lex_formula_honors_reference_style_option() {
    let opts = Object::new();
    Reflect::set(
        &opts,
        &JsValue::from_str("referenceStyle"),
        &JsValue::from_str("R1C1"),
    )
    .unwrap();

    let tokens_js = lex_formula("=R1C1", Some(opts.into())).unwrap();
    let tokens: Vec<TokenKindOnly> = serde_wasm_bindgen::from_value(tokens_js).unwrap();
    assert!(
        tokens.iter().any(|t| t.kind == "R1C1Cell"),
        "expected R1C1 lexing to emit R1C1Cell tokens, got {tokens:?}"
    );

    let default_js = lex_formula("=R1C1", None).unwrap();
    let default_tokens: Vec<TokenKindOnly> = serde_wasm_bindgen::from_value(default_js).unwrap();
    assert!(
        default_tokens.iter().all(|t| t.kind != "R1C1Cell"),
        "expected default A1 lexing to NOT emit R1C1Cell tokens, got {default_tokens:?}"
    );
}

#[wasm_bindgen_test]
fn lex_formula_accepts_full_parse_options_struct() {
    // Backward compatibility: `lexFormula` should still accept a fully-serialized
    // `formula_engine::ParseOptions` object (snake_case), not just the JS-friendly DTO.
    let opts = formula_engine::ParseOptions {
        locale: formula_engine::LocaleConfig::en_us(),
        reference_style: formula_engine::ReferenceStyle::R1C1,
        normalize_relative_to: None,
    };
    let opts_js = serde_wasm_bindgen::to_value(&opts).unwrap();

    let tokens_js = lex_formula("=R1C1", Some(opts_js)).unwrap();
    let tokens: Vec<TokenKindOnly> = serde_wasm_bindgen::from_value(tokens_js).unwrap();
    assert!(
        tokens.iter().any(|t| t.kind == "R1C1Cell"),
        "expected full ParseOptions R1C1 lexing to emit R1C1Cell tokens, got {tokens:?}"
    );
}

#[wasm_bindgen_test]
fn canonicalize_and_localize_formula_roundtrip_de_de() {
    let localized = "=SUMME(1,5;2)";
    let canonical = canonicalize_formula(localized, "de-DE", None).unwrap();
    assert_eq!(canonical, "=SUM(1.5,2)");

    let roundtrip = localize_formula(&canonical, "de-DE", None).unwrap();
    assert_eq!(roundtrip, localized);
}

#[wasm_bindgen_test]
fn canonicalize_and_localize_formula_roundtrip_fr_fr() {
    let localized = "=SOMME(1,5;2)";
    let canonical = canonicalize_formula(localized, "fr-FR", None).unwrap();
    assert_eq!(canonical, "=SUM(1.5,2)");

    let roundtrip = localize_formula(&canonical, "fr-FR", None).unwrap();
    assert_eq!(roundtrip, localized);
}

#[wasm_bindgen_test]
fn canonicalize_and_localize_formula_roundtrip_r1c1_reference_style() {
    let localized = "=SUMME(R1C1;R1C2)";
    let canonical = canonicalize_formula(localized, "de-DE", Some("R1C1".to_string())).unwrap();
    assert_eq!(canonical, "=SUM(R1C1,R1C2)");

    let roundtrip = localize_formula(&canonical, "de-DE", Some("R1C1".to_string())).unwrap();
    assert_eq!(roundtrip, localized);
}

#[wasm_bindgen_test]
fn canonicalize_formula_rejects_unknown_locale_id_with_supported_list() {
    let err = canonicalize_formula("=1+2", "xx-XX", None).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert!(message.contains("unknown localeId: xx-XX"));
    assert!(
        message.contains("Supported locale ids"),
        "expected actionable locale message, got {message:?}"
    );
}

#[wasm_bindgen_test]
fn rewrite_formulas_for_copy_delta_shifts_a1_references() {
    let requests = vec![json!({
        "formula": "=A2",
        "deltaRow": 1,
        "deltaCol": 1,
    })];
    let requests_js = to_js_value(&requests);
    let out_js = rewrite_formulas_for_copy_delta(requests_js).unwrap();
    let out: Vec<String> = serde_wasm_bindgen::from_value(out_js).unwrap();
    assert_eq!(out, vec!["=B3".to_string()]);
}

#[wasm_bindgen_test]
fn rewrite_formulas_for_copy_delta_shifts_row_and_column_ranges() {
    let requests = vec![
        json!({
            "formula": "=SUM(A:A)",
            "deltaRow": 0,
            "deltaCol": 1,
        }),
        json!({
            "formula": "=SUM(1:1)",
            "deltaRow": 1,
            "deltaCol": 0,
        }),
    ];
    let requests_js = to_js_value(&requests);
    let out_js = rewrite_formulas_for_copy_delta(requests_js).unwrap();
    let out: Vec<String> = serde_wasm_bindgen::from_value(out_js).unwrap();
    assert_eq!(out, vec!["=SUM(B:B)".to_string(), "=SUM(2:2)".to_string()]);
}

#[wasm_bindgen_test]
fn rewrite_formulas_for_copy_delta_drops_spill_postfix_when_reference_becomes_ref_error() {
    let requests = vec![json!({
        "formula": "=A1#",
        "deltaRow": 0,
        "deltaCol": -1,
    })];
    let requests_js = to_js_value(&requests);
    let out_js = rewrite_formulas_for_copy_delta(requests_js).unwrap();
    let out: Vec<String> = serde_wasm_bindgen::from_value(out_js).unwrap();
    assert_eq!(out, vec!["=#REF!".to_string()]);
}

#[wasm_bindgen_test]
fn parse_formula_partial_honors_locale_id_option() {
    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("localeId"), &JsValue::from_str("de-DE")).unwrap();

    let parsed_js =
        parse_formula_partial("=SUMME(1;2)".to_string(), None, Some(opts.into())).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    assert!(parsed.error.is_none(), "expected successful parse, got {parsed:?}");

    let default_js = parse_formula_partial("=SUMME(1;2)".to_string(), None, None).unwrap();
    let default_parsed: PartialParseResult = serde_wasm_bindgen::from_value(default_js).unwrap();
    let default_err = default_parsed.error.expect("expected parse error");
    assert_eq!(default_err.message, "Unexpected character `;`".to_string());
    // Span should point at the `;` in `=SUMME(1;2)`.
    assert_eq!(default_err.span.start, 8);
    assert_eq!(default_err.span.end, 9);
}

#[wasm_bindgen_test]
fn parse_formula_partial_rejects_unknown_locale_id_option() {
    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("localeId"), &JsValue::from_str("xx-XX")).unwrap();

    let err = parse_formula_partial("=1+2".to_string(), None, Some(opts.into())).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "unknown localeId: xx-XX");
}

#[wasm_bindgen_test]
fn parse_formula_partial_rejects_non_object_options() {
    let err = parse_formula_partial("=1+2".to_string(), None, Some(JsValue::from_f64(1.0))).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "options must be an object");
}

#[wasm_bindgen_test]
fn parse_formula_partial_rejects_unrecognized_options_object() {
    let opts = Object::new();
    Reflect::set(&opts, &JsValue::from_str("localeID"), &JsValue::from_str("de-DE")).unwrap();

    let err = parse_formula_partial("=1+2".to_string(), None, Some(opts.into())).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(
        message,
        "options must be { localeId?: string, referenceStyle?: \"A1\" | \"R1C1\" } or a ParseOptions object"
    );
}

#[wasm_bindgen_test]
fn parse_formula_partial_honors_reference_style_option() {
    #[derive(Debug, serde::Deserialize)]
    struct PartialParseWithAst {
        ast: formula_engine::Ast,
        error: Option<PartialParseError>,
    }

    let opts = Object::new();
    Reflect::set(
        &opts,
        &JsValue::from_str("referenceStyle"),
        &JsValue::from_str("R1C1"),
    )
    .unwrap();

    let parsed_js = parse_formula_partial("=R1C1".to_string(), None, Some(opts.into())).unwrap();
    let parsed: PartialParseWithAst = serde_wasm_bindgen::from_value(parsed_js).unwrap();
    assert!(parsed.error.is_none(), "expected successful parse, got {parsed:?}");
    assert!(
        matches!(parsed.ast.expr, formula_engine::Expr::CellRef(_)),
        "expected R1C1 parse to yield CellRef AST node, got {:?}",
        parsed.ast.expr
    );

    // Default reference style is A1, so `R1C1` should be treated as an identifier/name rather than
    // an R1C1 cell reference.
    let default_js = parse_formula_partial("=R1C1".to_string(), None, None).unwrap();
    let default_parsed: PartialParseWithAst = serde_wasm_bindgen::from_value(default_js).unwrap();
    assert!(
        default_parsed.error.is_none(),
        "expected successful parse, got {default_parsed:?}"
    );
    assert!(
        matches!(default_parsed.ast.expr, formula_engine::Expr::NameRef(_)),
        "expected default A1 parse to yield NameRef, got {:?}",
        default_parsed.ast.expr
    );
}

#[wasm_bindgen_test]
fn parse_formula_partial_accepts_full_parse_options_struct() {
    #[derive(Debug, serde::Deserialize)]
    struct PartialParseWithAst {
        ast: formula_engine::Ast,
        error: Option<PartialParseError>,
    }

    // Backward compatibility: `parseFormulaPartial` should still accept a fully-serialized
    // `formula_engine::ParseOptions` object (snake_case), not just the JS-friendly DTO.
    let opts = formula_engine::ParseOptions {
        locale: formula_engine::LocaleConfig::en_us(),
        reference_style: formula_engine::ReferenceStyle::R1C1,
        normalize_relative_to: None,
    };
    let opts_js = serde_wasm_bindgen::to_value(&opts).unwrap();

    let parsed_js = parse_formula_partial("=R1C1".to_string(), None, Some(opts_js)).unwrap();
    let parsed: PartialParseWithAst = serde_wasm_bindgen::from_value(parsed_js).unwrap();
    assert!(parsed.error.is_none(), "expected successful parse, got {parsed:?}");
    assert!(
        matches!(parsed.ast.expr, formula_engine::Expr::CellRef(_)),
        "expected full ParseOptions R1C1 parse to yield CellRef, got {:?}",
        parsed.ast.expr
    );
}

#[wasm_bindgen_test]
fn lex_formula_partial_returns_tokens_and_error_for_unterminated_string() {
    let js = formula_wasm::lex_formula_partial("=\"hello", None);
    let parsed: PartialLexResult = serde_wasm_bindgen::from_value(js).unwrap();

    assert!(!parsed.tokens.is_empty(), "expected at least one token");
    let err = parsed.error.expect("expected an error");
    assert_eq!(err.message, "Unterminated string literal".to_string());
    assert_eq!(err.span.start, 1);
    assert_eq!(err.span.end, 7);

    // The string token should span to end-of-input, offset by the leading '='.
    let string_token = parsed
        .tokens
        .iter()
        .find(|t| t.kind == "String")
        .expect("expected string token");
    assert_eq!(string_token.span.start, 1);
    assert_eq!(string_token.span.end, 7);

    // The error span should also cover the unterminated string literal.
    assert_eq!(err.span.start, string_token.span.start);
    assert_eq!(err.span.end, string_token.span.end);
}

#[wasm_bindgen_test]
fn recalculate_reports_changed_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 2.0);

    let cell_js = wb.get_cell("A2".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 2.0);
}

#[wasm_bindgen_test]
fn recalculate_returns_empty_when_no_cells_changed() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());
}

#[wasm_bindgen_test]
fn recalculate_reports_lambda_values_as_calc_error() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=LAMBDA(x,x)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_eq!(changes[0].value, JsonValue::String("#CALC!".to_string()));

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::String("#CALC!".to_string()));
}

#[wasm_bindgen_test]
fn recalculate_reports_dynamic_array_spills() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,2)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);

    // Spill outputs should not be treated as explicit inputs in the workbook JSON.
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert!(b1.input.is_null());
    assert_json_number(&b1.value, 2.0);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(cells.contains_key("A1"));
    assert!(!cells.contains_key("B1"));
}

#[wasm_bindgen_test]
fn recalculate_reports_spill_resize_clears_trailing_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,3)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    // Shrink the spill width from 3 -> 2; `C1` should be cleared and surfaced as a recalc delta.
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,2)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 3);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);
    assert_eq!(changes[2].sheet, DEFAULT_SHEET);
    assert_eq!(changes[2].address, "C1");
    assert!(changes[2].value.is_null());
}

#[wasm_bindgen_test]
fn recalculate_orders_changes_by_sheet_row_col() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(10.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();
    wb.set_cell(
        "A2".to_string(),
        JsValue::from_str("=A1*2"),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    // Establish initial formula values.
    wb.recalculate(None).unwrap();

    // Dirty both sheets before a single recalc tick so ordering is deterministic.
    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(11.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, "Sheet1");
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 4.0);
    assert_eq!(changes[1].sheet, "Sheet2");
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 22.0);
}

#[wasm_bindgen_test]
fn recalculate_orders_changes_by_row_col_within_sheet() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1+1"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    // Row-major: B1 (row 0, col 1) comes before A2 (row 1, col 0).
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "B1");
    assert_json_number(&changes[0].value, 3.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 4.0);
}

#[wasm_bindgen_test]
fn recalculate_does_not_filter_changes_by_sheet_argument() {
    let mut wb = WasmWorkbook::new();

    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(10.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();
    wb.set_cell(
        "A2".to_string(),
        JsValue::from_str("=A1*2"),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    // Dirty both sheets, then request a sheet-scoped recalc. The returned changes should still
    // include all sheets so the JS cache remains coherent across sheet switches.
    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(11.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    let changes_js = wb.recalculate(Some("sHeEt1".to_string())).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, "Sheet1");
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 4.0);
    assert_eq!(changes[1].sheet, "Sheet2");
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 22.0);
}

#[wasm_bindgen_test]
fn recalculate_ignores_unknown_sheet_argument() {
    let mut wb = WasmWorkbook::new();
    let changes_js = wb.recalculate(Some("MissingSheet".to_string())).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());
}

#[wasm_bindgen_test]
fn recalculate_reports_formula_edit_to_blank_value() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=1"), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_str("=A2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![CellChange {
            sheet: DEFAULT_SHEET.to_string(),
            address: "A1".to_string(),
            value: JsonValue::Null,
        }]
    );
}

#[wasm_bindgen_test]
fn recalculate_reports_cleared_spill_outputs_after_edit() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,3)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 3);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);
    assert_eq!(changes[2].sheet, DEFAULT_SHEET);
    assert_eq!(changes[2].address, "C1");
    assert_json_number(&changes[2].value, 3.0);

    // Overwrite a spill output cell with a literal value. This clears the spill footprint before
    // the next recalc, so `recalculate()` must still report the remaining spill outputs as blank.
    wb.set_cell("B1".to_string(), JsValue::from_f64(99.0), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!("#SPILL!"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                value: JsonValue::Null,
            },
        ]
    );
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_formulas_and_recalculates() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/formulas/formulas.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_json_number(&cell.value, 3.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_preserves_stale_formula_cache_until_recalc() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/formulas/formulas-stale-cache.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Before recalc, `getCell` should expose the cached value from the XLSX file.
    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_json_number(&cell.value, 999.0);

    // A manual recalc should replace the stale cached value with the computed one.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "C1");
    assert_json_number(&changes[0].value, 3.0);

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_json_number(&cell.value, 3.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_basic_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Should not error even though the fixture contains no formulas.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_json_number(&a1.input, 1.0);
    assert_json_number(&a1.value, 1.0);

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("Hello"));
    assert_eq!(b1.value, json!("Hello"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_bool_and_error_cells() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/bool-error.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // No formulas → no recalculation deltas.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!(true));
    assert_eq!(a1.value, json!(true));

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!(false));
    assert_eq!(a2.value, json!(false));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("#DIV/0!"));
    assert_eq!(b1.value, json!("#DIV/0!"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_extended_error_cells_with_semantics() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/extended-errors.xlsx"
    ));
    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    let cases = [
        ("A1", "#GETTING_DATA", 8.0),
        ("B1", "#FIELD!", 11.0),
        ("C1", "#CONNECT!", 12.0),
        ("D1", "#BLOCKED!", 13.0),
        ("E1", "#UNKNOWN!", 14.0),
    ];

    for (address, expected_error, _) in cases.iter().copied() {
        let cell_js = wb.get_cell(address.to_string(), None).unwrap();
        let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
        assert_eq!(cell.value, json!(expected_error));
        assert_eq!(cell.input, json!(expected_error));
    }

    // Populate formulas that distinguish between real error values vs. error-looking strings.
    for (address, _, _) in cases.iter().copied() {
        let col = &address[0..1];

        wb.set_cell(
            format!("{col}2"),
            JsValue::from_str(&format!("=ISERROR({address})")),
            None,
        )
        .unwrap();
        wb.set_cell(
            format!("{col}3"),
            JsValue::from_str(&format!("=ERROR.TYPE({address})")),
            None,
        )
        .unwrap();
        wb.set_cell(
            format!("{col}4"),
            JsValue::from_str(&format!("={address}+1")),
            None,
        )
        .unwrap();
    }

    wb.recalculate(None).unwrap();

    for (address, expected_error, expected_code) in cases.iter().copied() {
        let col = &address[0..1];

        let iserror_js = wb.get_cell(format!("{col}2"), None).unwrap();
        let iserror: CellData = serde_wasm_bindgen::from_value(iserror_js).unwrap();
        assert_eq!(iserror.value, json!(true));

        let type_js = wb.get_cell(format!("{col}3"), None).unwrap();
        let type_cell: CellData = serde_wasm_bindgen::from_value(type_js).unwrap();
        assert_json_number(&type_cell.value, expected_code);

        let arith_js = wb.get_cell(format!("{col}4"), None).unwrap();
        let arith: CellData = serde_wasm_bindgen::from_value(arith_js).unwrap();
        assert_eq!(arith.value, json!(expected_error));
    }
}

#[wasm_bindgen_test]
fn getting_data_error_literal_is_parsed_as_error() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("#GETTING_DATA"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("  #getting_data  "), None)
        .unwrap();

    // Use ISERROR to distinguish between real errors and error-looking text.
    wb.set_cell("B1".to_string(), JsValue::from_str("=ISERROR(A1)"), None)
        .unwrap();
    wb.set_cell("B2".to_string(), JsValue::from_str("=ISERROR(A2)"), None)
        .unwrap();

    wb.set_cell("C1".to_string(), JsValue::from_str("=ERROR.TYPE(A1)"), None)
        .unwrap();
    wb.set_cell("C2".to_string(), JsValue::from_str("=ERROR.TYPE(A2)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, json!(true));

    let b2_js = wb.get_cell("B2".to_string(), None).unwrap();
    let b2: CellData = serde_wasm_bindgen::from_value(b2_js).unwrap();
    assert_eq!(b2.value, json!(true));

    let c1_js = wb.get_cell("C1".to_string(), None).unwrap();
    let c1: CellData = serde_wasm_bindgen::from_value(c1_js).unwrap();
    assert_json_number(&c1.value, 8.0);

    let c2_js = wb.get_cell("C2".to_string(), None).unwrap();
    let c2: CellData = serde_wasm_bindgen::from_value(c2_js).unwrap();
    assert_json_number(&c2.value, 8.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_preserves_modern_error_values_as_errors() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/bool-error.xlsx"
    ));
    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Use ISERROR to distinguish between real errors and error-looking text.
    wb.set_cell("C2".to_string(), JsValue::from_str("=ISERROR(C1)"), None)
        .unwrap();
    wb.set_cell("D2".to_string(), JsValue::from_str("=ISERROR(D1)"), None)
        .unwrap();
    wb.set_cell("E2".to_string(), JsValue::from_str("=ISERROR(E1)"), None)
        .unwrap();
    wb.set_cell("F2".to_string(), JsValue::from_str("=ISERROR(F1)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    for addr in ["C2", "D2", "E2", "F2"] {
        let cell_js = wb.get_cell(addr.to_string(), None).unwrap();
        let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
        assert_eq!(
            cell.value,
            json!(true),
            "expected {addr} to evaluate to TRUE"
        );
    }
}

#[wasm_bindgen_test]
fn leading_apostrophe_forces_text_for_error_literals() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("#DIV/0!"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("'#DIV/0!"), None)
        .unwrap();
    wb.set_cell("A3".to_string(), JsValue::from_str("'#GETTING_DATA"), None)
        .unwrap();

    // Use ISERROR to distinguish between real errors and error-looking text.
    wb.set_cell("B1".to_string(), JsValue::from_str("=ISERROR(A1)"), None)
        .unwrap();
    wb.set_cell("B2".to_string(), JsValue::from_str("=ISERROR(A2)"), None)
        .unwrap();
    wb.set_cell("B3".to_string(), JsValue::from_str("=ISERROR(A3)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, json!(true));

    let b2_js = wb.get_cell("B2".to_string(), None).unwrap();
    let b2: CellData = serde_wasm_bindgen::from_value(b2_js).unwrap();
    assert_eq!(b2.value, json!(false));

    let b3_js = wb.get_cell("B3".to_string(), None).unwrap();
    let b3: CellData = serde_wasm_bindgen::from_value(b3_js).unwrap();
    assert_eq!(b3.value, json!(false));

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!("'#DIV/0!"));
    assert_eq!(a2.value, json!("#DIV/0!"));

    let a3_js = wb.get_cell("A3".to_string(), None).unwrap();
    let a3: CellData = serde_wasm_bindgen::from_value(a3_js).unwrap();
    assert_eq!(a3.input, json!("'#GETTING_DATA"));
    assert_eq!(a3.value, json!("#GETTING_DATA"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_extended_error_cells_as_errors() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/extended-errors.xlsx"
    ));
    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Verify the imported cached values are surfaced as error literals (not plain text).
    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!("#GETTING_DATA"));
    assert_eq!(a1.value, json!("#GETTING_DATA"));

    // Formulas should treat the value as an error.
    wb.set_cell("A2".to_string(), JsValue::from_str("=ISERROR(A1)"), None)
        .unwrap();
    wb.set_cell("A3".to_string(), JsValue::from_str("=ERROR.TYPE(A1)"), None)
        .unwrap();
    wb.set_cell("A4".to_string(), JsValue::from_str("=A1+1"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.value, json!(true));

    let a3_js = wb.get_cell("A3".to_string(), None).unwrap();
    let a3: CellData = serde_wasm_bindgen::from_value(a3_js).unwrap();
    assert_json_number(&a3.value, 8.0);

    let a4_js = wb.get_cell("A4".to_string(), None).unwrap();
    let a4: CellData = serde_wasm_bindgen::from_value(a4_js).unwrap();
    assert_eq!(a4.value, json!("#GETTING_DATA"));

    // Spot check one more non-classic error kind.
    wb.set_cell("B2".to_string(), JsValue::from_str("=ERROR.TYPE(B1)"), None)
        .unwrap();
    wb.recalculate(None).unwrap();
    let b2_js = wb.get_cell("B2".to_string(), None).unwrap();
    let b2: CellData = serde_wasm_bindgen::from_value(b2_js).unwrap();
    assert_json_number(&b2.value, 11.0);
}

#[wasm_bindgen_test]
fn scalar_protocol_parses_known_error_strings_but_not_unknown_hash_strings() {
    let mut wb = WasmWorkbook::new();

    // Known error literal → error value.
    wb.set_cell("A1".to_string(), JsValue::from_str("#BLOCKED!"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=ISERROR(A1)"), None)
        .unwrap();
    wb.set_cell("A3".to_string(), JsValue::from_str("=ERROR.TYPE(A1)"), None)
        .unwrap();
    wb.set_cell("A4".to_string(), JsValue::from_str("=A1+1"), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!("#BLOCKED!"));
    assert_eq!(a1.value, json!("#BLOCKED!"));

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.value, json!(true));

    let a3_js = wb.get_cell("A3".to_string(), None).unwrap();
    let a3: CellData = serde_wasm_bindgen::from_value(a3_js).unwrap();
    assert_json_number(&a3.value, 13.0);

    let a4_js = wb.get_cell("A4".to_string(), None).unwrap();
    let a4: CellData = serde_wasm_bindgen::from_value(a4_js).unwrap();
    assert_eq!(a4.value, json!("#BLOCKED!"));

    // Unknown `#FOO` strings must remain plain text.
    wb.set_cell("B1".to_string(), JsValue::from_str("#NOT_A_REAL_ERROR"), None)
        .unwrap();
    wb.set_cell("B2".to_string(), JsValue::from_str("=ISERROR(B1)"), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("#NOT_A_REAL_ERROR"));
    assert_eq!(b1.value, json!("#NOT_A_REAL_ERROR"));

    let b2_js = wb.get_cell("B2".to_string(), None).unwrap();
    let b2: CellData = serde_wasm_bindgen::from_value(b2_js).unwrap();
    assert_eq!(b2.value, json!(false));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_shared_strings_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/shared-strings.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Should not error even though the fixture contains only shared strings.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!("Hello"));
    assert_eq!(a1.value, json!("Hello"));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("World"));
    assert_eq!(b1.value, json!("World"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_shared_formula_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/formulas/shared-formula.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.recalculate(None).unwrap();

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!("=B2*2"));
    assert_json_number(&a2.value, 4.0);

    // Ensure shared-formula cells behave like real formulas (not frozen literal cached values).
    wb.set_cell("B2".to_string(), JsValue::from_f64(10.0), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!("=B2*2"));
    assert_json_number(&a2.value, 20.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_multi_sheet_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/multi-sheet.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let sheet2_a1_js = wb
        .get_cell("A1".to_string(), Some("Sheet2".to_string()))
        .unwrap();
    let sheet2_a1: CellData = serde_wasm_bindgen::from_value(sheet2_a1_js).unwrap();
    assert_json_number(&sheet2_a1.value, 2.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_defined_names() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/metadata/defined-names.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.set_cell("C1".to_string(), JsValue::from_str("=ZedName"), None)
        .unwrap();
    wb.set_cell("C2".to_string(), JsValue::from_str("=ErrName"), None)
        .unwrap();
    wb.set_cell("C3".to_string(), JsValue::from_str("=ERROR.TYPE(ErrName)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=ZedName"));
    assert_eq!(cell.value, json!("Hello"));

    let cell_js = wb.get_cell("C2".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=ErrName"));
    assert_eq!(cell.value, json!("#N/A"));

    let cell_js = wb.get_cell("C3".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=ERROR.TYPE(ErrName)"));
    assert_json_number(&cell.value, 7.0);
}

#[wasm_bindgen_test]
fn cross_sheet_formulas_recalculate() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(1.0),
        Some("Sheet1".to_string()),
    )
    .unwrap();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=Sheet1!A1*2"),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, "Sheet2");
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 2.0);

    let cell_js = wb
        .get_cell("A1".to_string(), Some("Sheet2".to_string()))
        .unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 2.0);
}

#[wasm_bindgen_test]
fn null_inputs_clear_cells_and_recalculate_dependents() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    // Clear A1 by setting it to `null` (empty cell in the JS protocol).
    wb.set_cell("A1".to_string(), JsValue::NULL, None).unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 0.0);

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}

#[wasm_bindgen_test]
fn from_json_treats_null_cells_as_absent() {
    let json_str = r#"{
        "sheets": {
            "Sheet1": {
                "cells": {
                    "A1": null,
                    "A2": "=A1*2"
                }
            }
        }
    }"#;

    let mut wb = WasmWorkbook::from_json(json_str).unwrap();
    wb.recalculate(None).unwrap();

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();

    // JSON import should not store explicit `null` cells.
    assert!(!cells.contains_key("A1"));
    assert!(cells.contains_key("A2"));
}

#[wasm_bindgen_test]
fn set_range_clears_null_entries() {
    let mut wb = WasmWorkbook::new();

    let values: Vec<Vec<JsonValue>> = vec![vec![json!(1), json!(2)]];
    wb.set_range(
        "A1:B1".to_string(),
        serde_wasm_bindgen::to_value(&values).unwrap(),
        None,
    )
    .unwrap();

    let cleared: Vec<Vec<JsonValue>> = vec![vec![JsonValue::Null, json!(2)]];
    wb.set_range(
        "A1:B1".to_string(),
        serde_wasm_bindgen::to_value(&cleared).unwrap(),
        None,
    )
    .unwrap();

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}

#[wasm_bindgen_test]
fn equals_sign_only_is_treated_as_literal_text_input() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("="), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("="));
    assert_eq!(cell.value, json!("="));

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    assert_eq!(parsed["sheets"][DEFAULT_SHEET]["cells"]["A1"], json!("="));
}

#[wasm_bindgen_test]
fn set_cells_bulk_updates_values_and_formulas() {
    #[derive(serde::Serialize)]
    struct Update {
        address: String,
        value: JsonValue,
        sheet: Option<String>,
    }

    let mut wb = WasmWorkbook::new();
    let updates = vec![
        Update {
            address: "A1".to_string(),
            value: json!(1),
            sheet: None,
        },
        Update {
            address: "A2".to_string(),
            value: json!("=A1*2"),
            sheet: None,
        },
        Update {
            address: "A1".to_string(),
            value: json!(10),
            sheet: Some("Sheet2".to_string()),
        },
        Update {
            address: "A2".to_string(),
            value: json!("=A1*2"),
            sheet: Some("Sheet2".to_string()),
        },
    ];

    wb.set_cells(serde_wasm_bindgen::to_value(&updates).unwrap())
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, "Sheet1");
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 2.0);
    assert_eq!(changes[1].sheet, "Sheet2");
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 20.0);

    let sheet1_a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let sheet1_a2: CellData = serde_wasm_bindgen::from_value(sheet1_a2_js).unwrap();
    assert_eq!(sheet1_a2.input, json!("=A1*2"));
    assert_json_number(&sheet1_a2.value, 2.0);
}

#[wasm_bindgen_test]
fn lex_formula_spans_are_utf16() {
    let formula = "=\"a😊b\"";
    let tokens_js = lex_formula(formula, None).unwrap();
    let tokens: Vec<LexToken> = serde_wasm_bindgen::from_value(tokens_js).unwrap();

    let utf16_len = formula.encode_utf16().count();
    let string_tok = tokens
        .iter()
        .find(|tok| tok.kind == "String")
        .unwrap_or_else(|| panic!("expected String token, got {tokens:?}"));
    assert_eq!(
        string_tok.value.as_ref().and_then(|v| v.as_str()),
        Some("a😊b")
    );
    // String token includes surrounding quotes, so it should start at the opening quote right after '='.
    assert_eq!(string_tok.span.start, 1);
    assert_eq!(string_tok.span.end, utf16_len);

    let eof_tok = tokens
        .iter()
        .find(|tok| tok.kind == "Eof")
        .unwrap_or_else(|| panic!("expected Eof token, got {tokens:?}"));
    assert_eq!(eof_tok.span.start, utf16_len);
    assert_eq!(eof_tok.span.end, utf16_len);
}

#[wasm_bindgen_test]
fn parse_formula_partial_reports_function_context() {
    // Cursor at end, inside SUM's second arg (0-indexed argIndex = 1).
    let formula = "=IF(sUm(A1,".to_string();
    let cursor = formula.encode_utf16().count() as u32;

    let parsed_js = parse_formula_partial(formula, Some(cursor), None).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    assert!(parsed.error.is_some());
    assert_eq!(
        parsed.context.function,
        Some(PartialParseFunctionContext {
            name: "SUM".to_string(),
            arg_index: 1,
        })
    );
}

#[wasm_bindgen_test]
fn parse_formula_partial_reports_error_span_utf16() {
    // Unterminated string with a surrogate-pair char in it (😊).
    let formula = "=\"a😊".to_string();
    let cursor = formula.encode_utf16().count() as u32;
    let utf16_len = cursor as usize;

    let parsed_js = parse_formula_partial(formula, Some(cursor), None).unwrap();
    let parsed: PartialParseResult = serde_wasm_bindgen::from_value(parsed_js).unwrap();

    let err = parsed
        .error
        .as_ref()
        .unwrap_or_else(|| panic!("expected error for unterminated string, got {parsed:?}"));
    assert!(err.message.contains("Unterminated string literal"));
    // Error span should cover from the opening quote to the end of input.
    assert_eq!(err.span.start, 1);
    assert_eq!(err.span.end, utf16_len);
}

#[wasm_bindgen_test]
fn rich_values_support_field_access_formulas() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Price": { "type": "number", "value": 12.5 }
            }
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1.Price"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 12.5);

    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(
        got,
        json!({
            "sheet": DEFAULT_SHEET,
            "address": "A1",
            "input": entity.clone(),
            "value": entity.clone(),
        })
    );
}

#[wasm_bindgen_test]
fn rich_values_support_image_inputs() {
    let mut wb = WasmWorkbook::new();

    let image = json!({
        "type": "image",
        "value": {
            "imageId": "image1.png",
            "altText": "Logo"
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&image),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    // Scalar getCell must keep returning scalar values/inputs.
    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert!(a1.input.is_null());
    assert_eq!(a1.value, JsonValue::String("Logo".to_string()));

    // Rich getter preserves the rich input schema, but the engine degrades the value to text today.
    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(
        got,
        json!({
            "sheet": DEFAULT_SHEET,
            "address": "A1",
            "input": image,
            "value": { "type": "string", "value": "Logo" },
        })
    );
}

#[wasm_bindgen_test]
fn rich_values_accept_scalar_cell_value_inputs() {
    let mut wb = WasmWorkbook::new();

    let number = json!({ "type": "number", "value": 42.0 });
    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&number),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    // Scalar API remains scalar-only and should store the scalar input.
    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_json_number(&a1.input, 42.0);
    assert_json_number(&a1.value, 42.0);

    // Rich API should preserve the typed schema.
    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(got["input"]["type"].as_str(), Some("number"));
    assert_json_number(&got["input"]["value"], 42.0);
    assert_eq!(got["value"]["type"].as_str(), Some("number"));
    assert_json_number(&got["value"]["value"], 42.0);
}

#[wasm_bindgen_test]
fn rich_values_accept_error_cell_value_inputs() {
    let mut wb = WasmWorkbook::new();

    let error = json!({ "type": "error", "value": "#FIELD!" });
    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&error),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    // Scalar API keeps returning scalar-ish values.
    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, JsonValue::String("#FIELD!".to_string()));
    assert_eq!(a1.value, JsonValue::String("#FIELD!".to_string()));

    // Rich API should round-trip the typed error schema.
    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(got["input"]["type"].as_str(), Some("error"));
    assert_eq!(got["input"]["value"].as_str(), Some("#FIELD!"));
    assert_eq!(got["value"]["type"].as_str(), Some("error"));
    assert_eq!(got["value"]["value"].as_str(), Some("#FIELD!"));
}

#[wasm_bindgen_test]
fn rich_values_typed_string_preserves_error_like_text() {
    let mut wb = WasmWorkbook::new();

    // Strings that look like error codes must be quote-prefixed in the scalar protocol so they
    // remain literal text (rather than being re-interpreted as errors).
    let text = json!({ "type": "string", "value": "#FIELD!" });
    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&text),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, JsonValue::String("'#FIELD!".to_string()));
    assert_eq!(a1.value, JsonValue::String("#FIELD!".to_string()));

    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(got["input"]["type"].as_str(), Some("string"));
    assert_eq!(got["input"]["value"].as_str(), Some("#FIELD!"));
    assert_eq!(got["value"]["type"].as_str(), Some("string"));
    assert_eq!(got["value"]["value"].as_str(), Some("#FIELD!"));
}

#[wasm_bindgen_test]
fn set_cell_rich_null_clears_previous_value() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "displayValue": "Acme",
            "properties": {}
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    wb.set_cell_rich("A1".to_string(), JsValue::NULL, Some(DEFAULT_SHEET.to_string()))
        .unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert!(a1.input.is_null());
    assert!(a1.value.is_null());

    let rich_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let rich: JsonValue = serde_wasm_bindgen::from_value(rich_js).unwrap();
    assert_eq!(
        rich,
        json!({
            "sheet": DEFAULT_SHEET,
            "address": "A1",
            "input": { "type": "empty" },
            "value": { "type": "empty" },
        })
    );
}

#[wasm_bindgen_test]
fn rich_values_support_bracketed_field_access_formulas() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Change%": { "type": "number", "value": 0.0133 }
            }
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str(r#"=A1.["Change%"]"#),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 0.0133);
}

#[wasm_bindgen_test]
fn calculate_pivot_returns_cell_writes_for_basic_row_sum() {
    #[derive(Debug, serde::Deserialize, PartialEq)]
    #[serde(rename_all = "camelCase")]
    struct PivotCalculationResult {
        writes: Vec<CellChange>,
    }

    let mut wb = WasmWorkbook::new();

    // Source data (headers + records).
    wb.set_cell("A1".to_string(), JsValue::from_str("Category"), None)
        .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("Amount"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("A"), None)
        .unwrap();
    wb.set_cell("B2".to_string(), JsValue::from_f64(10.0), None)
        .unwrap();
    wb.set_cell("A3".to_string(), JsValue::from_str("A"), None)
        .unwrap();
    wb.set_cell("B3".to_string(), JsValue::from_f64(5.0), None)
        .unwrap();
    wb.set_cell("A4".to_string(), JsValue::from_str("B"), None)
        .unwrap();
    wb.set_cell("B4".to_string(), JsValue::from_f64(7.0), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let config = formula_model::pivots::PivotConfig {
        row_fields: vec![formula_model::pivots::PivotField::new("Category")],
        column_fields: vec![],
        value_fields: vec![formula_model::pivots::ValueField {
            source_field: "Amount".to_string(),
            name: "Sum of Amount".to_string(),
            aggregation: formula_model::pivots::AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: formula_model::pivots::Layout::Tabular,
        subtotals: formula_model::pivots::SubtotalPosition::None,
        grand_totals: formula_model::pivots::GrandTotals {
            rows: true,
            columns: false,
        },
    };
    let config_js = serde_wasm_bindgen::to_value(&config).unwrap();

    let result_js = wb
        .calculate_pivot(
            DEFAULT_SHEET.to_string(),
            "A1:B4".to_string(),
            "D1".to_string(),
            config_js,
        )
        .unwrap();
    let result: PivotCalculationResult = serde_wasm_bindgen::from_value(result_js).unwrap();

    assert_eq!(
        result.writes,
        vec![
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D1".to_string(),
                value: json!("Category"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E1".to_string(),
                value: json!("Sum of Amount"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D2".to_string(),
                value: json!("A"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E2".to_string(),
                value: json!(15.0),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D3".to_string(),
                value: json!("B"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E3".to_string(),
                value: json!(7.0),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D4".to_string(),
                value: json!("Grand Total"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E4".to_string(),
                value: json!(22.0),
            },
        ]
    );
}

#[wasm_bindgen_test]
fn rich_values_support_nested_field_access_formulas() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Owner": {
                    "type": "record",
                    "value": {
                        "displayField": "Name",
                        "fields": {
                            "Name": { "type": "string", "value": "Alice" },
                            "Age": { "type": "number", "value": 42.0 }
                        }
                    }
                }
            }
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1.Owner.Age"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 42.0);
}

#[wasm_bindgen_test]
fn rich_values_missing_field_access_returns_field_error() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Price": { "type": "number", "value": 12.5 }
            }
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1.Nope"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String("#FIELD!".to_string()));
}

#[wasm_bindgen_test]
fn rich_values_roundtrip_through_wasm_exports() {
    let mut wb = WasmWorkbook::new();

    let entity = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Price": { "type": "number", "value": 178.5 },
                "Owner": {
                    "type": "record",
                    "value": {
                        "displayField": "Name",
                        "displayValue": "Alice",
                        "fields": {
                            "Name": { "type": "string", "value": "Alice" },
                            "Age": { "type": "number", "value": 42.0 }
                        }
                    }
                }
            }
        }
    });

    wb.set_cell_rich(
        "A1".to_string(),
        to_js_value(&entity),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();

    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: JsonValue = serde_wasm_bindgen::from_value(got_js).unwrap();
    // `serde_wasm_bindgen` may canonicalize JS numbers that are whole integers (e.g. `42.0`) into
    // integer JSON numbers (`42`). Normalize the expected payload through the same JS round-trip so
    // this assertion reflects the actual wasm boundary behavior.
    let expected_entity_js = to_js_value(&entity);
    let expected_entity: JsonValue = serde_wasm_bindgen::from_value(expected_entity_js).unwrap();
    assert_eq!(
        got,
        json!({
            "sheet": DEFAULT_SHEET,
            "address": "A1",
            "input": expected_entity.clone(),
            "value": expected_entity,
        })
    );

    // Scalar API remains scalar-only.
    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert!(cell.input.is_null());
    assert_eq!(cell.value, JsonValue::String("Apple Inc.".to_string()));
}

#[wasm_bindgen_test]
fn rich_values_accept_formula_model_cell_value_schema() {
    use formula_wasm::CellDataRich as WasmCellDataRich;

    let mut wb = WasmWorkbook::new();

    let typed = ModelCellValue::Entity(
        formula_model::EntityValue::new("Apple Inc.")
            .with_entity_type("stock")
            .with_entity_id("AAPL")
            .with_property("Price", 12.5),
    );

    wb.set_cell_rich(
        "A1".to_string(),
        serde_wasm_bindgen::to_value(&typed).unwrap(),
        Some(DEFAULT_SHEET.to_string()),
    )
    .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1.Price"), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 12.5);

    let got_js = wb
        .get_cell_rich("A1".to_string(), Some(DEFAULT_SHEET.to_string()))
        .unwrap();
    let got: WasmCellDataRich = serde_wasm_bindgen::from_value(got_js).unwrap();
    assert_eq!(got.sheet, DEFAULT_SHEET);
    assert_eq!(got.address, "A1");
    assert_eq!(got.input, typed);
    assert_eq!(got.value, got.input);
}

#[wasm_bindgen_test]
fn rich_values_reject_invalid_payloads() {
    let mut wb = WasmWorkbook::new();

    let invalid = json!({ "foo": "bar" });
    let err = wb
        .set_cell_rich(
            "A1".to_string(),
            to_js_value(&invalid),
            Some(DEFAULT_SHEET.to_string()),
        )
        .unwrap_err();

    let message = err
        .as_string()
        .unwrap_or_else(|| "missing error message".to_string());
    assert!(message.contains("invalid rich value"), "{message}");
}

#[wasm_bindgen_test]
fn goal_seek_solves_quadratic_and_updates_workbook_state() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1*A1"), None)
        .unwrap();

    let req = Object::new();
    Reflect::set(&req, &JsValue::from_str("targetCell"), &JsValue::from_str("B1")).unwrap();
    Reflect::set(&req, &JsValue::from_str("targetValue"), &JsValue::from_f64(25.0)).unwrap();
    Reflect::set(&req, &JsValue::from_str("changingCell"), &JsValue::from_str("A1")).unwrap();
    Reflect::set(&req, &JsValue::from_str("tolerance"), &JsValue::from_f64(1e-9)).unwrap();

    let result_js = wb.goal_seek(req.into()).unwrap();
    let result: JsonValue = serde_wasm_bindgen::from_value(result_js).unwrap();
    assert_eq!(result["success"].as_bool(), Some(true), "{result:?}");
    let solution = result["solution"]
        .as_f64()
        .unwrap_or_else(|| panic!("expected numeric solution, got {result:?}"));
    assert!(
        (solution - 5.0).abs() < 1e-6,
        "expected solution ≈ 5, got {solution}"
    );

    // After recalc, the workbook should reflect the solved value and the target formula result.
    wb.recalculate(None).unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    let a1_value = a1.value.as_f64().unwrap();
    assert!((a1_value - 5.0).abs() < 1e-6, "A1 = {a1_value}");
    let a1_input = a1.input.as_f64().unwrap();
    assert!((a1_input - 5.0).abs() < 1e-6, "A1 input = {a1_input}");

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    let b1_value = b1.value.as_f64().unwrap();
    assert!((b1_value - 25.0).abs() < 1e-6, "B1 = {b1_value}");
}
