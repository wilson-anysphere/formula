#![cfg(target_arch = "wasm32")]

use serde_json::json;
use serde_json::Value as JsonValue;
use js_sys::{Object, Reflect};
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_engine::pivot::{PivotFieldType, PivotSchema, PivotValue};
use formula_model::CellValue as ModelCellValue;
use formula_wasm::{
    canonicalize_formula, get_locale_info, lex_formula, localize_formula, parse_formula_partial,
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

#[wasm_bindgen_test]
fn style_number_format_null_clears_lower_layers_for_cell_format() {
    let mut wb = WasmWorkbook::new();

    let row_style_id = wb
        .intern_style(to_js_value(&json!({ "numberFormat": "0.00" })))
        .unwrap();
    wb.set_row_style_id(DEFAULT_SHEET.to_string(), 0, Some(row_style_id));

    let clear_style_id = wb
        .intern_style(to_js_value(&json!({ "numberFormat": JsonValue::Null })))
        .unwrap();
    assert_ne!(clear_style_id, 0, "expected explicit clear style to intern");
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), clear_style_id)
        .unwrap();

    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str(r#"=CELL("format",A1)"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, json!("G"));
}

#[wasm_bindgen_test]
fn style_alignment_horizontal_null_clears_lower_layers_for_cell_prefix() {
    let mut wb = WasmWorkbook::new();

    let row_style_id = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": "right" } })))
        .unwrap();
    wb.set_row_style_id(DEFAULT_SHEET.to_string(), 1, Some(row_style_id));

    let clear_style_id = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": JsonValue::Null } })))
        .unwrap();
    assert_ne!(clear_style_id, 0, "expected explicit clear style to intern");
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A2".to_string(), clear_style_id)
        .unwrap();

    wb.set_cell(
        "B2".to_string(),
        JsValue::from_str(r#"=CELL("prefix",A2)"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let b2_js = wb.get_cell("B2".to_string(), None).unwrap();
    let b2: CellData = serde_wasm_bindgen::from_value(b2_js).unwrap();
    assert_eq!(b2.value, json!(""));
}

#[wasm_bindgen_test]
fn style_locked_null_clears_lower_layers_for_cell_protect() {
    let mut wb = WasmWorkbook::new();

    let row_style_id = wb
        .intern_style(to_js_value(&json!({ "locked": false })))
        .unwrap();
    wb.set_row_style_id(DEFAULT_SHEET.to_string(), 2, Some(row_style_id));

    let clear_style_id = wb
        .intern_style(to_js_value(&json!({ "locked": JsonValue::Null })))
        .unwrap();
    assert_ne!(clear_style_id, 0, "expected explicit clear style to intern");
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A3".to_string(), clear_style_id)
        .unwrap();

    wb.set_cell(
        "B3".to_string(),
        JsValue::from_str(r#"=CELL("protect",A3)"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let b3_js = wb.get_cell("B3".to_string(), None).unwrap();
    let b3: CellData = serde_wasm_bindgen::from_value(b3_js).unwrap();
    assert_json_number(&b3.value, 1.0);
}

#[wasm_bindgen_test]
fn set_format_runs_by_col_accepts_valid_runs_and_null_clears() {
    let mut wb = WasmWorkbook::new();

    let style_id = wb
        .intern_style(to_js_value(&json!({ "numberFormat": "0.00" })))
        .unwrap();

    wb.set_format_runs_by_col(
        DEFAULT_SHEET.to_string(),
        0,
        to_js_value(&json!([{
            "startRow": 0,
            "endRowExclusive": 2,
            "styleId": style_id,
        }])),
    )
    .unwrap();

    // `null`/`undefined` clears the run list.
    wb.set_format_runs_by_col(DEFAULT_SHEET.to_string(), 0, JsValue::NULL)
        .unwrap();
}

#[wasm_bindgen_test]
fn set_format_runs_by_col_rejects_non_array_payloads() {
    let mut wb = WasmWorkbook::new();

    let err = wb
        .set_format_runs_by_col(
            DEFAULT_SHEET.to_string(),
            0,
            JsValue::from_str("not an array"),
        )
        .expect_err("expected setFormatRunsByCol to reject non-array input");

    assert_eq!(
        err.as_string().unwrap_or_default(),
        "setFormatRunsByCol: runs must be an array"
    );
}

#[wasm_bindgen_test]
fn set_format_runs_by_col_validates_run_ranges() {
    let mut wb = WasmWorkbook::new();

    let style_id = wb
        .intern_style(to_js_value(&json!({ "numberFormat": "0.00" })))
        .unwrap();

    let err = wb
        .set_format_runs_by_col(
            DEFAULT_SHEET.to_string(),
            0,
            to_js_value(&json!([{
                "startRow": 1,
                "endRowExclusive": 1,
                "styleId": style_id,
            }])),
        )
        .expect_err("expected setFormatRunsByCol to reject invalid run range");

    assert_eq!(
        err.as_string().unwrap_or_default(),
        "setFormatRunsByCol: runs[0].endRowExclusive must be greater than startRow"
    );
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

#[derive(Debug, serde::Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct CalcSettings {
    calculation_mode: String,
    calculate_before_save: bool,
    full_precision: bool,
    full_calc_on_load: bool,
    iterative: IterativeCalcSettings,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct IterativeCalcSettings {
    enabled: bool,
    max_iterations: u32,
    max_change: f64,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct GoalSeekResultDto {
    status: String,
    solution: f64,
    iterations: usize,
    final_output: f64,
    final_error: f64,
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct GoalSeekResponse {
    result: GoalSeekResultDto,
    changes: Vec<CellChange>,
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
    assert!(message.contains("unknown localeId: xx-XX"));
    assert!(
        message.contains("Supported locale ids"),
        "expected actionable locale message, got {message:?}"
    );
    assert!(
        message.contains("en-US"),
        "expected supported locale ids list to include en-US, got {message:?}"
    );
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
fn canonicalize_and_localize_error_literals_roundtrip_es_es_inverted_punctuation() {
    // es-ES uses inverted punctuation variants for many error literals.
    // Ensure these round-trip through the public wasm API.
    for localized in ["=#¡VALOR!", "=#¡valor!", "=#VALOR!", "=#valor!"] {
        assert_eq!(
            canonicalize_formula(localized, "es-ES", None).unwrap(),
            "=#VALUE!"
        );
    }
    assert_eq!(
        localize_formula("=#VALUE!", "es-ES", None).unwrap(),
        "=#¡VALOR!"
    );

    for localized in ["=#¿NOMBRE?", "=#¿nombre?", "=#NOMBRE?", "=#nombre?"] {
        assert_eq!(
            canonicalize_formula(localized, "es-ES", None).unwrap(),
            "=#NAME?"
        );
    }
    assert_eq!(
        localize_formula("=#NAME?", "es-ES", None).unwrap(),
        "=#¿NOMBRE?"
    );
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
fn canonicalize_formula_normalizes_locale_id_variants() {
    let localized = "=SUMME(1,5;2)";
    let expected = "=SUM(1.5,2)";

    // JS callers may provide OS/browser locale spellings or case/whitespace variants. Ensure the
    // wasm public API remains compatible with the engine's locale-id normalization rules.
    for locale_id in [
        "de_DE.UTF-8",
        "de_DE.UTF-8@euro",
        "de_DE@euro",
        "de_DE",
        "  de-DE  ",
        "DE-de",
        "de",
    ] {
        let canonical = canonicalize_formula(localized, locale_id, None).unwrap();
        assert_eq!(
            canonical, expected,
            "expected canonicalize_formula to accept locale id variant {locale_id:?}"
        );
    }
}

#[wasm_bindgen_test]
fn set_locale_accepts_normalized_locale_ids_and_aliases() {
    // POSIX-like locale tags should work (normalization drops the encoding/modifier suffix).
    for locale_id in ["de_DE.UTF-8", "de_DE.UTF-8@euro", "de_DE@euro"] {
        let mut wb = WasmWorkbook::new();
        assert!(
            wb.set_locale(locale_id.to_string()),
            "expected set_locale to accept POSIX locale id variant {locale_id:?}"
        );
    }

    // Aliases (added in Task 456).
    let mut failures = Vec::new();
    for locale_id in ["C", "POSIX"] {
        let mut wb = WasmWorkbook::new();
        if !wb.set_locale(locale_id.to_string()) {
            failures.push(locale_id);
        }
    }
    assert!(
        failures.is_empty(),
        "expected set_locale to accept locale aliases: {failures:?}"
    );
}

#[wasm_bindgen_test]
fn canonicalize_formula_accepts_locale_aliases() {
    // Ensure canonicalizeFormula accepts the same alias forms as the engine's locale
    // normalization (not just workbook.setLocale).
    for locale_id in ["C", "POSIX"] {
        let canonical = canonicalize_formula("=1+2", locale_id, None).unwrap();
        assert_eq!(
            canonical, "=1+2",
            "expected canonicalize_formula to accept locale alias {locale_id:?}"
        );
    }
}

#[derive(Debug, serde::Deserialize)]
#[serde(rename_all = "camelCase")]
struct LocaleInfo {
    locale_id: String,
}

#[wasm_bindgen_test]
fn get_locale_info_accepts_chinese_script_locale_ids() {
    // When the region is missing, BCP-47 script subtags should resolve to the engine's canonical
    // region defaults.
    for (locale_id, expected_canonical) in [("zh-Hant", "zh-TW"), ("zh-Hans", "zh-CN")] {
        let info_js = get_locale_info(locale_id).unwrap();
        let info: LocaleInfo = serde_wasm_bindgen::from_value(info_js).unwrap();
        assert_eq!(
            info.locale_id, expected_canonical,
            "expected get_locale_info({locale_id:?}) to canonicalize to {expected_canonical:?}"
        );

        // Also ensure other public APIs accept these locale ids.
        let canonical = canonicalize_formula("=1+2", locale_id, None).unwrap();
        assert_eq!(canonical, "=1+2");
    }
}

#[wasm_bindgen_test]
fn canonicalize_formula_accepts_posix_locale_id_variants_at_api_boundary() {
    // JS callers may supply POSIX locale ids that include encoding/modifier suffixes, as well as
    // case/whitespace variants. Ensure these are accepted by the public wasm API.
    assert_eq!(
        canonicalize_formula("=SUMME(1;2)", "de_DE.UTF-8", None).unwrap(),
        "=SUM(1,2)"
    );
    assert_eq!(
        canonicalize_formula("=SUMME(1;2)", "de_DE@euro", None).unwrap(),
        "=SUM(1,2)"
    );
    assert_eq!(
        canonicalize_formula("=SUMME(1;2)", "  DE-de  ", None).unwrap(),
        "=SUM(1,2)"
    );
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
    assert!(message.contains("unknown localeId: xx-XX"));
    assert!(
        message.contains("Supported locale ids"),
        "expected actionable locale message, got {message:?}"
    );
    assert!(
        message.contains("en-US"),
        "expected supported locale ids list to include en-US, got {message:?}"
    );
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
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), 42)
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
    assert_eq!(wb.get_cell_style_id("A1".to_string(), None).unwrap(), 42);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}

#[wasm_bindgen_test]
fn null_inputs_preserve_cell_style_metadata_in_engine() {
    let mut wb = WasmWorkbook::new();

    // Use a fixed 2-decimal number format so `CELL("format")` returns a stable, Excel-like
    // classification code ("F2"). Also mark the cell as unlocked so `CELL("protect")` returns 0.
    let style = formula_model::Style {
        number_format: Some("0.00".to_string()),
        protection: Some(formula_model::Protection {
            locked: false,
            hidden: false,
        }),
        ..Default::default()
    };
    let style_id = wb.intern_style(to_js_value(&style)).unwrap();
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), style_id)
        .unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str(r#"=CELL("format",A1)"#),
        None,
    )
    .unwrap();
    wb.set_cell(
        "C1".to_string(),
        JsValue::from_str(r#"=CELL("protect",A1)"#),
        None,
    )
    .unwrap();
    wb.set_cell(
        "D1".to_string(),
        JsValue::from_str(r#"=ISBLANK(A1)"#),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();
    let cell_js = wb.get_cell("B1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::String("F2".to_string()));

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 0.0);

    let cell_js = wb.get_cell("D1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::Bool(false));

    // Clear A1 by setting it to `null` (empty cell in the JS protocol). Formatting should remain
    // intact so `CELL("format",A1)`/`CELL("protect",A1)` continue to observe style metadata.
    wb.set_cell("A1".to_string(), JsValue::NULL, None).unwrap();
    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("B1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::String("F2".to_string()));

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 0.0);

    let cell_js = wb.get_cell("D1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::Bool(true));

    // `toJson()` should remain sparse: clearing a value should not serialize an explicit blank,
    // even if the engine preserved formatting metadata.
    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}

#[wasm_bindgen_test]
fn cell_protect_respects_explicit_locked_overrides() {
    let mut wb = WasmWorkbook::new();

    wb.set_cell("A1".to_string(), JsValue::from_str("x"), None)
        .unwrap();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str(r#"=CELL("protect",A1)"#),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 1.0);

    // Row-level style marks the cell as unlocked.
    let unlocked = wb
        .intern_style(to_js_value(&json!({ "protection": { "locked": false } })))
        .unwrap();
    wb.set_row_style_id(DEFAULT_SHEET.to_string(), 0, Some(unlocked));
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 0.0);

    // Cell-level style should be able to override the inherited unlocked flag by explicitly
    // clearing the lock flag back to Excel's default locked=true.
    let clear = wb
        .intern_style(to_js_value(&json!({ "protection": { "locked": null } })))
        .unwrap();
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), clear)
        .unwrap();
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 1.0);
}

#[wasm_bindgen_test]
fn cell_prefix_respects_effective_alignment_and_explicit_clears() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("x"), None)
        .unwrap();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str(r#"=CELL("prefix",A1)"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String(String::new()));

    // Row alignment should affect CELL("prefix").
    let style_right = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": "right" } })))
        .unwrap();
    wb.set_row_style_id(DEFAULT_SHEET.to_string(), 0, Some(style_right));
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String("\"".to_string()));

    // Range-run alignment should override row/col styles.
    let style_center = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": "center" } })))
        .unwrap();
    wb.set_format_runs_by_col(
        DEFAULT_SHEET.to_string(),
        0,
        to_js_value(&json!([{ "startRow": 0, "endRowExclusive": 10, "styleId": style_center }])),
    )
    .unwrap();
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String("^".to_string()));

    // Cell-level alignment wins over the range run.
    let style_fill = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": "fill" } })))
        .unwrap();
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), style_fill)
        .unwrap();
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String("\\".to_string()));

    // Explicit clear (`horizontal: null`) should override inherited formatting and revert to General.
    let style_clear = wb
        .intern_style(to_js_value(&json!({ "alignment": { "horizontal": null } })))
        .unwrap();
    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), style_clear)
        .unwrap();
    wb.recalculate(None).unwrap();
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.value, JsonValue::String(String::new()));
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
            source_field: "Amount".into(),
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
                value: json!(15),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D3".to_string(),
                value: json!("B"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E3".to_string(),
                value: json!(7),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "D4".to_string(),
                value: json!("Grand Total"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "E4".to_string(),
                value: json!(22),
            },
        ]
    );
}

#[wasm_bindgen_test]
fn get_pivot_schema_reports_field_types_and_limits_samples() {
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

    // Only sample the first two records.
    let schema_js = wb
        .get_pivot_schema(DEFAULT_SHEET.to_string(), "A1:B4".to_string(), Some(2))
        .unwrap();
    let schema: PivotSchema = serde_wasm_bindgen::from_value(schema_js).unwrap();

    assert_eq!(schema.record_count, 3);
    assert_eq!(schema.fields.len(), 2);

    assert_eq!(schema.fields[0].name, "Category");
    assert_eq!(schema.fields[0].field_type, PivotFieldType::Text);
    assert_eq!(
        schema.fields[0].sample_values,
        vec![
            PivotValue::Text("A".to_string()),
            PivotValue::Text("A".to_string()),
        ]
    );

    assert_eq!(schema.fields[1].name, "Amount");
    assert_eq!(schema.fields[1].field_type, PivotFieldType::Number);
    assert_eq!(
        schema.fields[1].sample_values,
        vec![PivotValue::Number(10.0), PivotValue::Number(5.0),]
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
fn from_xlsx_bytes_imports_style_and_column_metadata() {
    let bytes = include_bytes!("../../../fixtures/xlsx/metadata/style-only-cell.xlsx");
    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Recalculate to evaluate the workbook's formulas.
    wb.recalculate(None).unwrap();

    // A1 is a style-only cell (no value/formula) with "locked=false". If style-only cells are
    // dropped during import, this will incorrectly evaluate to 1 (locked).
    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 0.0);

    // The fixture also sets a column width + hidden flag for column A. Ensure that metadata is
    // imported into the engine so CELL("width") sees it.
    wb.set_cell(
        "C1".to_string(),
        JsValue::from_str(r#"=CELL("width",A1)"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();
    let c1_js = wb.get_cell("C1".to_string(), None).unwrap();
    let c1: CellData = serde_wasm_bindgen::from_value(c1_js).unwrap();

    // Hidden columns report a width of 0.
    assert_json_number(&c1.value, 0.0);

    // Unhide the column; CELL("width") should now report the stored width with the Excel
    // fractional marker for an explicit width override (`+0.1`).
    wb.set_col_hidden(DEFAULT_SHEET.to_string(), 0, false).unwrap();
    wb.recalculate(None).unwrap();
    let c1_js = wb.get_cell("C1".to_string(), None).unwrap();
    let c1: CellData = serde_wasm_bindgen::from_value(c1_js).unwrap();
    assert_json_number(&c1.value, 20.1);

    // Style-only cells should *not* appear in the legacy scalar IO schema returned by `toJson()`.
    let json = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&json).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"]
        .as_object()
        .expect("expected cells object");
    assert!(
        !cells.contains_key("A1"),
        "style-only cell A1 should be omitted from toJson() but was present: {cells:?}"
    );

    // Workbook file metadata should round-trip into functions like CELL("filename") / INFO("directory").
    wb.set_cell(
        "D1".to_string(),
        JsValue::from_str(r#"=CELL("filename")"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();
    let d1_js = wb.get_cell("D1".to_string(), None).unwrap();
    let d1: CellData = serde_wasm_bindgen::from_value(d1_js).unwrap();
    assert_eq!(d1.value, JsonValue::String(String::new()));

    wb.set_workbook_file_metadata(
        JsValue::from_str(r#"C:\foo"#),
        JsValue::from_str("Book1.xlsx"),
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let d1_js = wb.get_cell("D1".to_string(), None).unwrap();
    let d1: CellData = serde_wasm_bindgen::from_value(d1_js).unwrap();
    assert_eq!(
        d1.value,
        JsonValue::String(r#"C:\foo\[Book1.xlsx]Sheet1"#.to_string())
    );

    wb.set_cell(
        "E1".to_string(),
        JsValue::from_str(r#"=INFO("directory")"#),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();
    let e1_js = wb.get_cell("E1".to_string(), None).unwrap();
    let e1: CellData = serde_wasm_bindgen::from_value(e1_js).unwrap();
    assert_eq!(e1.value, JsonValue::String(r#"C:\foo\"#.to_string()));
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
    let result: GoalSeekResponse = serde_wasm_bindgen::from_value(result_js).unwrap();
    assert_eq!(result.result.status, "Converged", "{result:?}");
    let solution = result.result.solution;
    assert!(
        (solution - 5.0).abs() < 1e-6,
        "expected solution ≈ 5, got {solution}"
    );
    assert!(
        result.result.iterations > 0,
        "expected goal seek to perform at least one iteration, got {}",
        result.result.iterations
    );
    assert!(
        (result.result.final_output - 25.0).abs() < 1e-6,
        "expected final output ≈ 25, got {}",
        result.result.final_output
    );
    assert!(
        result.result.final_error.abs() < 1e-6,
        "expected final error to be near zero, got {}",
        result.result.final_error
    );

    // Goal seek should include final input/output changes for cache updates.
    let a1_change = result
        .changes
        .iter()
        .find(|c| c.sheet == DEFAULT_SHEET && c.address == "A1")
        .expect("expected A1 change");
    assert!((a1_change.value.as_f64().unwrap() - 5.0).abs() < 1e-6);
    let b1_change = result
        .changes
        .iter()
        .find(|c| c.sheet == DEFAULT_SHEET && c.address == "B1")
        .expect("expected B1 change");
    assert!((b1_change.value.as_f64().unwrap() - 25.0).abs() < 1e-6);

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

#[wasm_bindgen_test]
fn goal_seek_rejects_non_object_request() {
    let mut wb = WasmWorkbook::new();
    let err = wb.goal_seek(JsValue::from_f64(1.0)).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert!(
        message.contains("expected") || message.contains("object"),
        "{message}"
    );
}

#[wasm_bindgen_test]
fn goal_seek_rejects_non_finite_target_value() {
    let mut wb = WasmWorkbook::new();
    let req = Object::new();
    Reflect::set(&req, &JsValue::from_str("targetCell"), &JsValue::from_str("B1")).unwrap();
    Reflect::set(&req, &JsValue::from_str("targetValue"), &JsValue::from_f64(f64::NAN)).unwrap();
    Reflect::set(&req, &JsValue::from_str("changingCell"), &JsValue::from_str("A1")).unwrap();
    let err = wb.goal_seek(req.into()).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "targetValue must be a finite number");
}

#[wasm_bindgen_test]
fn goal_seek_rejects_invalid_addresses() {
    let mut wb = WasmWorkbook::new();
    let req = Object::new();
    Reflect::set(&req, &JsValue::from_str("targetCell"), &JsValue::from_str("A0")).unwrap();
    Reflect::set(&req, &JsValue::from_str("targetValue"), &JsValue::from_f64(1.0)).unwrap();
    Reflect::set(&req, &JsValue::from_str("changingCell"), &JsValue::from_str("A1")).unwrap();
    let err = wb.goal_seek(req.into()).unwrap_err();
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert_eq!(message, "invalid cell address: A0");
}

#[wasm_bindgen_test]
fn get_pivot_field_items_deduplicates_and_sorts() {
    let mut wb = WasmWorkbook::new();

    // Source data: a simple 2-column dataset with repeated category values and one blank.
    wb.set_cell("A1".to_string(), JsValue::from_str("Category"), None)
        .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("Value"), None)
        .unwrap();

    wb.set_cell("A2".to_string(), JsValue::from_str("B"), None)
        .unwrap();
    wb.set_cell("B2".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();

    wb.set_cell("A3".to_string(), JsValue::from_str("A"), None)
        .unwrap();
    wb.set_cell("B3".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();

    // Duplicate category value.
    wb.set_cell("A4".to_string(), JsValue::from_str("B"), None)
        .unwrap();
    wb.set_cell("B4".to_string(), JsValue::from_f64(3.0), None)
        .unwrap();

    // Blank category.
    wb.set_cell("A5".to_string(), JsValue::NULL, None).unwrap();
    wb.set_cell("B5".to_string(), JsValue::from_f64(4.0), None)
        .unwrap();

    let items_js = wb
        .get_pivot_field_items(
            DEFAULT_SHEET.to_string(),
            "A1:B5".to_string(),
            "Category".to_string(),
        )
        .unwrap();
    let items: Vec<PivotValue> = serde_wasm_bindgen::from_value(items_js).unwrap();

    // Pivot item ordering is stable: text values sort alphabetically (case-insensitive), and
    // blanks always appear last.
    assert_eq!(
        items,
        vec![
            PivotValue::Text("A".to_string()),
            PivotValue::Text("B".to_string()),
            PivotValue::Blank
        ]
    );
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_populates_calc_settings_via_get_calc_settings() {
    let bytes = include_bytes!("../../formula-xlsx/tests/fixtures/calc_settings.xlsx");
    let wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    let settings_js = wb.get_calc_settings().unwrap();
    let settings: CalcSettings = serde_wasm_bindgen::from_value(settings_js).unwrap();

    assert_eq!(settings.calculation_mode, "manual");
    assert!(settings.calculate_before_save);
    assert!(settings.full_precision);
    assert!(!settings.full_calc_on_load);
    assert!(settings.iterative.enabled);
    assert_eq!(settings.iterative.max_iterations, 10);
    assert!((settings.iterative.max_change - 0.0001).abs() < 1e-12);
}

#[wasm_bindgen_test]
fn from_encrypted_xlsx_bytes_decrypts_and_loads_workbook() {
    let plaintext: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsx/tests/fixtures/rt_simple.xlsx"
    ));
    let password = "secret-password";

    // Encrypt the fixture in-process so the test can run fully in wasm without committed binary
    // encrypted fixtures.
    //
    // Use modest (non-production) parameters to keep wasm-bindgen-test runtime reasonable.
    let encrypted = formula_office_crypto::encrypt_package_to_ole(
        plaintext,
        password,
        formula_office_crypto::EncryptOptions {
            key_bits: 128,
            hash_algorithm: formula_office_crypto::HashAlgorithm::Sha256,
            spin_count: 1_000,
            ..Default::default()
        },
    )
    .expect("encrypt");
    let wb =
        WasmWorkbook::from_encrypted_xlsx_bytes(&encrypted, password.to_string()).expect("load");

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.value, JsonValue::String("Hello".to_string()));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 42.0);
}

#[wasm_bindgen_test]
fn from_encrypted_xlsx_bytes_rejects_invalid_password() {
    let plaintext: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsx/tests/fixtures/rt_simple.xlsx"
    ));
    let password = "secret-password";

    let encrypted = formula_office_crypto::encrypt_package_to_ole(
        plaintext,
        password,
        formula_office_crypto::EncryptOptions {
            key_bits: 128,
            hash_algorithm: formula_office_crypto::HashAlgorithm::Sha256,
            spin_count: 1_000,
            ..Default::default()
        },
    )
    .expect("encrypt");

    let err = match WasmWorkbook::from_encrypted_xlsx_bytes(&encrypted, "wrong-password".to_string()) {
        Ok(_) => panic!("expected invalid password error"),
        Err(err) => err,
    };
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert!(
        message.to_lowercase().contains("invalid password"),
        "expected invalid password error, got {message:?}"
    );
}

#[wasm_bindgen_test]
fn from_encrypted_xlsx_bytes_opens_xlsb_payload() {
    let plaintext: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let password = "secret-password";

    let encrypted = formula_office_crypto::encrypt_package_to_ole(
        plaintext,
        password,
        formula_office_crypto::EncryptOptions {
            key_bits: 128,
            hash_algorithm: formula_office_crypto::HashAlgorithm::Sha256,
            spin_count: 1_000,
            ..Default::default()
        },
    )
    .expect("encrypt");

    let mut wb = WasmWorkbook::from_encrypted_xlsx_bytes(&encrypted, password.to_string()).unwrap();
    wb.recalculate(None).unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.value, json!("Hello"));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_json_number(&b1.value, 42.5);

    let c1_js = wb.get_cell("C1".to_string(), None).unwrap();
    let c1: CellData = serde_wasm_bindgen::from_value(c1_js).unwrap();
    assert_json_number(&c1.value, 85.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_reports_password_required_for_encrypted_inputs() {
    let plaintext: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsx/tests/fixtures/rt_simple.xlsx"
    ));
    let password = "secret-password";

    let encrypted = formula_office_crypto::encrypt_package_to_ole(
        plaintext,
        password,
        formula_office_crypto::EncryptOptions {
            key_bits: 128,
            hash_algorithm: formula_office_crypto::HashAlgorithm::Sha256,
            spin_count: 1_000,
            ..Default::default()
        },
    )
    .expect("encrypt");

    let err = match WasmWorkbook::from_xlsx_bytes(&encrypted) {
        Ok(_) => panic!("expected from_xlsx_bytes to fail for encrypted workbook"),
        Err(err) => err,
    };
    let message = err
        .as_string()
        .unwrap_or_else(|| format!("unexpected error value: {err:?}"));
    assert!(
        message.contains("fromEncryptedXlsxBytes") || message.to_lowercase().contains("password"),
        "expected encrypted workbook error to mention fromEncryptedXlsxBytes/password, got {message:?}"
    );
}

#[wasm_bindgen_test]
fn cell_filename_updates_after_set_workbook_file_metadata() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=CELL(\"filename\")"),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    let before_js = wb.get_cell("A1".to_string(), None).unwrap();
    let before: CellData = serde_wasm_bindgen::from_value(before_js).unwrap();
    assert_eq!(before.value, JsonValue::String("".to_string()));

    wb.set_workbook_file_metadata(JsValue::from_str("/tmp"), JsValue::from_str("book.xlsx"))
        .unwrap();
    wb.recalculate(None).unwrap();

    let after_js = wb.get_cell("A1".to_string(), None).unwrap();
    let after: CellData = serde_wasm_bindgen::from_value(after_js).unwrap();
    assert_eq!(
        after.value,
        JsonValue::String(format!("/tmp/[book.xlsx]{DEFAULT_SHEET}"))
    );
}

#[wasm_bindgen_test]
fn cell_filename_reflects_sheet_display_name() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=CELL(\"filename\")"),
        None,
    )
    .unwrap();
    wb.set_workbook_file_metadata(JsValue::from_str("/tmp"), JsValue::from_str("book.xlsx"))
        .unwrap();
    wb.recalculate(None).unwrap();

    let before_js = wb.get_cell("A1".to_string(), None).unwrap();
    let before: CellData = serde_wasm_bindgen::from_value(before_js).unwrap();
    assert_eq!(
        before.value,
        JsonValue::String(format!("/tmp/[book.xlsx]{DEFAULT_SHEET}"))
    );

    wb.set_sheet_display_name(DEFAULT_SHEET.to_string(), "Summary".to_string())
        .unwrap();
    wb.recalculate(None).unwrap();

    let after_js = wb.get_cell("A1".to_string(), None).unwrap();
    let after: CellData = serde_wasm_bindgen::from_value(after_js).unwrap();
    assert_eq!(after.value, JsonValue::String("/tmp/[book.xlsx]Summary".to_string()));
}

#[wasm_bindgen_test]
fn cell_format_reflects_intern_style_and_set_cell_style_id() {
    let mut wb = WasmWorkbook::new();

    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str("=CELL(\"format\",A1)"),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    let before_js = wb.get_cell("B1".to_string(), None).unwrap();
    let before: CellData = serde_wasm_bindgen::from_value(before_js).unwrap();
    assert_eq!(before.value, JsonValue::String("G".to_string()));

    // The worker-side DTO uses `numberFormat` (camelCase), but accept `number_format` too for
    // backward compatibility with Rust/serde shapes.
    let fmt_camel = Object::new();
    Reflect::set(
        &fmt_camel,
        &JsValue::from_str("numberFormat"),
        &JsValue::from_str("0.00"),
    )
    .unwrap();
    let style_id_camel = wb.intern_style(fmt_camel.into()).unwrap();

    let fmt_snake = Object::new();
    Reflect::set(
        &fmt_snake,
        &JsValue::from_str("number_format"),
        &JsValue::from_str("0.00"),
    )
    .unwrap();
    let style_id_snake = wb.intern_style(fmt_snake.into()).unwrap();
    assert_eq!(style_id_camel, style_id_snake);

    wb.set_cell_style_id(DEFAULT_SHEET.to_string(), "A1".to_string(), style_id_camel)
        .unwrap();
    wb.recalculate(None).unwrap();

    let after_js = wb.get_cell("B1".to_string(), None).unwrap();
    let after: CellData = serde_wasm_bindgen::from_value(after_js).unwrap();
    assert_eq!(after.value, JsonValue::String("F2".to_string()));
}

#[wasm_bindgen_test]
fn cell_width_reflects_set_col_width_chars() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "B1".to_string(),
        JsValue::from_str("=CELL(\"width\",A1)"),
        None,
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    wb.set_col_width_chars(
        DEFAULT_SHEET.to_string(),
        0,
        JsValue::from_f64(16.42578125),
    )
    .unwrap();
    wb.recalculate(None).unwrap();

    let after_js = wb.get_cell("B1".to_string(), None).unwrap();
    let after: CellData = serde_wasm_bindgen::from_value(after_js).unwrap();
    // Excel's `CELL("width")` returns the integer part of the width (rounded down) and uses the
    // first decimal digit as a flag for whether the width is an explicit per-column override.
    // See `crates/formula-engine/src/functions/information/worksheet.rs`.
    assert_json_number(&after.value, 16.1);

    // Clearing the override should revert to the default width and clear the "custom width" flag.
    wb.set_col_width_chars(DEFAULT_SHEET.to_string(), 0, JsValue::NULL)
        .unwrap();
    wb.recalculate(None).unwrap();
    let cleared_js = wb.get_cell("B1".to_string(), None).unwrap();
    let cleared: CellData = serde_wasm_bindgen::from_value(cleared_js).unwrap();
    let cleared_width = cleared
        .value
        .as_f64()
        .unwrap_or_else(|| panic!("expected number, got {:?}", cleared.value));
    assert!(
        (cleared_width - 8.0).abs() < 1e-6,
        "expected cleared width to revert to default; got {cleared_width}"
    );
}
