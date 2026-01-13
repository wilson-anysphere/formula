use std::collections::{BTreeMap, BTreeSet, HashMap};

use formula_engine::{
    CellAddr, Coord, EditError as EngineEditError, EditOp as EngineEditOp,
    EditResult as EngineEditResult, Engine, ErrorKind, NameDefinition, NameScope, ParseOptions,
    RecalcMode, Span as EngineSpan, Token, TokenKind, Value as EngineValue,
};
use formula_engine::editing::rewrite::rewrite_formula_for_copy_delta;
use formula_engine::locale::{
    canonicalize_formula_with_style, get_locale, localize_formula_with_style, FormulaLocale,
    ValueLocaleConfig, DE_DE, EN_US, ES_ES, FR_FR,
};
use formula_engine::pivot as pivot_engine;
use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams};
use formula_engine::what_if::EngineWhatIfModel;
use formula_model::{
    display_formula_text, CellRef, CellValue, DateSystem, DefinedNameScope, Range, EXCEL_MAX_COLS,
    EXCEL_MAX_ROWS,
};
use js_sys::{Array, Object, Reflect};
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use wasm_bindgen::prelude::*;

#[cfg(feature = "dax")]
mod dax;
#[cfg(feature = "dax")]
pub use dax::{DaxFilterContext, DaxModel, WasmDaxDataModel};

pub const DEFAULT_SHEET: &str = "Sheet1";

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellData {
    pub sheet: String,
    pub address: String,
    pub input: JsonValue,
    pub value: JsonValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellDataRich {
    pub sheet: String,
    pub address: String,
    pub input: CellValue,
    pub value: CellValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellChange {
    pub sheet: String,
    pub address: String,
    pub value: JsonValue,
}

fn js_err(message: impl ToString) -> JsValue {
    JsValue::from_str(&message.to_string())
}

fn require_formula_locale(locale_id: &str) -> Result<&'static FormulaLocale, JsValue> {
    get_locale(locale_id).ok_or_else(|| {
        js_err(format!(
            "unknown localeId: {locale_id}. Supported locale ids: {}, {}, {}, {}",
            EN_US.id, DE_DE.id, FR_FR.id, ES_ES.id
        ))
    })
}

fn parse_reference_style(
    reference_style: Option<String>,
) -> Result<formula_engine::ReferenceStyle, JsValue> {
    match reference_style.as_deref().unwrap_or("A1") {
        "A1" => Ok(formula_engine::ReferenceStyle::A1),
        "R1C1" => Ok(formula_engine::ReferenceStyle::R1C1),
        other => Err(js_err(format!(
            "invalid referenceStyle: {other}. Expected \"A1\" or \"R1C1\""
        ))),
    }
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ParseOptionsJsDto {
    #[serde(default)]
    locale_id: Option<String>,
    #[serde(default)]
    reference_style: Option<formula_engine::ReferenceStyle>,
}
fn parse_options_from_js(options: Option<JsValue>) -> Result<ParseOptions, JsValue> {
    let Some(value) = options else {
        return Ok(ParseOptions::default());
    };
    if value.is_undefined() || value.is_null() {
        return Ok(ParseOptions::default());
    }

    // Prefer a small JS-friendly options object. This keeps callers from having to construct
    // `formula_engine::ParseOptions` directly in JS.
    //
    // Supported shape:
    //   { localeId?: string, referenceStyle?: "A1" | "R1C1" }
    //
    // For backward compatibility, also accept a fully-serialized `ParseOptions`.
    let obj = value
        .dyn_into::<Object>()
        .map_err(|_| js_err("options must be an object".to_string()))?;
    let keys = js_sys::Object::keys(&obj);
    if keys.length() == 0 {
        return Ok(ParseOptions::default());
    }

    let has_locale_id = Reflect::has(&obj, &JsValue::from_str("localeId")).unwrap_or(false);
    let has_ref_style = Reflect::has(&obj, &JsValue::from_str("referenceStyle")).unwrap_or(false);
    if has_locale_id || has_ref_style {
        let dto: ParseOptionsJsDto =
            serde_wasm_bindgen::from_value(obj.into()).map_err(|err| js_err(err.to_string()))?;
        let mut opts = ParseOptions::default();
        if let Some(locale_id) = dto.locale_id {
            let locale = get_locale(&locale_id)
                .ok_or_else(|| js_err(format!("unknown localeId: {locale_id}")))?;
            opts.locale = locale.config.clone();
        }
        if let Some(style) = dto.reference_style {
            opts.reference_style = style;
        }
        return Ok(opts);
    }

    let looks_like_parse_options = Reflect::has(&obj, &JsValue::from_str("locale"))
        .unwrap_or(false)
        || Reflect::has(&obj, &JsValue::from_str("reference_style")).unwrap_or(false)
        || Reflect::has(&obj, &JsValue::from_str("normalize_relative_to")).unwrap_or(false);
    if looks_like_parse_options {
        // Fall back to the full ParseOptions struct for advanced callers.
        return serde_wasm_bindgen::from_value(obj.into()).map_err(|err| js_err(err.to_string()));
    }

    Err(js_err(
        "options must be { localeId?: string, referenceStyle?: \"A1\" | \"R1C1\" } or a ParseOptions object",
    ))
}

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
enum GoalSeekRecalcModeDto {
    SingleThreaded,
    MultiThreaded,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct GoalSeekRequestDto {
    target_cell: String,
    target_value: f64,
    changing_cell: String,
    #[serde(default)]
    sheet: Option<String>,
    #[serde(default)]
    tolerance: Option<f64>,
    #[serde(default)]
    max_iterations: Option<usize>,
    #[serde(default)]
    recalc_mode: Option<GoalSeekRecalcModeDto>,
}

fn edit_error_to_string(err: EngineEditError) -> String {
    match err {
        EngineEditError::SheetNotFound(sheet) => format!("sheet not found: {sheet}"),
        EngineEditError::InvalidCount => "invalid count".to_string(),
        EngineEditError::InvalidRange => "invalid range".to_string(),
        EngineEditError::OverlappingMove => "overlapping move".to_string(),
        EngineEditError::Engine(message) => message,
    }
}

#[cfg(target_arch = "wasm32")]
fn ensure_rust_constructors_run() {
    use std::sync::Once;

    static ONCE: Once = Once::new();
    ONCE.call_once(|| {
        // `inventory` (used by `formula-engine` for its built-in function registry)
        // relies on `.init_array` constructors on wasm. Some runtimes (notably
        // `wasm-bindgen-test`) do not automatically invoke them, which leaves the
        // function registry empty. Call the generated constructor trampoline when
        // needed so spreadsheet functions like `SUM()` work under wasm.
        //
        // Note: some runtimes can leave the registry *partially* populated. Avoid checking
        // "is there any function?" and instead probe for a small set of representative built-ins.
        let mut has_sum = false;
        let mut has_sequence = false;
        for spec in formula_engine::functions::iter_function_specs() {
            match spec.name {
                "SUM" => has_sum = true,
                "SEQUENCE" => has_sequence = true,
                _ => {}
            }
            if has_sum && has_sequence {
                return;
            }
        }

        extern "C" {
            fn __wasm_call_ctors();
        }

        // SAFETY: `__wasm_call_ctors` is generated by the Rust/Wasm toolchain to run global
        // constructors. This is required for `inventory`-style registries (used by `formula-engine`)
        // to be populated under wasm-bindgen-test.
        unsafe { __wasm_call_ctors() }

        let mut has_sum = false;
        let mut has_sequence = false;
        for spec in formula_engine::functions::iter_function_specs() {
            match spec.name {
                "SUM" => has_sum = true,
                "SEQUENCE" => has_sequence = true,
                _ => {}
            }
        }
        debug_assert!(
            has_sum && has_sequence,
            "formula-engine inventory registry did not populate after calling __wasm_call_ctors"
        );
    });
}

#[cfg(not(target_arch = "wasm32"))]
fn ensure_rust_constructors_run() {}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen(start)]
pub fn wasm_start() {
    ensure_rust_constructors_run();
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
pub struct Utf16Span {
    pub start: u32,
    pub end: u32,
}

#[derive(Clone, Debug)]
struct Utf16IndexMap {
    /// Monotonic mapping from UTF-8 byte offsets (Rust) to UTF-16 code-unit offsets (JS).
    ///
    /// Contains `(0, 0)` and `(s.len(), s.encode_utf16().count())`, plus an entry at every UTF-8
    /// character boundary.
    byte_to_utf16: Vec<(usize, usize)>,
}

impl Utf16IndexMap {
    fn new(s: &str) -> Self {
        let mut byte_to_utf16 = Vec::with_capacity(s.chars().count() + 2);
        byte_to_utf16.push((0, 0));
        let mut utf16: usize = 0;
        for (byte_idx, ch) in s.char_indices() {
            if byte_idx != 0 {
                byte_to_utf16.push((byte_idx, utf16));
            }
            utf16 = utf16.saturating_add(ch.len_utf16());
        }
        byte_to_utf16.push((s.len(), utf16));
        Self { byte_to_utf16 }
    }

    fn byte_to_utf16(&self, byte_offset: usize) -> usize {
        match self
            .byte_to_utf16
            .binary_search_by_key(&byte_offset, |(byte, _)| *byte)
        {
            Ok(idx) => self.byte_to_utf16[idx].1,
            Err(idx) => {
                // Token spans should always land on UTF-8 boundaries, but prefer a best-effort
                // fallback rather than panicking in production.
                if idx == 0 {
                    0
                } else {
                    self.byte_to_utf16[idx - 1].1
                }
            }
        }
    }
}

fn engine_span_to_utf16(span: EngineSpan, utf16_map: &Utf16IndexMap) -> Utf16Span {
    Utf16Span {
        start: utf16_map.byte_to_utf16(span.start) as u32,
        end: utf16_map.byte_to_utf16(span.end) as u32,
    }
}

fn add_byte_offset(span: EngineSpan, delta: usize) -> EngineSpan {
    EngineSpan {
        start: span.start.saturating_add(delta),
        end: span.end.saturating_add(delta),
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "kind")]
enum LexTokenDto {
    Number {
        span: Utf16Span,
        value: String,
    },
    String {
        span: Utf16Span,
        value: String,
    },
    Boolean {
        span: Utf16Span,
        value: bool,
    },
    Error {
        span: Utf16Span,
        value: String,
    },
    Cell {
        span: Utf16Span,
        row: u32,
        col: u32,
        row_abs: bool,
        col_abs: bool,
    },
    R1C1Cell {
        span: Utf16Span,
        row: CoordDto,
        col: CoordDto,
    },
    R1C1Row {
        span: Utf16Span,
        row: CoordDto,
    },
    R1C1Col {
        span: Utf16Span,
        col: CoordDto,
    },
    Ident {
        span: Utf16Span,
        value: String,
    },
    QuotedIdent {
        span: Utf16Span,
        value: String,
    },
    Whitespace {
        span: Utf16Span,
        value: String,
    },
    Intersect {
        span: Utf16Span,
        value: String,
    },
    LParen {
        span: Utf16Span,
    },
    RParen {
        span: Utf16Span,
    },
    LBrace {
        span: Utf16Span,
    },
    RBrace {
        span: Utf16Span,
    },
    LBracket {
        span: Utf16Span,
    },
    RBracket {
        span: Utf16Span,
    },
    Bang {
        span: Utf16Span,
    },
    Colon {
        span: Utf16Span,
    },
    Dot {
        span: Utf16Span,
    },
    ArgSep {
        span: Utf16Span,
    },
    Union {
        span: Utf16Span,
    },
    ArrayRowSep {
        span: Utf16Span,
    },
    ArrayColSep {
        span: Utf16Span,
    },
    Plus {
        span: Utf16Span,
    },
    Minus {
        span: Utf16Span,
    },
    Star {
        span: Utf16Span,
    },
    Slash {
        span: Utf16Span,
    },
    Caret {
        span: Utf16Span,
    },
    Amp {
        span: Utf16Span,
    },
    Percent {
        span: Utf16Span,
    },
    Hash {
        span: Utf16Span,
    },
    Eq {
        span: Utf16Span,
    },
    Ne {
        span: Utf16Span,
    },
    Lt {
        span: Utf16Span,
    },
    Gt {
        span: Utf16Span,
    },
    Le {
        span: Utf16Span,
    },
    Ge {
        span: Utf16Span,
    },
    At {
        span: Utf16Span,
    },
    Eof {
        span: Utf16Span,
    },
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(tag = "kind")]
enum CoordDto {
    A1 { index: u32, abs: bool },
    Offset { delta: i32 },
}

impl From<Coord> for CoordDto {
    fn from(coord: Coord) -> Self {
        match coord {
            Coord::A1 { index, abs } => CoordDto::A1 { index, abs },
            Coord::Offset(delta) => CoordDto::Offset { delta },
        }
    }
}

fn token_to_dto(token: Token, byte_offset: usize, utf16_map: &Utf16IndexMap) -> LexTokenDto {
    let span = engine_span_to_utf16(add_byte_offset(token.span, byte_offset), utf16_map);
    match token.kind {
        TokenKind::Number(raw) => LexTokenDto::Number { span, value: raw },
        TokenKind::String(value) => LexTokenDto::String { span, value },
        TokenKind::Boolean(value) => LexTokenDto::Boolean { span, value },
        TokenKind::Error(value) => LexTokenDto::Error { span, value },
        TokenKind::Cell(cell) => LexTokenDto::Cell {
            span,
            row: cell.row,
            col: cell.col,
            row_abs: cell.row_abs,
            col_abs: cell.col_abs,
        },
        TokenKind::R1C1Cell(cell) => LexTokenDto::R1C1Cell {
            span,
            row: cell.row.into(),
            col: cell.col.into(),
        },
        TokenKind::R1C1Row(row) => LexTokenDto::R1C1Row {
            span,
            row: row.row.into(),
        },
        TokenKind::R1C1Col(col) => LexTokenDto::R1C1Col {
            span,
            col: col.col.into(),
        },
        TokenKind::Ident(value) => LexTokenDto::Ident { span, value },
        TokenKind::QuotedIdent(value) => LexTokenDto::QuotedIdent { span, value },
        TokenKind::Whitespace(value) => LexTokenDto::Whitespace { span, value },
        TokenKind::Intersect(value) => LexTokenDto::Intersect { span, value },
        TokenKind::LParen => LexTokenDto::LParen { span },
        TokenKind::RParen => LexTokenDto::RParen { span },
        TokenKind::LBrace => LexTokenDto::LBrace { span },
        TokenKind::RBrace => LexTokenDto::RBrace { span },
        TokenKind::LBracket => LexTokenDto::LBracket { span },
        TokenKind::RBracket => LexTokenDto::RBracket { span },
        TokenKind::Bang => LexTokenDto::Bang { span },
        TokenKind::Colon => LexTokenDto::Colon { span },
        TokenKind::Dot => LexTokenDto::Dot { span },
        TokenKind::ArgSep => LexTokenDto::ArgSep { span },
        TokenKind::Union => LexTokenDto::Union { span },
        TokenKind::ArrayRowSep => LexTokenDto::ArrayRowSep { span },
        TokenKind::ArrayColSep => LexTokenDto::ArrayColSep { span },
        TokenKind::Plus => LexTokenDto::Plus { span },
        TokenKind::Minus => LexTokenDto::Minus { span },
        TokenKind::Star => LexTokenDto::Star { span },
        TokenKind::Slash => LexTokenDto::Slash { span },
        TokenKind::Caret => LexTokenDto::Caret { span },
        TokenKind::Amp => LexTokenDto::Amp { span },
        TokenKind::Percent => LexTokenDto::Percent { span },
        TokenKind::Hash => LexTokenDto::Hash { span },
        TokenKind::Eq => LexTokenDto::Eq { span },
        TokenKind::Ne => LexTokenDto::Ne { span },
        TokenKind::Lt => LexTokenDto::Lt { span },
        TokenKind::Gt => LexTokenDto::Gt { span },
        TokenKind::Le => LexTokenDto::Le { span },
        TokenKind::Ge => LexTokenDto::Ge { span },
        TokenKind::At => LexTokenDto::At { span },
        TokenKind::Eof => LexTokenDto::Eof { span },
    }
}

#[wasm_bindgen(js_name = "lexFormula")]
pub fn lex_formula(formula: &str, opts: Option<JsValue>) -> Result<JsValue, JsValue> {
    // `parseFormulaPartial`/`lexFormula` can be used without instantiating a workbook. Ensure the
    // function registry constructors ran for wasm-bindgen-test environments.
    ensure_rust_constructors_run();

    let opts = parse_options_from_js(opts)?;
    let (expr_src, byte_offset) = if let Some(rest) = formula.strip_prefix('=') {
        (rest, 1usize)
    } else {
        (formula, 0usize)
    };

    let utf16_map = Utf16IndexMap::new(formula);

    let tokens = formula_engine::lex(expr_src, &opts).map_err(|err| js_err(err.to_string()))?;
    let out: Vec<LexTokenDto> = tokens
        .into_iter()
        .map(|tok| token_to_dto(tok, byte_offset, &utf16_map))
        .collect();

    serde_wasm_bindgen::to_value(&out).map_err(|err| js_err(err.to_string()))
}

#[derive(Debug, Serialize)]
struct WasmLexError {
    message: String,
    span: Utf16Span,
}

#[derive(Debug, Serialize)]
struct WasmPartialLex {
    tokens: Vec<LexTokenDto>,
    error: Option<WasmLexError>,
}

/// Best-effort lexer used for editor syntax highlighting.
///
/// This mirrors `lexFormula` but never throws: on errors it returns the tokens produced so far plus
/// the first encountered lexer error.
#[wasm_bindgen(js_name = "lexFormulaPartial")]
pub fn lex_formula_partial(formula: &str, opts: Option<JsValue>) -> JsValue {
    // `parseFormulaPartial`/`lexFormula` can be used without instantiating a workbook. Ensure the
    // function registry constructors ran for wasm-bindgen-test environments.
    ensure_rust_constructors_run();

    // Best-effort: treat option parsing failures as "use defaults" so this API never throws.
    let opts = parse_options_from_js(opts).unwrap_or_default();

    let (expr_src, byte_offset) = if let Some(rest) = formula.strip_prefix('=') {
        (rest, 1usize)
    } else {
        (formula, 0usize)
    };

    let utf16_map = Utf16IndexMap::new(formula);
    let partial = formula_engine::lex_partial(expr_src, &opts);

    let tokens: Vec<LexTokenDto> = partial
        .tokens
        .into_iter()
        .map(|tok| token_to_dto(tok, byte_offset, &utf16_map))
        .collect();

    let error = partial.error.map(|err| WasmLexError {
        message: err.message,
        span: engine_span_to_utf16(add_byte_offset(err.span, byte_offset), &utf16_map),
    });

    let out = WasmPartialLex { tokens, error };
    use serde::ser::Serialize as _;
    out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
        .unwrap_or_else(|err| js_err(err.to_string()))
}

/// Canonicalize a localized formula into the engine's persisted form.
///
/// Canonical form uses:
/// - English function names (e.g. `SUM`)
/// - `,` as argument separator
/// - `.` as decimal separator
///
/// `referenceStyle` controls how cell references are tokenized (`A1` vs `R1C1`).
#[wasm_bindgen(js_name = "canonicalizeFormula")]
pub fn canonicalize_formula(
    formula: &str,
    locale_id: &str,
    reference_style: Option<String>,
) -> Result<String, JsValue> {
    ensure_rust_constructors_run();
    let locale = require_formula_locale(locale_id)?;
    let reference_style = parse_reference_style(reference_style)?;
    canonicalize_formula_with_style(formula, locale, reference_style)
        .map_err(|err| js_err(err.to_string()))
}

/// Localize a canonical (English) formula into a locale-specific display form.
///
/// `referenceStyle` controls how cell references are tokenized (`A1` vs `R1C1`).
#[wasm_bindgen(js_name = "localizeFormula")]
pub fn localize_formula(
    formula: &str,
    locale_id: &str,
    reference_style: Option<String>,
) -> Result<String, JsValue> {
    ensure_rust_constructors_run();
    let locale = require_formula_locale(locale_id)?;
    let reference_style = parse_reference_style(reference_style)?;
    localize_formula_with_style(formula, locale, reference_style)
        .map_err(|err| js_err(err.to_string()))
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RewriteFormulaForCopyDeltaRequestDto {
    formula: String,
    delta_row: i32,
    delta_col: i32,
}

/// Rewrite a batch of formulas as if they were copied by `(deltaRow, deltaCol)`.
///
/// This is used by UI layers (clipboard paste, fill handle) that need the engine's formula
/// shifting semantics without mutating workbook state.
#[wasm_bindgen(js_name = "rewriteFormulasForCopyDelta")]
pub fn rewrite_formulas_for_copy_delta(requests: JsValue) -> Result<JsValue, JsValue> {
    ensure_rust_constructors_run();
    let requests: Vec<RewriteFormulaForCopyDeltaRequestDto> =
        serde_wasm_bindgen::from_value(requests).map_err(|err| js_err(err.to_string()))?;

    let origin = CellAddr::new(0, 0);
    let mut out: Vec<String> = Vec::with_capacity(requests.len());
    for req in requests {
        let (rewritten, _) = rewrite_formula_for_copy_delta(
            &req.formula,
            DEFAULT_SHEET,
            origin,
            req.delta_row,
            req.delta_col,
        );
        out.push(rewritten);
    }

    serde_wasm_bindgen::to_value(&out).map_err(|err| js_err(err.to_string()))
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct FormulaCellKey {
    sheet: String,
    row: u32,
    col: u32,
}

impl FormulaCellKey {
    fn new(sheet: String, cell: CellRef) -> Self {
        Self {
            sheet,
            row: cell.row,
            col: cell.col,
        }
    }

    fn address(&self) -> String {
        CellRef::new(self.row, self.col).to_a1()
    }
}

fn is_scalar_json(value: &JsonValue) -> bool {
    matches!(
        value,
        JsonValue::Null | JsonValue::Bool(_) | JsonValue::Number(_) | JsonValue::String(_)
    )
}

fn is_formula_input(value: &JsonValue) -> bool {
    value.as_str().is_some_and(|s| {
        let trimmed = s.trim_start();
        let Some(rest) = trimmed.strip_prefix('=') else {
            return false;
        };
        !rest.trim().is_empty()
    })
}

fn normalize_sheet_key(name: &str) -> String {
    name.to_ascii_uppercase()
}

/// Encode a literal text string as a scalar workbook `input` value.
///
/// The legacy JS worker protocol treats strings that look like formulas (leading `=`, ignoring
/// whitespace) and error codes (e.g. `#REF!`) as structured inputs. To preserve non-formula rich
/// inputs through `toJson`/`fromJson` round-trips we apply Excel's quote prefix (`'`) when needed.
fn encode_scalar_text_input(text: &str) -> String {
    // If the desired text itself starts with a quote prefix, double it so the scalar path keeps a
    // leading apostrophe after `json_to_engine_value` strips one.
    if text.starts_with('\'') {
        return format!("'{text}");
    }

    let candidate = JsonValue::String(text.to_string());
    if is_formula_input(&candidate) || ErrorKind::from_code(text).is_some() {
        format!("'{text}")
    } else {
        text.to_string()
    }
}
fn json_to_engine_value(value: &JsonValue) -> EngineValue {
    match value {
        JsonValue::Null => EngineValue::Blank,
        JsonValue::Bool(b) => EngineValue::Bool(*b),
        JsonValue::Number(n) => EngineValue::Number(n.as_f64().unwrap_or(0.0)),
        JsonValue::String(s) => {
            // Excel-style quote prefix: a leading apostrophe forces the value to be treated as
            // literal text (even if it looks like an error code or formula).
            if let Some(rest) = s.strip_prefix('\'') {
                return EngineValue::Text(rest.to_string());
            }

            if let Some(kind) = ErrorKind::from_code(s) {
                return EngineValue::Error(kind);
            }

            EngineValue::Text(s.clone())
        }
        JsonValue::Array(_) | JsonValue::Object(_) => {
            // Should be unreachable due to `is_scalar_json` validation, but keep a fallback.
            EngineValue::Blank
        }
    }
}

fn engine_value_to_json(value: EngineValue) -> JsonValue {
    match value {
        EngineValue::Blank => JsonValue::Null,
        EngineValue::Bool(b) => JsonValue::Bool(b),
        EngineValue::Text(s) => JsonValue::String(s),
        EngineValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        EngineValue::Entity(entity) => JsonValue::String(entity.display),
        EngineValue::Record(record) => JsonValue::String(record.display),
        EngineValue::Error(kind) => JsonValue::String(kind.as_code().to_string()),
        // Arrays should generally be spilled into grid cells. If one reaches the JS boundary,
        // degrade to its top-left value so callers still get a scalar.
        EngineValue::Array(arr) => engine_value_to_json(arr.top_left()),
        // The JS worker protocol only supports scalar-ish values today.
        //
        // Degrade any rich/non-scalar value (references, lambdas, entities, records, etc.) to its
        // display string so existing `getCell` / `recalculate` callers keep receiving scalars.
        other => JsonValue::String(other.to_string()),
    }
}

fn engine_value_to_pivot_value(value: EngineValue) -> pivot_engine::PivotValue {
    match value {
        EngineValue::Blank => pivot_engine::PivotValue::Blank,
        EngineValue::Bool(b) => pivot_engine::PivotValue::Bool(b),
        EngineValue::Number(n) => pivot_engine::PivotValue::Number(n),
        EngineValue::Text(s) => pivot_engine::PivotValue::Text(s),
        EngineValue::Entity(entity) => pivot_engine::PivotValue::Text(entity.display),
        EngineValue::Record(record) => pivot_engine::PivotValue::Text(record.display),
        EngineValue::Error(kind) => pivot_engine::PivotValue::Text(kind.as_code().to_string()),
        // Arrays should generally be spilled into grid cells. If one reaches pivot extraction,
        // degrade to its top-left value for a scalar-like experience.
        EngineValue::Array(arr) => engine_value_to_pivot_value(arr.top_left()),
        // Degrade any other rich/non-scalar value (references, lambdas, spill markers, etc.) to
        // its display string so the pivot engine can treat it as a grouping key.
        other => pivot_engine::PivotValue::Text(other.to_string()),
    }
}

fn pivot_value_to_json(value: pivot_engine::PivotValue) -> JsonValue {
    match value {
        pivot_engine::PivotValue::Blank => JsonValue::Null,
        pivot_engine::PivotValue::Bool(b) => JsonValue::Bool(b),
        pivot_engine::PivotValue::Text(s) => JsonValue::String(s),
        pivot_engine::PivotValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        // The scalar WASM protocol has no first-class date type today; represent dates using their
        // ISO string form (YYYY-MM-DD).
        pivot_engine::PivotValue::Date(d) => JsonValue::String(d.to_string()),
    }
}

fn pivot_key_part_model_to_engine(part: &formula_model::pivots::PivotKeyPart) -> pivot_engine::PivotKeyPart {
    match part {
        formula_model::pivots::PivotKeyPart::Blank => pivot_engine::PivotKeyPart::Blank,
        formula_model::pivots::PivotKeyPart::Number(bits) => pivot_engine::PivotKeyPart::Number(*bits),
        formula_model::pivots::PivotKeyPart::Date(d) => pivot_engine::PivotKeyPart::Date(*d),
        formula_model::pivots::PivotKeyPart::Text(s) => pivot_engine::PivotKeyPart::Text(s.clone()),
        formula_model::pivots::PivotKeyPart::Bool(b) => pivot_engine::PivotKeyPart::Bool(*b),
    }
}

fn pivot_sort_order_model_to_engine(
    order: formula_model::pivots::SortOrder,
) -> pivot_engine::SortOrder {
    match order {
        formula_model::pivots::SortOrder::Ascending => pivot_engine::SortOrder::Ascending,
        formula_model::pivots::SortOrder::Descending => pivot_engine::SortOrder::Descending,
        formula_model::pivots::SortOrder::Manual => pivot_engine::SortOrder::Manual,
    }
}

fn pivot_field_model_to_engine(field: &formula_model::pivots::PivotField) -> pivot_engine::PivotField {
    pivot_engine::PivotField {
        source_field: field.source_field.clone(),
        sort_order: pivot_sort_order_model_to_engine(field.sort_order),
        manual_sort: field
            .manual_sort
            .as_ref()
            .map(|items| items.iter().map(pivot_key_part_model_to_engine).collect()),
    }
}

fn pivot_aggregation_model_to_engine(
    agg: formula_model::pivots::AggregationType,
) -> pivot_engine::AggregationType {
    match agg {
        formula_model::pivots::AggregationType::Sum => pivot_engine::AggregationType::Sum,
        formula_model::pivots::AggregationType::Count => pivot_engine::AggregationType::Count,
        formula_model::pivots::AggregationType::Average => pivot_engine::AggregationType::Average,
        formula_model::pivots::AggregationType::Max => pivot_engine::AggregationType::Max,
        formula_model::pivots::AggregationType::Min => pivot_engine::AggregationType::Min,
        formula_model::pivots::AggregationType::Product => pivot_engine::AggregationType::Product,
        formula_model::pivots::AggregationType::CountNumbers => {
            pivot_engine::AggregationType::CountNumbers
        }
        formula_model::pivots::AggregationType::StdDev => pivot_engine::AggregationType::StdDev,
        formula_model::pivots::AggregationType::StdDevP => pivot_engine::AggregationType::StdDevP,
        formula_model::pivots::AggregationType::Var => pivot_engine::AggregationType::Var,
        formula_model::pivots::AggregationType::VarP => pivot_engine::AggregationType::VarP,
    }
}

fn pivot_value_field_model_to_engine(field: &formula_model::pivots::ValueField) -> pivot_engine::ValueField {
    pivot_engine::ValueField {
        source_field: field.source_field.clone(),
        name: field.name.clone(),
        aggregation: pivot_aggregation_model_to_engine(field.aggregation),
        show_as: field.show_as,
        base_field: field.base_field.clone(),
        base_item: field.base_item.clone(),
    }
}

fn pivot_filter_field_model_to_engine(field: &formula_model::pivots::FilterField) -> pivot_engine::FilterField {
    pivot_engine::FilterField {
        source_field: field.source_field.clone(),
        allowed: field.allowed.as_ref().map(|allowed| {
            allowed
                .iter()
                .map(pivot_key_part_model_to_engine)
                .collect::<std::collections::HashSet<_>>()
        }),
    }
}

fn pivot_layout_model_to_engine(layout: formula_model::pivots::Layout) -> pivot_engine::Layout {
    match layout {
        formula_model::pivots::Layout::Compact => pivot_engine::Layout::Compact,
        // `Outline` is not yet supported by the pivot engine; treat it as tabular output.
        formula_model::pivots::Layout::Outline | formula_model::pivots::Layout::Tabular => {
            pivot_engine::Layout::Tabular
        }
    }
}

fn pivot_subtotals_model_to_engine(
    position: formula_model::pivots::SubtotalPosition,
) -> pivot_engine::SubtotalPosition {
    match position {
        formula_model::pivots::SubtotalPosition::Top => pivot_engine::SubtotalPosition::Top,
        formula_model::pivots::SubtotalPosition::Bottom => pivot_engine::SubtotalPosition::Bottom,
        formula_model::pivots::SubtotalPosition::None => pivot_engine::SubtotalPosition::None,
    }
}

fn pivot_config_model_to_engine(cfg: &formula_model::pivots::PivotConfig) -> pivot_engine::PivotConfig {
    pivot_engine::PivotConfig {
        row_fields: cfg.row_fields.iter().map(pivot_field_model_to_engine).collect(),
        column_fields: cfg
            .column_fields
            .iter()
            .map(pivot_field_model_to_engine)
            .collect(),
        value_fields: cfg
            .value_fields
            .iter()
            .map(pivot_value_field_model_to_engine)
            .collect(),
        filter_fields: cfg
            .filter_fields
            .iter()
            .map(pivot_filter_field_model_to_engine)
            .collect(),
        calculated_fields: cfg
            .calculated_fields
            .iter()
            .map(|f| pivot_engine::CalculatedField {
                name: f.name.clone(),
                formula: f.formula.clone(),
            })
            .collect(),
        calculated_items: cfg
            .calculated_items
            .iter()
            .map(|it| pivot_engine::CalculatedItem {
                field: it.field.clone(),
                name: it.name.clone(),
                formula: it.formula.clone(),
            })
            .collect(),
        layout: pivot_layout_model_to_engine(cfg.layout),
        subtotals: pivot_subtotals_model_to_engine(cfg.subtotals),
        grand_totals: pivot_engine::GrandTotals {
            rows: cfg.grand_totals.rows,
            columns: cfg.grand_totals.columns,
        },
    }
}

/// Convert an engine value into a scalar workbook `input` representation.
///
/// This differs from [`engine_value_to_json`] for text values: workbook inputs must preserve
/// "quote prefix" escaping so strings that look like formulas/errors (or begin with an
/// apostrophe) survive `toJson`/`fromJson` round-trips without changing semantics.
///
/// Returns `None` for values that cannot be represented in the legacy scalar input map (e.g.
/// rich values like entities/records, lambdas, references). Callers should treat `None` as
/// "remove from the sparse input map".
fn engine_value_to_scalar_json_input(value: EngineValue) -> Option<JsonValue> {
    match value {
        EngineValue::Blank => None,
        EngineValue::Bool(b) => Some(JsonValue::Bool(b)),
        EngineValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .or_else(|| Some(JsonValue::String(ErrorKind::Num.as_code().to_string()))),
        EngineValue::Text(s) => Some(JsonValue::String(encode_scalar_text_input(&s))),
        EngineValue::Error(kind) => Some(JsonValue::String(kind.as_code().to_string())),
        EngineValue::Array(arr) => engine_value_to_scalar_json_input(arr.top_left()),
        // Rich/non-scalar values are not representable in the scalar input map.
        _ => None,
    }
}
fn model_error_to_engine(err: formula_model::ErrorValue) -> ErrorKind {
    match err {
        formula_model::ErrorValue::Null => ErrorKind::Null,
        formula_model::ErrorValue::Div0 => ErrorKind::Div0,
        formula_model::ErrorValue::Value => ErrorKind::Value,
        formula_model::ErrorValue::Ref => ErrorKind::Ref,
        formula_model::ErrorValue::Name => ErrorKind::Name,
        formula_model::ErrorValue::Num => ErrorKind::Num,
        formula_model::ErrorValue::NA => ErrorKind::NA,
        formula_model::ErrorValue::GettingData => ErrorKind::GettingData,
        formula_model::ErrorValue::Spill => ErrorKind::Spill,
        formula_model::ErrorValue::Calc => ErrorKind::Calc,
        formula_model::ErrorValue::Field => ErrorKind::Field,
        formula_model::ErrorValue::Connect => ErrorKind::Connect,
        formula_model::ErrorValue::Blocked => ErrorKind::Blocked,
        formula_model::ErrorValue::Unknown => ErrorKind::Unknown,
    }
}

fn engine_error_to_model(kind: ErrorKind) -> formula_model::ErrorValue {
    match kind {
        ErrorKind::Null => formula_model::ErrorValue::Null,
        ErrorKind::Div0 => formula_model::ErrorValue::Div0,
        ErrorKind::Value => formula_model::ErrorValue::Value,
        ErrorKind::Ref => formula_model::ErrorValue::Ref,
        ErrorKind::Name => formula_model::ErrorValue::Name,
        ErrorKind::Num => formula_model::ErrorValue::Num,
        ErrorKind::NA => formula_model::ErrorValue::NA,
        ErrorKind::GettingData => formula_model::ErrorValue::GettingData,
        ErrorKind::Spill => formula_model::ErrorValue::Spill,
        ErrorKind::Calc => formula_model::ErrorValue::Calc,
        ErrorKind::Field => formula_model::ErrorValue::Field,
        ErrorKind::Connect => formula_model::ErrorValue::Connect,
        ErrorKind::Blocked => formula_model::ErrorValue::Blocked,
        ErrorKind::Unknown => formula_model::ErrorValue::Unknown,
    }
}

fn scalar_json_to_cell_value_input(value: &JsonValue) -> CellValue {
    match value {
        JsonValue::Null => CellValue::Empty,
        JsonValue::Bool(b) => CellValue::Boolean(*b),
        JsonValue::Number(n) => CellValue::Number(n.as_f64().unwrap_or(0.0)),
        JsonValue::String(s) => {
            // Excel-style quote prefix: a leading apostrophe forces literal text.
            if let Some(rest) = s.strip_prefix('\'') {
                return CellValue::String(rest.to_string());
            }
            if let Some(kind) = ErrorKind::from_code(s) {
                return CellValue::Error(engine_error_to_model(kind));
            }
            CellValue::String(s.clone())
        }
        JsonValue::Array(_) | JsonValue::Object(_) => CellValue::Empty,
    }
}

fn engine_value_to_cell_value_rich(value: EngineValue) -> CellValue {
    match value {
        EngineValue::Blank => CellValue::Empty,
        EngineValue::Bool(b) => CellValue::Boolean(b),
        EngineValue::Number(n) => CellValue::Number(n),
        EngineValue::Text(s) => CellValue::String(s),
        EngineValue::Error(kind) => CellValue::Error(engine_error_to_model(kind)),
        EngineValue::Entity(entity) => {
            let mut properties = BTreeMap::new();
            for (k, v) in entity.fields {
                properties.insert(k, engine_value_to_cell_value_rich(v));
            }
            CellValue::Entity(formula_model::EntityValue {
                entity_type: entity.entity_type.unwrap_or_default(),
                entity_id: entity.entity_id.unwrap_or_default(),
                display_value: entity.display,
                properties,
            })
        }
        EngineValue::Record(record) => {
            let mut fields = BTreeMap::new();
            for (k, v) in record.fields {
                fields.insert(k, engine_value_to_cell_value_rich(v));
            }
            CellValue::Record(formula_model::RecordValue {
                fields,
                display_field: record.display_field,
                display_value: record.display,
            })
        }
        EngineValue::Array(arr) => {
            let mut iter = arr.values.into_iter();
            let mut data = Vec::with_capacity(arr.rows);
            for _ in 0..arr.rows {
                let mut row = Vec::with_capacity(arr.cols);
                for _ in 0..arr.cols {
                    let next = iter.next().unwrap_or(EngineValue::Blank);
                    row.push(engine_value_to_cell_value_rich(next));
                }
                data.push(row);
            }
            CellValue::Array(formula_model::ArrayValue { data })
        }
        EngineValue::Spill { origin } => CellValue::Spill(formula_model::SpillValue { origin }),
        other => CellValue::String(other.to_string()),
    }
}

/// Convert a `formula-model` [`CellValue`] (including entity/record rich values) into a
/// `formula-engine` runtime [`Value`](formula_engine::Value).
///
/// Arrays are mapped into engine arrays (2D row-major). Note that the scalar JS-facing protocol
/// (`getCell`/`recalculate`) still degrades arrays to their top-left value.
fn cell_value_to_engine_rich(value: &CellValue) -> Result<EngineValue, JsValue> {
    match value {
        CellValue::Empty => Ok(EngineValue::Blank),
        CellValue::Number(n) => Ok(EngineValue::Number(*n)),
        CellValue::String(s) => Ok(EngineValue::Text(s.clone())),
        CellValue::Boolean(b) => Ok(EngineValue::Bool(*b)),
        CellValue::Error(err) => Ok(EngineValue::Error(model_error_to_engine(*err))),
        CellValue::RichText(rt) => Ok(EngineValue::Text(rt.plain_text().to_string())),
        CellValue::Entity(entity) => {
            let mut fields = HashMap::new();
            for (k, v) in &entity.properties {
                fields.insert(k.clone(), cell_value_to_engine_rich(v)?);
            }
            Ok(EngineValue::Entity(formula_engine::value::EntityValue {
                display: entity.display_value.clone(),
                entity_type: (!entity.entity_type.is_empty()).then(|| entity.entity_type.clone()),
                entity_id: (!entity.entity_id.is_empty()).then(|| entity.entity_id.clone()),
                fields,
            }))
        }
        CellValue::Record(record) => {
            let mut fields = HashMap::new();
            for (k, v) in &record.fields {
                fields.insert(k.clone(), cell_value_to_engine_rich(v)?);
            }
            Ok(EngineValue::Record(formula_engine::value::RecordValue {
                display: record.to_string(),
                display_field: record.display_field.clone(),
                fields,
            }))
        }
        CellValue::Image(image) => Ok(EngineValue::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        )),
        CellValue::Array(arr) => {
            let rows = arr.data.len();
            let cols = arr.data.first().map(|r| r.len()).unwrap_or(0);
            if arr.data.iter().any(|r| r.len() != cols) {
                return Err(js_err(
                    "invalid array CellValue: expected a rectangular 2D array",
                ));
            }

            let mut values = Vec::with_capacity(rows.saturating_mul(cols));
            for row in &arr.data {
                for v in row {
                    values.push(cell_value_to_engine_rich(v)?);
                }
            }

            Ok(EngineValue::Array(formula_engine::value::Array::new(
                rows, cols, values,
            )))
        }
        CellValue::Spill(spill) => Ok(EngineValue::Spill {
            origin: spill.origin,
        }),
    }
}

fn cell_value_to_engine(value: &CellValue) -> EngineValue {
    match value {
        CellValue::Empty => EngineValue::Blank,
        CellValue::Number(n) => EngineValue::Number(*n),
        CellValue::String(s) => EngineValue::Text(s.clone()),
        CellValue::Boolean(b) => EngineValue::Bool(*b),
        CellValue::Error(err) => EngineValue::Error(model_error_to_engine(*err)),
        CellValue::RichText(rt) => EngineValue::Text(rt.plain_text().to_string()),
        CellValue::Entity(_) | CellValue::Record(_) => cell_value_to_engine_rich(value)
            .unwrap_or_else(|_| EngineValue::Error(ErrorKind::Value)),
        CellValue::Image(image) => EngineValue::Text(
            image
                .alt_text
                .clone()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "[Image]".to_string()),
        ),
        // The workbook model can store cached array/spill results, but the WASM worker API only
        // supports scalar values today. Treat these as spill errors so downstream formulas see an
        // error rather than silently treating an array as a string.
        CellValue::Array(_) | CellValue::Spill(_) => EngineValue::Error(ErrorKind::Spill),
    }
}

fn cell_value_to_scalar_json_input(value: &CellValue) -> JsonValue {
    match value {
        CellValue::Empty => JsonValue::Null,
        CellValue::Number(n) => serde_json::Number::from_f64(*n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        CellValue::Boolean(b) => JsonValue::Bool(*b),
        CellValue::String(s) => JsonValue::String(encode_scalar_text_input(s)),
        CellValue::Error(err) => JsonValue::String(err.as_str().to_string()),
        CellValue::RichText(rt) => {
            JsonValue::String(encode_scalar_text_input(&rt.plain_text().to_string()))
        }
        CellValue::Entity(entity) => {
            JsonValue::String(encode_scalar_text_input(&entity.display_value))
        }
        CellValue::Record(record) => {
            let display = record.to_string();
            JsonValue::String(encode_scalar_text_input(&display))
        }
        CellValue::Image(image) => {
            let display = image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or("[Image]");
            JsonValue::String(encode_scalar_text_input(display))
        }
        // Degrade arrays to their top-left value so `getCell`/`toJson` remain scalar-compatible.
        CellValue::Array(arr) => arr
            .data
            .first()
            .and_then(|row| row.first())
            .map(cell_value_to_scalar_json_input)
            .unwrap_or(JsonValue::Null),
        // Preserve the scalar spill error in legacy IO paths.
        CellValue::Spill(_) => JsonValue::String(ErrorKind::Spill.as_code().to_string()),
    }
}

struct WorkbookState {
    engine: Engine,
    formula_locale: &'static FormulaLocale,
    /// Workbook input state for `toJson`/`getCell.input`.
    ///
    /// Mirrors the simple JSON workbook schema consumed by `packages/engine`.
    sheets: BTreeMap<String, BTreeMap<String, JsonValue>>,
    /// Case-insensitive mapping (Excel semantics) from sheet key -> display name.
    sheet_lookup: HashMap<String, String>,
    /// Spill cells that were cleared by edits since the last recalc.
    ///
    /// `Engine::recalculate_with_value_changes` can only diff values across a recalc tick; when a
    /// spill is cleared as part of `setCell`/`setRange` we stash the affected cells so the next
    /// `recalculate()` call can return `CellChange[]` entries that blank out any now-stale spill
    /// outputs in the JS cache.
    pending_spill_clears: BTreeSet<FormulaCellKey>,
    /// Formula cells that were edited since the last recalc, keyed by their previous visible value.
    ///
    /// The JS frontend applies `directChange` updates for literal edits but not for formulas; the
    /// WASM bridge resets formula cells to blank until the next `recalculate()` so `getCell` matches
    /// the existing semantics. This can hide "value cleared" edits when the new formula result is
    /// also blank, so we keep the previous value here and explicitly diff it against the post-recalc
    /// value.
    pending_formula_baselines: BTreeMap<FormulaCellKey, JsonValue>,
    /// Rich cell input values set via `setCellRich`.
    ///
    /// This is stored separately from `sheets` to keep legacy scalar IO (`toJson`/`getCell`) stable.
    sheets_rich: BTreeMap<String, BTreeMap<String, CellValue>>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(tag = "type")]
enum EditOpDto {
    InsertRows {
        sheet: String,
        row: u32,
        count: u32,
    },
    DeleteRows {
        sheet: String,
        row: u32,
        count: u32,
    },
    InsertCols {
        sheet: String,
        col: u32,
        count: u32,
    },
    DeleteCols {
        sheet: String,
        col: u32,
        count: u32,
    },
    InsertCellsShiftRight {
        sheet: String,
        range: String,
    },
    InsertCellsShiftDown {
        sheet: String,
        range: String,
    },
    DeleteCellsShiftLeft {
        sheet: String,
        range: String,
    },
    DeleteCellsShiftUp {
        sheet: String,
        range: String,
    },
    MoveRange {
        sheet: String,
        src: String,
        #[serde(rename = "dstTopLeft")]
        dst_top_left: String,
    },
    CopyRange {
        sheet: String,
        src: String,
        #[serde(rename = "dstTopLeft")]
        dst_top_left: String,
    },
    Fill {
        sheet: String,
        src: String,
        dst: String,
    },
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct EditResultDto {
    changed_cells: Vec<EditCellChangeDto>,
    moved_ranges: Vec<EditMovedRangeDto>,
    formula_rewrites: Vec<EditFormulaRewriteDto>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct EditCellChangeDto {
    sheet: String,
    address: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    before: Option<EditCellSnapshotDto>,
    #[serde(skip_serializing_if = "Option::is_none")]
    after: Option<EditCellSnapshotDto>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct EditCellSnapshotDto {
    value: JsonValue,
    #[serde(skip_serializing_if = "Option::is_none")]
    formula: Option<String>,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct EditMovedRangeDto {
    sheet: String,
    from: String,
    to: String,
}

#[derive(Clone, Debug, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct EditFormulaRewriteDto {
    sheet: String,
    address: String,
    before: String,
    after: String,
}

impl WorkbookState {
    fn new_empty() -> Self {
        ensure_rust_constructors_run();
        Self {
            engine: Engine::new(),
            formula_locale: &EN_US,
            sheets: BTreeMap::new(),
            sheets_rich: BTreeMap::new(),
            sheet_lookup: HashMap::new(),
            pending_spill_clears: BTreeSet::new(),
            pending_formula_baselines: BTreeMap::new(),
        }
    }

    fn new_with_default_sheet() -> Self {
        let mut wb = Self::new_empty();
        wb.ensure_sheet(DEFAULT_SHEET);
        wb
    }

    fn ensure_sheet(&mut self, name: &str) -> String {
        let key = normalize_sheet_key(name);
        if let Some(existing) = self.sheet_lookup.get(&key) {
            return existing.clone();
        }

        let display = name.to_string();
        self.sheet_lookup.insert(key, display.clone());
        self.sheets.entry(display.clone()).or_default();
        self.sheets_rich.entry(display.clone()).or_default();
        self.engine.ensure_sheet(&display);
        display
    }

    fn set_sheet_dimensions_internal(
        &mut self,
        name: &str,
        rows: u32,
        cols: u32,
    ) -> Result<(), JsValue> {
        let sheet = self.ensure_sheet(name);
        self.engine
            .set_sheet_dimensions(&sheet, rows, cols)
            .map_err(|err| js_err(err.to_string()))
    }

    fn get_sheet_dimensions_internal(&self, name: &str) -> Result<(u32, u32), JsValue> {
        let sheet = self.require_sheet(name)?;
        self.engine
            .sheet_dimensions(sheet)
            .ok_or_else(|| js_err(format!("missing sheet: {name}")))
    }

    fn resolve_sheet(&self, name: &str) -> Option<&str> {
        let key = normalize_sheet_key(name);
        self.sheet_lookup.get(&key).map(String::as_str)
    }

    fn require_sheet(&self, name: &str) -> Result<&str, JsValue> {
        self.resolve_sheet(name)
            .ok_or_else(|| js_err(format!("missing sheet: {name}")))
    }

    fn parse_address(address: &str) -> Result<CellRef, JsValue> {
        CellRef::from_a1(address).map_err(|_| js_err(format!("invalid cell address: {address}")))
    }

    fn parse_range(range: &str) -> Result<Range, JsValue> {
        Range::from_a1(range).map_err(|_| js_err(format!("invalid range: {range}")))
    }

    fn read_range_values_as_pivot_values(
        &self,
        sheet: &str,
        range: &Range,
    ) -> Vec<Vec<pivot_engine::PivotValue>> {
        let mut out = Vec::with_capacity(range.height() as usize);
        for row in range.start.row..=range.end.row {
            let mut row_values = Vec::with_capacity(range.width() as usize);
            for col in range.start.col..=range.end.col {
                let address = CellRef::new(row, col).to_a1();
                let value = self.engine.get_cell_value(sheet, &address);
                row_values.push(engine_value_to_pivot_value(value));
            }
            out.push(row_values);
        }
        out
    }

    fn get_pivot_schema_internal(
        &self,
        sheet: &str,
        source_range_a1: &str,
        sample_size: usize,
    ) -> Result<pivot_engine::PivotSchema, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let range = Self::parse_range(source_range_a1)?;
        let source = self.read_range_values_as_pivot_values(&sheet, &range);
        let cache =
            pivot_engine::PivotCache::from_range(&source).map_err(|err| js_err(err.to_string()))?;
        Ok(cache.schema(sample_size))
    }

    fn calculate_pivot_writes_internal(
        &self,
        sheet: &str,
        source_range_a1: &str,
        destination_top_left_a1: &str,
        config: &pivot_engine::PivotConfig,
    ) -> Result<Vec<CellChange>, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let range = Self::parse_range(source_range_a1)?;
        let destination = Self::parse_address(destination_top_left_a1)?;

        let source = self.read_range_values_as_pivot_values(&sheet, &range);
        let cache =
            pivot_engine::PivotCache::from_range(&source).map_err(|err| js_err(err.to_string()))?;
        let result =
            pivot_engine::PivotEngine::calculate(&cache, config).map_err(|err| js_err(err.to_string()))?;

        let writes = result.to_cell_writes(pivot_engine::CellRef {
            row: destination.row,
            col: destination.col,
        });

        let mut out = Vec::with_capacity(writes.len());
        for write in writes {
            out.push(CellChange {
                sheet: sheet.clone(),
                address: CellRef::new(write.row, write.col).to_a1(),
                value: pivot_value_to_json(write.value),
            });
        }
        Ok(out)
    }
    fn set_cell_internal(
        &mut self,
        sheet: &str,
        address: &str,
        input: JsonValue,
    ) -> Result<(), JsValue> {
        if !is_scalar_json(&input) {
            return Err(js_err(format!("invalid cell value: {address}")));
        }

        let sheet = self.ensure_sheet(sheet);
        let cell_ref = Self::parse_address(address)?;
        let address = cell_ref.to_a1();

        // Legacy scalar edits overwrite any previous rich input for this cell.
        if let Some(rich_cells) = self.sheets_rich.get_mut(&sheet) {
            rich_cells.remove(&address);
        }

        if let Some((origin, end)) = self.engine.spill_range(&sheet, &address) {
            let edited_row = cell_ref.row;
            let edited_col = cell_ref.col;
            let edited_is_formula = is_formula_input(&input);
            for row in origin.row..=end.row {
                for col in origin.col..=end.col {
                    // Skip the origin cell (top-left); we only need to clear spill outputs.
                    if row == origin.row && col == origin.col {
                        continue;
                    }
                    // If the user overwrote a spill output cell with a literal value, don't emit a
                    // spill-clear change for that cell; the caller already knows its new input.
                    if !edited_is_formula && row == edited_row && col == edited_col {
                        continue;
                    }
                    self.pending_spill_clears
                        .insert(FormulaCellKey::new(sheet.clone(), CellRef::new(row, col)));
                }
            }
        }

        let sheet_cells = self
            .sheets
            .get_mut(&sheet)
            .expect("sheet just ensured must exist");

        // `null` represents an empty cell in the JS protocol. Preserve sparse semantics by
        // removing the stored entry instead of storing an explicit blank.
        if input.is_null() {
            self.engine
                .clear_cell(&sheet, &address)
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.remove(&address);
            // If this cell was previously tracked as part of a spill-clear batch, drop it so we
            // don't report direct input edits as recalc changes.
            self.pending_spill_clears
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            self.pending_formula_baselines
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            return Ok(());
        }

        if is_formula_input(&input) {
            let raw = input.as_str().expect("formula input must be string");
            // Match `formula-model`'s display semantics so the worker protocol doesn't
            // drift from other layers (trim both ends, strip a single leading '=', and
            // treat bare '=' as empty).
            let normalized = display_formula_text(raw);
            if normalized.is_empty() {
                // This should be unreachable because `is_formula_input` requires
                // non-whitespace content after '=', but keep a defensive fallback so
                // we never store a literal "=" formula.
                self.engine
                    .clear_cell(&sheet, &address)
                    .map_err(|err| js_err(err.to_string()))?;
                sheet_cells.remove(&address);
                self.pending_spill_clears
                    .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
                self.pending_formula_baselines
                    .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
                return Ok(());
            }

            let canonical = if self.formula_locale.id == EN_US.id {
                normalized
            } else {
                canonicalize_formula_with_style(
                    &normalized,
                    self.formula_locale,
                    formula_engine::ReferenceStyle::A1,
                )
                .map_err(|err| js_err(err.to_string()))?
            };

            let key = FormulaCellKey::new(sheet.clone(), cell_ref);
            self.pending_formula_baselines
                .entry(key)
                .or_insert_with(|| {
                    engine_value_to_json(self.engine.get_cell_value(&sheet, &address))
                });

            // Reset the stored value to blank so `getCell` returns null until the next recalc,
            // matching the existing worker semantics.
            self.engine
                .set_cell_value(&sheet, &address, EngineValue::Blank)
                .map_err(|err| js_err(err.to_string()))?;
            self.engine
                .set_cell_formula(&sheet, &address, &canonical)
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.insert(address.clone(), JsonValue::String(canonical));
            return Ok(());
        }

        // Non-formula scalar value.
        self.engine
            .set_cell_value(&sheet, &address, json_to_engine_value(&input))
            .map_err(|err| js_err(err.to_string()))?;

        sheet_cells.insert(address.clone(), input);
        // If this cell was previously tracked as part of a spill-clear batch (e.g. a multi-cell
        // paste over a spill range), drop it so we don't report direct input edits as recalc
        // changes.
        self.pending_spill_clears
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
        self.pending_formula_baselines
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
        Ok(())
    }

    fn set_cell_rich_internal(
        &mut self,
        sheet: &str,
        address: &str,
        input: CellValue,
    ) -> Result<(), JsValue> {
        // Preserve the legacy scalar JS worker protocol by delegating for values that can already
        // be represented as scalars. This keeps behavior consistent for numbers, booleans, strings,
        // rich text, and error values while still allowing structured rich values (entity/record,
        // images, arrays) to round-trip through `getCellRich`.
        if matches!(
            &input,
            CellValue::Empty
                | CellValue::Number(_)
                | CellValue::Boolean(_)
                | CellValue::String(_)
                | CellValue::Error(_)
                | CellValue::RichText(_)
        ) {
            let scalar_input = cell_value_to_scalar_json_input(&input);
            self.set_cell_internal(sheet, address, scalar_input)?;

            // Preserve the typed representation for `getCellRich.input`.
            //
            // Note: For rich text values, the engine currently only stores the plain string value.
            // Persisting the input here allows callers to round-trip rich text styling even though
            // `getCellRich.value` will still reflect the scalar engine value.
            if !input.is_empty() {
                let sheet = self.ensure_sheet(sheet);
                let address = Self::parse_address(address)?.to_a1();
                self.sheets_rich
                    .entry(sheet)
                    .or_default()
                    .insert(address, input);
            }

            return Ok(());
        }

        let sheet = self.ensure_sheet(sheet);
        let cell_ref = Self::parse_address(address)?;
        let address = cell_ref.to_a1();

        if let Some((origin, end)) = self.engine.spill_range(&sheet, &address) {
            let edited_row = cell_ref.row;
            let edited_col = cell_ref.col;
            for row in origin.row..=end.row {
                for col in origin.col..=end.col {
                    // Skip the origin cell (top-left); we only need to clear spill outputs.
                    if row == origin.row && col == origin.col {
                        continue;
                    }
                    // If the user overwrote a spill output cell with a literal value, don't emit a
                    // spill-clear change for that cell; the caller already knows its new input.
                    if row == edited_row && col == edited_col {
                        continue;
                    }
                    self.pending_spill_clears
                        .insert(FormulaCellKey::new(sheet.clone(), CellRef::new(row, col)));
                }
            }
        }

        let sheet_cells = self
            .sheets
            .get_mut(&sheet)
            .expect("sheet just ensured must exist");
        let sheet_cells_rich = self
            .sheets_rich
            .get_mut(&sheet)
            .expect("sheet just ensured must exist");

        // Convert model cell value into the engine's runtime value.
        //
        // NOTE: Today we do not support directly setting dynamic arrays/spill markers via the WASM
        // worker API. If callers send `array`/`spill` values, feed a `#SPILL!` error into the engine
        // but still store the rich input for round-tripping through `getCellRich`.
        let engine_value = match &input {
            CellValue::Array(_) | CellValue::Spill(_) => EngineValue::Error(ErrorKind::Spill),
            CellValue::Image(image) => EngineValue::Text(
                image
                    .alt_text
                    .clone()
                    .filter(|s| !s.is_empty())
                    .unwrap_or_else(|| "[Image]".to_string()),
            ),
            _ => cell_value_to_engine_rich(&input)?,
        };
        self.engine
            .set_cell_value(&sheet, &address, engine_value)
            .map_err(|err| js_err(err.to_string()))?;

        // Rich values are not representable in the scalar workbook input schema; preserve scalar
        // compatibility by removing any stored scalar input for this cell.
        sheet_cells.remove(&address);

        // Store the full rich input for `getCellRich.input`.
        sheet_cells_rich.insert(address.clone(), input);

        self.pending_spill_clears
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
        self.pending_formula_baselines
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
        Ok(())
    }
    fn get_cell_data(&self, sheet: &str, address: &str) -> Result<CellData, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let address = Self::parse_address(address)?.to_a1();

        let input = self
            .sheets
            .get(&sheet)
            .and_then(|cells| cells.get(&address))
            .cloned()
            .unwrap_or(JsonValue::Null);

        let value = engine_value_to_json(self.engine.get_cell_value(&sheet, &address));

        Ok(CellData {
            sheet,
            address,
            input,
            value,
        })
    }

    fn get_cell_rich_data(&self, sheet: &str, address: &str) -> Result<CellDataRich, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let address = Self::parse_address(address)?.to_a1();

        let input = self
            .sheets_rich
            .get(&sheet)
            .and_then(|cells| cells.get(&address))
            .cloned()
            .unwrap_or_else(|| {
                let scalar = self
                    .sheets
                    .get(&sheet)
                    .and_then(|cells| cells.get(&address))
                    .cloned()
                    .unwrap_or(JsonValue::Null);
                scalar_json_to_cell_value_input(&scalar)
            });

        let value = engine_value_to_cell_value_rich(self.engine.get_cell_value(&sheet, &address));

        Ok(CellDataRich {
            sheet,
            address,
            input,
            value,
        })
    }

    fn recalculate_internal(&mut self, sheet: Option<&str>) -> Result<Vec<CellChange>, JsValue> {
        // The JS worker protocol historically accepted a `sheet` argument for API symmetry, but
        // callers rely on `recalculate()` returning *all* value changes across the workbook so
        // client-side caches stay coherent across sheet switches.
        //
        // Therefore we intentionally ignore `sheet` here (and do not validate it).
        let _ = sheet;

        let recalc_changes = self.engine.recalculate_with_value_changes_single_threaded();
        let mut by_cell: BTreeMap<FormulaCellKey, JsonValue> = BTreeMap::new();

        for change in recalc_changes {
            by_cell.insert(
                FormulaCellKey {
                    sheet: change.sheet,
                    row: change.addr.row,
                    col: change.addr.col,
                },
                engine_value_to_json(change.value),
            );
        }

        let pending_spills = std::mem::take(&mut self.pending_spill_clears);
        for key in pending_spills {
            if by_cell.contains_key(&key) {
                continue;
            }
            let address = key.address();
            let value = engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
            by_cell.insert(key, value);
        }

        let pending_formulas = std::mem::take(&mut self.pending_formula_baselines);
        for (key, before) in pending_formulas {
            if by_cell.contains_key(&key) {
                continue;
            }
            let address = key.address();
            let after = engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
            if after != before {
                by_cell.insert(key, after);
            }
        }

        let changes: Vec<CellChange> = by_cell
            .into_iter()
            .map(|(key, value)| {
                let address = key.address();
                CellChange {
                    sheet: key.sheet,
                    address,
                    value,
                }
            })
            .collect();

        Ok(changes)
    }

    fn collect_spill_output_cells(&self) -> BTreeSet<FormulaCellKey> {
        let mut out = BTreeSet::new();
        for (sheet_name, cells) in &self.sheets {
            for (address, input) in cells {
                if !is_formula_input(input) {
                    continue;
                }
                let Some((origin, end)) = self.engine.spill_range(sheet_name, address) else {
                    continue;
                };
                for row in origin.row..=end.row {
                    for col in origin.col..=end.col {
                        if row == origin.row && col == origin.col {
                            continue;
                        }
                        out.insert(FormulaCellKey::new(
                            sheet_name.clone(),
                            CellRef::new(row, col),
                        ));
                    }
                }
            }
        }
        out
    }

    fn edit_op_from_dto(&mut self, dto: EditOpDto) -> Result<EngineEditOp, JsValue> {
        match dto {
            EditOpDto::InsertRows { sheet, row, count } => {
                let sheet = self.ensure_sheet(&sheet);
                Ok(EngineEditOp::InsertRows { sheet, row, count })
            }
            EditOpDto::DeleteRows { sheet, row, count } => {
                let sheet = self.ensure_sheet(&sheet);
                Ok(EngineEditOp::DeleteRows { sheet, row, count })
            }
            EditOpDto::InsertCols { sheet, col, count } => {
                let sheet = self.ensure_sheet(&sheet);
                Ok(EngineEditOp::InsertCols { sheet, col, count })
            }
            EditOpDto::DeleteCols { sheet, col, count } => {
                let sheet = self.ensure_sheet(&sheet);
                Ok(EngineEditOp::DeleteCols { sheet, col, count })
            }
            EditOpDto::InsertCellsShiftRight { sheet, range } => {
                let sheet = self.ensure_sheet(&sheet);
                let range = Self::parse_range(&range)?;
                Ok(EngineEditOp::InsertCellsShiftRight { sheet, range })
            }
            EditOpDto::InsertCellsShiftDown { sheet, range } => {
                let sheet = self.ensure_sheet(&sheet);
                let range = Self::parse_range(&range)?;
                Ok(EngineEditOp::InsertCellsShiftDown { sheet, range })
            }
            EditOpDto::DeleteCellsShiftLeft { sheet, range } => {
                let sheet = self.ensure_sheet(&sheet);
                let range = Self::parse_range(&range)?;
                Ok(EngineEditOp::DeleteCellsShiftLeft { sheet, range })
            }
            EditOpDto::DeleteCellsShiftUp { sheet, range } => {
                let sheet = self.ensure_sheet(&sheet);
                let range = Self::parse_range(&range)?;
                Ok(EngineEditOp::DeleteCellsShiftUp { sheet, range })
            }
            EditOpDto::MoveRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet = self.ensure_sheet(&sheet);
                let src = Self::parse_range(&src)?;
                let dst_top_left = Self::parse_address(&dst_top_left)?;
                Ok(EngineEditOp::MoveRange {
                    sheet,
                    src,
                    dst_top_left,
                })
            }
            EditOpDto::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet = self.ensure_sheet(&sheet);
                let src = Self::parse_range(&src)?;
                let dst_top_left = Self::parse_address(&dst_top_left)?;
                Ok(EngineEditOp::CopyRange {
                    sheet,
                    src,
                    dst_top_left,
                })
            }
            EditOpDto::Fill { sheet, src, dst } => {
                let sheet = self.ensure_sheet(&sheet);
                let src = Self::parse_range(&src)?;
                let dst = Self::parse_range(&dst)?;
                Ok(EngineEditOp::Fill { sheet, src, dst })
            }
        }
    }

    fn remap_pending_keys_for_edit(&mut self, op: &EngineEditOp) {
        fn remap_key(key: &FormulaCellKey, op: &EngineEditOp) -> Option<FormulaCellKey> {
            match op {
                EngineEditOp::InsertRows { sheet, row, count } if &key.sheet == sheet => {
                    if key.row >= *row {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row + *count,
                            col: key.col,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::DeleteRows { sheet, row, count } if &key.sheet == sheet => {
                    let start = *row;
                    let end_exclusive = row.saturating_add(*count);
                    if key.row >= start && key.row < end_exclusive {
                        None
                    } else if key.row >= end_exclusive {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row - *count,
                            col: key.col,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::InsertCols { sheet, col, count } if &key.sheet == sheet => {
                    if key.col >= *col {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row,
                            col: key.col + *count,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::DeleteCols { sheet, col, count } if &key.sheet == sheet => {
                    let start = *col;
                    let end_exclusive = col.saturating_add(*count);
                    if key.col >= start && key.col < end_exclusive {
                        None
                    } else if key.col >= end_exclusive {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row,
                            col: key.col - *count,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::InsertCellsShiftRight { sheet, range } if &key.sheet == sheet => {
                    let width = range.width();
                    if key.row >= range.start.row
                        && key.row <= range.end.row
                        && key.col >= range.start.col
                    {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row,
                            col: key.col + width,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::InsertCellsShiftDown { sheet, range } if &key.sheet == sheet => {
                    let height = range.height();
                    if key.col >= range.start.col
                        && key.col <= range.end.col
                        && key.row >= range.start.row
                    {
                        Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row: key.row + height,
                            col: key.col,
                        })
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::DeleteCellsShiftLeft { sheet, range } if &key.sheet == sheet => {
                    let width = range.width();
                    if key.row >= range.start.row && key.row <= range.end.row {
                        if key.col >= range.start.col && key.col <= range.end.col {
                            None
                        } else if key.col > range.end.col {
                            Some(FormulaCellKey {
                                sheet: key.sheet.clone(),
                                row: key.row,
                                col: key.col - width,
                            })
                        } else {
                            Some(key.clone())
                        }
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::DeleteCellsShiftUp { sheet, range } if &key.sheet == sheet => {
                    let height = range.height();
                    if key.col >= range.start.col && key.col <= range.end.col {
                        if key.row >= range.start.row && key.row <= range.end.row {
                            None
                        } else if key.row > range.end.row {
                            Some(FormulaCellKey {
                                sheet: key.sheet.clone(),
                                row: key.row - height,
                                col: key.col,
                            })
                        } else {
                            Some(key.clone())
                        }
                    } else {
                        Some(key.clone())
                    }
                }
                EngineEditOp::MoveRange {
                    sheet,
                    src,
                    dst_top_left,
                } if &key.sheet == sheet => {
                    let delta_row = dst_top_left.row as i64 - src.start.row as i64;
                    let delta_col = dst_top_left.col as i64 - src.start.col as i64;

                    let dst_end = CellRef::new(
                        dst_top_left.row + src.height().saturating_sub(1),
                        dst_top_left.col + src.width().saturating_sub(1),
                    );
                    let dst = Range::new(*dst_top_left, dst_end);

                    let in_src = key.row >= src.start.row
                        && key.row <= src.end.row
                        && key.col >= src.start.col
                        && key.col <= src.end.col;
                    if in_src {
                        let row = (key.row as i64 + delta_row).max(0) as u32;
                        let col = (key.col as i64 + delta_col).max(0) as u32;
                        return Some(FormulaCellKey {
                            sheet: key.sheet.clone(),
                            row,
                            col,
                        });
                    }

                    // Destination range contents are overwritten by the move.
                    let in_dst = key.row >= dst.start.row
                        && key.row <= dst.end.row
                        && key.col >= dst.start.col
                        && key.col <= dst.end.col;
                    if in_dst {
                        return None;
                    }

                    Some(key.clone())
                }
                EngineEditOp::CopyRange {
                    sheet,
                    src,
                    dst_top_left,
                } if &key.sheet == sheet => {
                    let dst_end = CellRef::new(
                        dst_top_left.row + src.height().saturating_sub(1),
                        dst_top_left.col + src.width().saturating_sub(1),
                    );
                    let dst = Range::new(*dst_top_left, dst_end);
                    let in_dst = key.row >= dst.start.row
                        && key.row <= dst.end.row
                        && key.col >= dst.start.col
                        && key.col <= dst.end.col;
                    if in_dst {
                        // Destination range contents are overwritten by the copy.
                        return None;
                    }
                    Some(key.clone())
                }
                EngineEditOp::Fill { sheet, src, dst } if &key.sheet == sheet => {
                    let in_dst = key.row >= dst.start.row
                        && key.row <= dst.end.row
                        && key.col >= dst.start.col
                        && key.col <= dst.end.col;
                    if !in_dst {
                        return Some(key.clone());
                    }

                    let in_src = key.row >= src.start.row
                        && key.row <= src.end.row
                        && key.col >= src.start.col
                        && key.col <= src.end.col;
                    if in_src {
                        // Preserve the source range; only the surrounding destination range is overwritten.
                        return Some(key.clone());
                    }

                    None
                }
                _ => Some(key.clone()),
            }
        }

        let mut next_spills = BTreeSet::new();
        for key in std::mem::take(&mut self.pending_spill_clears) {
            if let Some(remapped) = remap_key(&key, op) {
                next_spills.insert(remapped);
            }
        }
        self.pending_spill_clears = next_spills;

        let mut next_formulas = BTreeMap::new();
        for (key, baseline) in std::mem::take(&mut self.pending_formula_baselines) {
            if let Some(remapped) = remap_key(&key, op) {
                // If multiple keys map to the same cell, keep the earliest baseline.
                next_formulas.entry(remapped).or_insert(baseline);
            }
        }
        self.pending_formula_baselines = next_formulas;

        // Remap rich inputs to follow the same structural edit semantics as the engine.
        match op {
            EngineEditOp::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let dst_end = CellRef::new(
                    dst_top_left.row + src.height().saturating_sub(1),
                    dst_top_left.col + src.width().saturating_sub(1),
                );
                let dst = Range::new(*dst_top_left, dst_end);

                let mut next_rich: BTreeMap<String, BTreeMap<String, CellValue>> = BTreeMap::new();
                for (sheet_name, cells) in std::mem::take(&mut self.sheets_rich) {
                    if &sheet_name != sheet {
                        next_rich.insert(sheet_name, cells);
                        continue;
                    }

                    let mut copied: Vec<(u32, u32, CellValue)> = Vec::new();
                    for (address, value) in &cells {
                        let Ok(cell_ref) = CellRef::from_a1(address) else {
                            continue;
                        };
                        let in_src = cell_ref.row >= src.start.row
                            && cell_ref.row <= src.end.row
                            && cell_ref.col >= src.start.col
                            && cell_ref.col <= src.end.col;
                        if in_src {
                            copied.push((cell_ref.row, cell_ref.col, value.clone()));
                        }
                    }

                    let mut new_cells = BTreeMap::new();
                    for (address, value) in cells {
                        let Ok(cell_ref) = CellRef::from_a1(&address) else {
                            new_cells.insert(address, value);
                            continue;
                        };
                        let in_dst = cell_ref.row >= dst.start.row
                            && cell_ref.row <= dst.end.row
                            && cell_ref.col >= dst.start.col
                            && cell_ref.col <= dst.end.col;
                        if !in_dst {
                            new_cells.insert(address, value);
                        }
                    }

                    for (row, col, value) in copied {
                        let dest_row = dst_top_left.row + (row - src.start.row);
                        let dest_col = dst_top_left.col + (col - src.start.col);
                        let address = CellRef::new(dest_row, dest_col).to_a1();
                        new_cells.insert(address, value);
                    }

                    next_rich.insert(sheet_name, new_cells);
                }
                self.sheets_rich = next_rich;
            }
            _ => {
                let mut next_rich: BTreeMap<String, BTreeMap<String, CellValue>> = BTreeMap::new();
                for (sheet_name, cells) in std::mem::take(&mut self.sheets_rich) {
                    for (address, input) in cells {
                        let Ok(cell_ref) = CellRef::from_a1(&address) else {
                            continue;
                        };
                        let key = FormulaCellKey {
                            sheet: sheet_name.clone(),
                            row: cell_ref.row,
                            col: cell_ref.col,
                        };
                        let Some(remapped) = remap_key(&key, op) else {
                            continue;
                        };
                        let remapped_address = CellRef::new(remapped.row, remapped.col).to_a1();
                        next_rich
                            .entry(remapped.sheet)
                            .or_default()
                            .insert(remapped_address, input);
                    }
                }
                self.sheets_rich = next_rich;
            }
        }
    }

    fn apply_operation_internal(&mut self, dto: EditOpDto) -> Result<EditResultDto, JsValue> {
        let spill_outputs_before = self.collect_spill_output_cells();
        let op = self.edit_op_from_dto(dto)?;
        self.remap_pending_keys_for_edit(&op);

        let result: EngineEditResult = self
            .engine
            .apply_operation(op)
            .map_err(|err| js_err(edit_error_to_string(err)))?;

        // Update the persisted input map used by `toJson` and `getCell.input`.
        for change in &result.changed_cells {
            let sheet = self.ensure_sheet(&change.sheet);
            let address = change.cell.to_a1();
            let sheet_cells = self
                .sheets
                .get_mut(&sheet)
                .expect("sheet just ensured must exist");

            match &change.after {
                None => {
                    sheet_cells.remove(&address);
                }
                Some(after) => {
                    if let Some(formula) = after.formula.as_deref() {
                        sheet_cells.insert(address.clone(), JsonValue::String(formula.to_string()));
                    } else {
                        let Some(value) = engine_value_to_scalar_json_input(after.value.clone())
                        else {
                            sheet_cells.remove(&address);
                            continue;
                        };
                        sheet_cells.insert(address.clone(), value);
                    }
                }
            }
        }

        // Preserve the WASM worker semantics where formula cells return blank values until the next
        // explicit `recalculate()` call.
        for change in &result.changed_cells {
            let Some(after) = &change.after else {
                let sheet = self.ensure_sheet(&change.sheet);
                self.pending_spill_clears
                    .remove(&FormulaCellKey::new(sheet.clone(), change.cell));
                self.pending_formula_baselines
                    .remove(&FormulaCellKey::new(sheet.clone(), change.cell));
                continue;
            };

            let sheet = self.ensure_sheet(&change.sheet);
            let address = change.cell.to_a1();
            let key = FormulaCellKey::new(sheet.clone(), change.cell);

            if let Some(formula) = after.formula.as_deref() {
                self.pending_formula_baselines
                    .entry(key)
                    .or_insert_with(|| {
                        engine_value_to_json(self.engine.get_cell_value(&sheet, &address))
                    });

                // Reset stored value to blank while preserving the formula. This matches the
                // `setCell` behavior where formula results are treated as unknown until recalc.
                self.engine
                    .set_cell_value(&sheet, &address, EngineValue::Blank)
                    .map_err(|err| js_err(err.to_string()))?;
                self.engine
                    .set_cell_formula(&sheet, &address, formula)
                    .map_err(|err| js_err(err.to_string()))?;
            } else {
                // This cell is now a literal (or empty) value; remove any stale baseline.
                self.pending_formula_baselines.remove(&key);
            }
        }

        // Spill ranges are maintained by the engine across recalc ticks, but `apply_operation`
        // rebuilds the dependency graph (and clears spill metadata). Capture spill output cells from
        // before the edit so the next `recalculate()` call can emit `null` deltas for any now-stale
        // spill values that would otherwise be lost.
        for key in spill_outputs_before {
            let address = key.address();
            let has_input = self
                .sheets
                .get(&key.sheet)
                .and_then(|cells| cells.get(&address))
                .is_some();
            if has_input {
                continue;
            }
            self.pending_spill_clears.insert(key);
        }

        // Convert to JS-friendly DTO.
        let mut changed_cells = Vec::with_capacity(result.changed_cells.len());
        for change in &result.changed_cells {
            let address = change.cell.to_a1();
            let before = change.before.as_ref().map(|snap| EditCellSnapshotDto {
                value: engine_value_to_json(snap.value.clone()),
                formula: snap.formula.clone(),
            });
            let after = change.after.as_ref().map(|snap| {
                let is_formula = snap.formula.is_some();
                EditCellSnapshotDto {
                    value: if is_formula {
                        JsonValue::Null
                    } else {
                        engine_value_to_json(snap.value.clone())
                    },
                    formula: snap.formula.clone(),
                }
            });

            changed_cells.push(EditCellChangeDto {
                sheet: change.sheet.clone(),
                address,
                before,
                after,
            });
        }

        let moved_ranges = result
            .moved_ranges
            .iter()
            .map(|m| EditMovedRangeDto {
                sheet: m.sheet.clone(),
                from: m.from.to_string(),
                to: m.to.to_string(),
            })
            .collect();

        let formula_rewrites = result
            .formula_rewrites
            .iter()
            .map(|r| EditFormulaRewriteDto {
                sheet: r.sheet.clone(),
                address: r.cell.to_a1(),
                before: r.before.clone(),
                after: r.after.clone(),
            })
            .collect();

        Ok(EditResultDto {
            changed_cells,
            moved_ranges,
            formula_rewrites,
        })
    }

    fn set_locale_id(&mut self, locale_id: &str) -> bool {
        let Some(formula_locale) = get_locale(locale_id) else {
            return false;
        };
        let Some(value_locale) = ValueLocaleConfig::for_locale_id(locale_id) else {
            return false;
        };

        self.formula_locale = formula_locale;
        self.engine.set_locale_config(formula_locale.config.clone());
        self.engine.set_value_locale(value_locale);
        true
    }
}

fn json_scalar_to_js(value: &JsonValue) -> JsValue {
    match value {
        JsonValue::Null => JsValue::NULL,
        JsonValue::Bool(b) => JsValue::from_bool(*b),
        JsonValue::Number(n) => n.as_f64().map(JsValue::from_f64).unwrap_or(JsValue::NULL),
        JsonValue::String(s) => JsValue::from_str(s),
        // The engine protocol only supports scalars; fall back to `null` for any
        // unexpected values to avoid surfacing `undefined`.
        _ => JsValue::NULL,
    }
}

fn object_set(obj: &Object, key: &str, value: &JsValue) -> Result<(), JsValue> {
    Reflect::set(obj, &JsValue::from_str(key), value).map(|_| ())
}

fn cell_data_to_js(cell: &CellData) -> Result<JsValue, JsValue> {
    let obj = Object::new();
    object_set(&obj, "sheet", &JsValue::from_str(&cell.sheet))?;
    object_set(&obj, "address", &JsValue::from_str(&cell.address))?;
    object_set(&obj, "input", &json_scalar_to_js(&cell.input))?;
    object_set(&obj, "value", &json_scalar_to_js(&cell.value))?;
    Ok(obj.into())
}

fn cell_change_to_js(change: &CellChange) -> Result<JsValue, JsValue> {
    let obj = Object::new();
    object_set(&obj, "sheet", &JsValue::from_str(&change.sheet))?;
    object_set(&obj, "address", &JsValue::from_str(&change.address))?;
    object_set(&obj, "value", &json_scalar_to_js(&change.value))?;
    Ok(obj.into())
}

fn utf16_cursor_to_byte_index(s: &str, cursor_utf16: u32) -> usize {
    let cursor_utf16 = cursor_utf16 as usize;
    if cursor_utf16 == 0 {
        return 0;
    }

    let mut seen_utf16: usize = 0;
    for (byte_idx, ch) in s.char_indices() {
        let ch_utf16 = ch.len_utf16();
        if seen_utf16 + ch_utf16 > cursor_utf16 {
            // Cursor points into the middle of this char (possible for surrogate pairs).
            // Clamp to the previous valid UTF-8 boundary.
            return byte_idx;
        }
        seen_utf16 += ch_utf16;
        if seen_utf16 == cursor_utf16 {
            return byte_idx + ch.len_utf8();
        }
    }
    s.len()
}

fn byte_index_to_utf16_cursor(s: &str, byte_idx: usize) -> usize {
    let mut byte_idx = byte_idx.min(s.len());
    while byte_idx > 0 && !s.is_char_boundary(byte_idx) {
        byte_idx -= 1;
    }
    s[..byte_idx].encode_utf16().count()
}

fn is_ident_start_char(c: char) -> bool {
    matches!(c, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z') || (!c.is_ascii() && c.is_alphabetic())
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(
        c,
        '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'
    ) || (!c.is_ascii() && c.is_alphanumeric())
}

#[derive(Debug)]
struct FallbackFunctionFrame {
    name: String,
    paren_depth: usize,
    arg_index: usize,
    brace_depth: usize,
    bracket_depth: usize,
}

fn scan_fallback_function_context(
    formula_prefix: &str,
    arg_separator: char,
) -> Option<formula_engine::FunctionContext> {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum Mode {
        Normal,
        String,
        QuotedIdent,
    }

    let mut mode = Mode::Normal;
    let mut paren_depth: usize = 0;
    let mut brace_depth: usize = 0;
    let mut bracket_depth: usize = 0;
    let mut stack: Vec<FallbackFunctionFrame> = Vec::new();

    let mut i: usize = 0;
    while i < formula_prefix.len() {
        let ch = formula_prefix[i..]
            .chars()
            .next()
            .expect("char_indices iteration should always yield a char");
        let ch_len = ch.len_utf8();

        match mode {
            Mode::String => {
                if ch == '"' {
                    let next_i = i + ch_len;
                    if next_i < formula_prefix.len()
                        && formula_prefix[next_i..].chars().next() == Some('"')
                    {
                        // Escaped quote within a string literal: `""`.
                        i = next_i + 1;
                    } else {
                        // Closing quote.
                        mode = Mode::Normal;
                        i = next_i;
                    }
                    continue;
                }

                i += ch_len;
                continue;
            }
            Mode::QuotedIdent => {
                if ch == '\'' {
                    let next_i = i + ch_len;
                    if next_i < formula_prefix.len()
                        && formula_prefix[next_i..].chars().next() == Some('\'')
                    {
                        // Escaped quote within a quoted identifier: `''`.
                        i = next_i + 1;
                    } else {
                        mode = Mode::Normal;
                        i = next_i;
                    }
                    continue;
                }

                i += ch_len;
                continue;
            }
            Mode::Normal => {
                // In the engine lexer, quotes are treated as literal characters inside
                // structured reference brackets, so only treat them as string/quoted-ident
                // openers when we're not in a bracket segment.
                if bracket_depth == 0 {
                    if ch == '"' {
                        mode = Mode::String;
                        i += ch_len;
                        continue;
                    }
                    if ch == '\'' {
                        mode = Mode::QuotedIdent;
                        i += ch_len;
                        continue;
                    }
                }

                if bracket_depth > 0 {
                    // Mirror `formula-engine`'s lexer behavior: inside structured-ref/workbook
                    // brackets, treat everything as raw text except nested bracket open/close.
                    match ch {
                        '[' => bracket_depth += 1,
                        ']' => {
                            // Excel escapes `]` inside structured references as `]]`. At the
                            // outermost bracket depth, treat a double `]]` as a literal `]` rather
                            // than the end of the bracket segment.
                            if bracket_depth == 1 && formula_prefix[i..].starts_with("]]") {
                                i += 2;
                                continue;
                            }
                            bracket_depth = bracket_depth.saturating_sub(1);
                        }
                        _ => {}
                    }
                    i += ch_len;
                    continue;
                }

                match ch {
                    '[' => {
                        bracket_depth += 1;
                        i += ch_len;
                    }
                    ']' => {
                        if bracket_depth > 0 {
                            bracket_depth -= 1;
                        }
                        i += ch_len;
                    }
                    '{' => {
                        brace_depth += 1;
                        i += ch_len;
                    }
                    '}' => {
                        if brace_depth > 0 {
                            brace_depth -= 1;
                        }
                        i += ch_len;
                    }
                    '(' => {
                        paren_depth += 1;
                        i += ch_len;
                    }
                    ')' => {
                        if paren_depth > 0 {
                            if stack
                                .last()
                                .is_some_and(|frame| frame.paren_depth == paren_depth)
                            {
                                stack.pop();
                            }
                            paren_depth -= 1;
                        }
                        i += ch_len;
                    }
                    c if c == arg_separator => {
                        if let Some(frame) = stack.last_mut() {
                            // Count only separators that are at the "top level" within the call.
                            if paren_depth == frame.paren_depth
                                && brace_depth == frame.brace_depth
                                && bracket_depth == frame.bracket_depth
                            {
                                frame.arg_index += 1;
                            }
                        }
                        i += ch_len;
                    }
                    c if is_ident_start_char(c) => {
                        let start = i;
                        let mut end = i + ch_len;
                        while end < formula_prefix.len() {
                            let next = formula_prefix[end..]
                                .chars()
                                .next()
                                .expect("slice must start at char boundary");
                            if is_ident_cont_char(next) {
                                end += next.len_utf8();
                            } else {
                                break;
                            }
                        }

                        let ident = &formula_prefix[start..end];

                        // Look ahead for `(`, allowing whitespace between.
                        let mut j = end;
                        while j < formula_prefix.len() {
                            let next = formula_prefix[j..]
                                .chars()
                                .next()
                                .expect("slice must start at char boundary");
                            if next.is_whitespace() {
                                j += next.len_utf8();
                            } else {
                                break;
                            }
                        }

                        if j < formula_prefix.len()
                            && formula_prefix[j..].chars().next() == Some('(')
                        {
                            paren_depth += 1;
                            stack.push(FallbackFunctionFrame {
                                name: ident.to_ascii_uppercase(),
                                paren_depth,
                                arg_index: 0,
                                brace_depth,
                                bracket_depth,
                            });
                            // Skip whitespace + `(`.
                            i = j + 1;
                        } else {
                            i = end;
                        }
                    }
                    _ => {
                        i += ch_len;
                    }
                }
            }
        }
    }

    stack.last().map(|frame| formula_engine::FunctionContext {
        name: frame.name.clone(),
        arg_index: frame.arg_index,
    })
}

#[derive(Debug, Serialize)]
struct WasmSpan {
    start: usize,
    end: usize,
}

#[derive(Debug, Serialize)]
struct WasmParseError {
    message: String,
    span: WasmSpan,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct WasmFunctionContext {
    name: String,
    arg_index: usize,
}

#[derive(Debug, Serialize)]
struct WasmParseContext {
    function: Option<WasmFunctionContext>,
}

#[derive(Debug, Serialize)]
struct WasmPartialParse {
    ast: formula_engine::Ast,
    error: Option<WasmParseError>,
    context: WasmParseContext,
}
#[wasm_bindgen(js_name = "parseFormulaPartial")]
pub fn parse_formula_partial(
    formula: String,
    cursor: Option<u32>,
    opts: Option<JsValue>,
) -> Result<JsValue, JsValue> {
    ensure_rust_constructors_run();

    let opts = parse_options_from_js(opts)?;

    // Cursor is expressed in UTF-16 code units by JS callers.
    let formula_utf16_len = formula.encode_utf16().count() as u32;
    let cursor_utf16 = cursor.unwrap_or(formula_utf16_len).min(formula_utf16_len);
    let byte_cursor = utf16_cursor_to_byte_index(&formula, cursor_utf16);
    let prefix = &formula[..byte_cursor];

    let mut parsed = formula_engine::parse_formula_partial(prefix, opts.clone());
    if parsed.context.function.is_none() {
        let lex_error = parsed.error.as_ref().is_some_and(|err| {
            matches!(
                err.message.as_str(),
                "Unterminated string literal" | "Unterminated quoted identifier"
            )
        });
        if lex_error {
            parsed.context.function =
                scan_fallback_function_context(prefix, opts.locale.arg_separator);
        }
    }

    let error = parsed.error.map(|err| WasmParseError {
        message: err.message,
        span: WasmSpan {
            start: byte_index_to_utf16_cursor(prefix, err.span.start),
            end: byte_index_to_utf16_cursor(prefix, err.span.end),
        },
    });

    let context = WasmParseContext {
        function: parsed.context.function.map(|ctx| WasmFunctionContext {
            name: ctx.name.to_ascii_uppercase(),
            arg_index: ctx.arg_index,
        }),
    };

    let out = WasmPartialParse {
        ast: parsed.ast,
        error,
        context,
    };

    use serde::ser::Serialize as _;
    out.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
        .map_err(|err| js_err(err.to_string()))
}
#[wasm_bindgen]
pub struct WasmWorkbook {
    inner: WorkbookState,
}

#[wasm_bindgen]
impl WasmWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> WasmWorkbook {
        WasmWorkbook {
            inner: WorkbookState::new_with_default_sheet(),
        }
    }

    #[wasm_bindgen(js_name = "setLocale")]
    pub fn set_locale(&mut self, locale_id: String) -> bool {
        self.inner.set_locale_id(&locale_id)
    }

    #[wasm_bindgen(js_name = "fromJson")]
    pub fn from_json(json: &str) -> Result<WasmWorkbook, JsValue> {
        #[derive(Debug, Deserialize)]
        struct WorkbookJson {
            sheets: BTreeMap<String, SheetJson>,
        }

        #[derive(Debug, Deserialize)]
        struct SheetJson {
            #[serde(default, rename = "rowCount")]
            row_count: Option<u32>,
            #[serde(default, rename = "colCount")]
            col_count: Option<u32>,
            cells: BTreeMap<String, JsonValue>,
        }

        let parsed: WorkbookJson = serde_json::from_str(json)
            .map_err(|err| js_err(format!("invalid workbook json: {err}")))?;

        let mut wb = WorkbookState::new_empty();

        // Create all sheets up-front so cross-sheet formula references resolve correctly.
        for sheet_name in parsed.sheets.keys() {
            wb.ensure_sheet(sheet_name);
        }

        for (sheet_name, sheet) in parsed.sheets {
            // Apply sheet dimensions (when provided) before importing cells so large addresses
            // can be set without pre-populating the full grid.
            if sheet.row_count.is_some() || sheet.col_count.is_some() {
                let rows = sheet.row_count.unwrap_or(EXCEL_MAX_ROWS);
                let cols = sheet.col_count.unwrap_or(EXCEL_MAX_COLS);
                if rows != EXCEL_MAX_ROWS || cols != EXCEL_MAX_COLS {
                    wb.set_sheet_dimensions_internal(&sheet_name, rows, cols)?;
                }
            }

            for (address, input) in sheet.cells {
                if !is_scalar_json(&input) {
                    return Err(js_err(format!("invalid cell value: {address}")));
                }
                if input.is_null() {
                    // `null` cells are treated as absent (sparse semantics).
                    continue;
                }
                wb.set_cell_internal(&sheet_name, &address, input)?;
            }
        }

        if wb.sheets.is_empty() {
            wb.ensure_sheet(DEFAULT_SHEET);
        }

        Ok(WasmWorkbook { inner: wb })
    }

    #[wasm_bindgen(js_name = "fromXlsxBytes")]
    pub fn from_xlsx_bytes(bytes: &[u8]) -> Result<WasmWorkbook, JsValue> {
        let model = formula_xlsx::read_workbook_model_from_bytes(bytes)
            .map_err(|err| js_err(err.to_string()))?;

        let mut wb = WorkbookState::new_empty();

        // Date system influences date serials for NOW/TODAY/DATE, etc.
        wb.engine.set_date_system(match model.date_system {
            DateSystem::Excel1900 => formula_engine::date::ExcelDateSystem::EXCEL_1900,
            DateSystem::Excel1904 => formula_engine::date::ExcelDateSystem::Excel1904,
        });

        // Create all sheets up-front so formulas can resolve cross-sheet references.
        for sheet in &model.sheets {
            wb.ensure_sheet(&sheet.name);
        }

        // Apply per-sheet dimensions (logical grid size) before importing cells/formulas so
        // whole-column/row semantics (`A:A`, `1:1`) resolve correctly for large sheets.
        for sheet in &model.sheets {
            if sheet.row_count != EXCEL_MAX_ROWS || sheet.col_count != EXCEL_MAX_COLS {
                wb.set_sheet_dimensions_internal(&sheet.name, sheet.row_count, sheet.col_count)?;
            }
        }

        // Import Excel tables (structured reference metadata) before formulas are compiled so
        // expressions like `Table1[Col]` and `[@Col]` resolve correctly.
        for sheet in &model.sheets {
            let sheet_name = wb
                .resolve_sheet(&sheet.name)
                .expect("sheet just ensured must resolve")
                .to_string();
            wb.engine
                .set_sheet_tables(&sheet_name, sheet.tables.clone());
        }

        // Best-effort defined names.
        let mut sheet_names_by_id: HashMap<u32, String> = HashMap::new();
        for sheet in &model.sheets {
            sheet_names_by_id.insert(sheet.id, sheet.name.clone());
        }

        for name in &model.defined_names {
            let scope = match name.scope {
                DefinedNameScope::Workbook => NameScope::Workbook,
                DefinedNameScope::Sheet(sheet_id) => {
                    let Some(sheet_name) = sheet_names_by_id.get(&sheet_id) else {
                        continue;
                    };
                    NameScope::Sheet(sheet_name)
                }
            };

            let refers_to = name.refers_to.trim();
            if refers_to.is_empty() {
                continue;
            }

            // Best-effort heuristic:
            // - numeric/bool constants are imported as constants
            // - everything else is imported as a reference-like expression
            let definition = if refers_to.eq_ignore_ascii_case("TRUE") {
                NameDefinition::Constant(EngineValue::Bool(true))
            } else if refers_to.eq_ignore_ascii_case("FALSE") {
                NameDefinition::Constant(EngineValue::Bool(false))
            } else if let Ok(n) = refers_to.parse::<f64>() {
                NameDefinition::Constant(EngineValue::Number(n))
            } else if let Ok(err) = refers_to.parse::<formula_model::ErrorValue>() {
                NameDefinition::Constant(EngineValue::Error(err.into()))
            } else {
                NameDefinition::Reference(refers_to.to_string())
            };

            let _ = wb.engine.define_name(&name.name, scope, definition);
        }

        for sheet in &model.sheets {
            let sheet_name = wb
                .resolve_sheet(&sheet.name)
                .expect("sheet just ensured must resolve")
                .to_string();

            for (cell_ref, cell) in sheet.iter_cells() {
                let address = cell_ref.to_a1();

                // Skip style-only cells (not representable in this WASM DTO surface).
                let has_formula = cell.formula.is_some();
                let has_value = !cell.value.is_empty();
                if !has_formula && !has_value {
                    continue;
                }

                // Seed cached values first (including cached formula results).
                wb.engine
                    .set_cell_value(&sheet_name, &address, cell_value_to_engine(&cell.value))
                    .map_err(|err| js_err(err.to_string()))?;

                if let Some(formula) = cell.formula.as_deref() {
                    // `formula-model` stores formulas without a leading '='.
                    let display = display_formula_text(formula);
                    if !display.is_empty() {
                        // Best-effort: if the formula fails to parse (unsupported syntax), leave the
                        // cached value and still store the display formula in the input map.
                        let _ = wb.engine.set_cell_formula(&sheet_name, &address, &display);

                        let sheet_cells = wb
                            .sheets
                            .get_mut(&sheet_name)
                            .expect("sheet just ensured must exist");
                        sheet_cells.insert(address.clone(), JsonValue::String(display));
                        continue;
                    }
                }

                // Non-formula cell; store scalar value as input.
                let sheet_cells = wb
                    .sheets
                    .get_mut(&sheet_name)
                    .expect("sheet just ensured must exist");
                sheet_cells.insert(address, cell_value_to_scalar_json_input(&cell.value));
            }
        }

        if wb.sheets.is_empty() {
            wb.ensure_sheet(DEFAULT_SHEET);
        }

        Ok(WasmWorkbook { inner: wb })
    }

    #[wasm_bindgen(js_name = "setSheetDimensions")]
    pub fn set_sheet_dimensions(
        &mut self,
        sheet_name: String,
        rows: u32,
        cols: u32,
    ) -> Result<(), JsValue> {
        self.inner
            .set_sheet_dimensions_internal(&sheet_name, rows, cols)
    }

    #[wasm_bindgen(js_name = "getSheetDimensions")]
    pub fn get_sheet_dimensions(&self, sheet_name: String) -> Result<JsValue, JsValue> {
        let (rows, cols) = self.inner.get_sheet_dimensions_internal(&sheet_name)?;
        let obj = Object::new();
        object_set(&obj, "rows", &JsValue::from_f64(rows as f64))?;
        object_set(&obj, "cols", &JsValue::from_f64(cols as f64))?;
        Ok(obj.into())
    }

    #[wasm_bindgen(js_name = "toJson")]
    pub fn to_json(&self) -> Result<String, JsValue> {
        #[derive(Serialize)]
        struct WorkbookJson {
            sheets: BTreeMap<String, SheetJson>,
        }

        #[derive(Serialize)]
        struct SheetJson {
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "rowCount")]
            row_count: Option<u32>,
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "colCount")]
            col_count: Option<u32>,
            cells: BTreeMap<String, JsonValue>,
        }

        let mut sheets = BTreeMap::new();
        for (sheet_name, cells) in &self.inner.sheets {
            let mut out_cells = BTreeMap::new();
            for (address, input) in cells {
                // Ensure we never serialize explicit `null` cells; empty cells are
                // omitted from the sparse workbook representation.
                if input.is_null() {
                    continue;
                }
                out_cells.insert(address.clone(), input.clone());
            }
            let (rows, cols) = self
                .inner
                .engine
                .sheet_dimensions(sheet_name)
                .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS));
            let row_count = (rows != EXCEL_MAX_ROWS).then_some(rows);
            let col_count = (cols != EXCEL_MAX_COLS).then_some(cols);
            sheets.insert(
                sheet_name.clone(),
                SheetJson {
                    row_count,
                    col_count,
                    cells: out_cells,
                },
            );
        }

        serde_json::to_string(&WorkbookJson { sheets })
            .map_err(|err| js_err(format!("invalid workbook json: {err}")))
    }

    #[wasm_bindgen(js_name = "getCell")]
    pub fn get_cell(&self, address: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let cell = self.inner.get_cell_data(sheet, &address)?;
        cell_data_to_js(&cell)
    }
    #[wasm_bindgen(js_name = "setCell")]
    pub fn set_cell(
        &mut self,
        address: String,
        input: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        if input.is_null() {
            return self
                .inner
                .set_cell_internal(sheet, &address, JsonValue::Null);
        }
        let input: JsonValue =
            serde_wasm_bindgen::from_value(input).map_err(|err| js_err(err.to_string()))?;
        self.inner.set_cell_internal(sheet, &address, input)
    }

    #[wasm_bindgen(js_name = "setCellRich")]
    pub fn set_cell_rich(
        &mut self,
        address: String,
        value: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        if value.is_null() || value.is_undefined() {
            // Preserve sparse semantics: treat null/undefined as clearing the cell.
            return self
                .inner
                .set_cell_rich_internal(sheet, &address, CellValue::Empty);
        }

        let input: CellValue = serde_wasm_bindgen::from_value(value)
            .map_err(|err| js_err(format!("invalid rich value: {err}")))?;
        self.inner.set_cell_rich_internal(sheet, &address, input)
    }

    #[wasm_bindgen(js_name = "getCellRich")]
    pub fn get_cell_rich(
        &self,
        address: String,
        sheet: Option<String>,
    ) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let cell = self.inner.get_cell_rich_data(sheet, &address)?;
        use serde::ser::Serialize as _;
        cell.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "setCells")]
    pub fn set_cells(&mut self, updates: JsValue) -> Result<(), JsValue> {
        #[derive(Deserialize)]
        struct CellUpdate {
            address: String,
            value: JsonValue,
            sheet: Option<String>,
        }

        let updates: Vec<CellUpdate> =
            serde_wasm_bindgen::from_value(updates).map_err(|err| js_err(err.to_string()))?;

        for update in updates {
            let sheet = update.sheet.as_deref().unwrap_or(DEFAULT_SHEET);
            self.inner
                .set_cell_internal(sheet, &update.address, update.value)?;
        }

        Ok(())
    }

    #[wasm_bindgen(js_name = "getRange")]
    pub fn get_range(&self, range: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet)?.to_string();
        let range = WorkbookState::parse_range(&range)?;

        let outer = Array::new();
        for row in range.start.row..=range.end.row {
            let inner = Array::new();
            for col in range.start.col..=range.end.col {
                let addr = CellRef::new(row, col).to_a1();
                let cell = self.inner.get_cell_data(&sheet, &addr)?;
                inner.push(&cell_data_to_js(&cell)?);
            }
            outer.push(&inner);
        }

        Ok(outer.into())
    }

    #[wasm_bindgen(js_name = "setRange")]
    pub fn set_range(
        &mut self,
        range: String,
        values: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let range_parsed = WorkbookState::parse_range(&range)?;

        let values: Vec<Vec<JsonValue>> =
            serde_wasm_bindgen::from_value(values).map_err(|err| js_err(err.to_string()))?;

        let expected_rows = range_parsed.height() as usize;
        let expected_cols = range_parsed.width() as usize;
        if values.len() != expected_rows || values.iter().any(|row| row.len() != expected_cols) {
            return Err(js_err(format!(
                "invalid range: range {range} expects {expected_rows}x{expected_cols} values"
            )));
        }

        for (r_idx, row_values) in values.into_iter().enumerate() {
            for (c_idx, input) in row_values.into_iter().enumerate() {
                let row = range_parsed.start.row + r_idx as u32;
                let col = range_parsed.start.col + c_idx as u32;
                let addr = CellRef::new(row, col).to_a1();
                self.inner.set_cell_internal(sheet, &addr, input)?;
            }
        }

        Ok(())
    }

    #[wasm_bindgen(js_name = "goalSeek")]
    pub fn goal_seek(&mut self, request: JsValue) -> Result<JsValue, JsValue> {
        if request.is_null() || request.is_undefined() {
            return Err(js_err("goalSeek request must be an object"));
        }
        let request: GoalSeekRequestDto = serde_wasm_bindgen::from_value(request)
            .map_err(|err| js_err(format!("invalid goalSeek request: {err}")))?;

        let sheet_name = request.sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet_name)?.to_string();

        let target_cell_raw = request.target_cell.trim();
        if target_cell_raw.is_empty() {
            return Err(js_err("targetCell must be a non-empty string"));
        }
        if target_cell_raw.contains('!') {
            return Err(js_err("targetCell must be an A1 address without a sheet prefix"));
        }
        let target_cell = CellRef::from_a1(target_cell_raw)
            .map_err(|_| js_err(format!("invalid targetCell address: {target_cell_raw}")))?
            .to_a1();

        let changing_cell_raw = request.changing_cell.trim();
        if changing_cell_raw.is_empty() {
            return Err(js_err("changingCell must be a non-empty string"));
        }
        if changing_cell_raw.contains('!') {
            return Err(js_err("changingCell must be an A1 address without a sheet prefix"));
        }
        let changing_cell_ref = CellRef::from_a1(changing_cell_raw)
            .map_err(|_| js_err(format!("invalid changingCell address: {changing_cell_raw}")))?;
        let changing_cell = changing_cell_ref.to_a1();

        if !request.target_value.is_finite() {
            return Err(js_err("targetValue must be a finite number"));
        }

        let mut params =
            GoalSeekParams::new(target_cell.as_str(), request.target_value, changing_cell.as_str());
        if let Some(tolerance) = request.tolerance {
            if !tolerance.is_finite() {
                return Err(js_err("tolerance must be a finite number"));
            }
            if !(tolerance > 0.0) {
                return Err(js_err("tolerance must be > 0"));
            }
            params.tolerance = tolerance;
        }
        if let Some(max_iterations) = request.max_iterations {
            if max_iterations == 0 {
                return Err(js_err("maxIterations must be > 0"));
            }
            params.max_iterations = max_iterations;
        }

        let recalc_mode = request.recalc_mode.map(|mode| match mode {
            GoalSeekRecalcModeDto::SingleThreaded => RecalcMode::SingleThreaded,
            GoalSeekRecalcModeDto::MultiThreaded => RecalcMode::MultiThreaded,
        });

        let result = {
            let mut model = EngineWhatIfModel::new(&mut self.inner.engine, sheet.clone());
            if let Some(mode) = recalc_mode {
                model = model.with_recalc_mode(mode);
            }
            GoalSeek::solve(&mut model, params).map_err(|err| js_err(err.to_string()))?
        };

        let solution_json = serde_json::Number::from_f64(result.solution)
            .map(JsonValue::Number)
            .ok_or_else(|| js_err("goalSeek produced a non-finite solution"))?;

        // Ensure the JS-facing workbook input state matches the what-if engine updates.
        if let Some(rich_cells) = self.inner.sheets_rich.get_mut(&sheet) {
            rich_cells.remove(&changing_cell);
        }
        if let Some(sheet_cells) = self.inner.sheets.get_mut(&sheet) {
            sheet_cells.insert(changing_cell.clone(), solution_json);
        }

        let key = FormulaCellKey::new(sheet.clone(), changing_cell_ref);
        self.inner.pending_spill_clears.remove(&key);
        self.inner.pending_formula_baselines.remove(&key);

        let out = Object::new();
        object_set(&out, "success", &JsValue::from_bool(result.success()))?;
        object_set(&out, "status", &JsValue::from_str(&format!("{:?}", result.status)))?;
        object_set(&out, "solution", &JsValue::from_f64(result.solution))?;
        object_set(
            &out,
            "iterations",
            &JsValue::from_f64(result.iterations as f64),
        )?;
        object_set(&out, "finalError", &JsValue::from_f64(result.final_error))?;
        object_set(&out, "finalOutput", &JsValue::from_f64(result.final_output))?;
        Ok(out.into())
    }

    #[wasm_bindgen(js_name = "getPivotSchema")]
    pub fn get_pivot_schema(
        &self,
        sheet: String,
        source_range_a1: String,
        sample_size: Option<u32>,
    ) -> Result<JsValue, JsValue> {
        ensure_rust_constructors_run();
        let sample_size = sample_size.map(|s| s as usize).unwrap_or(20);
        let schema = self
            .inner
            .get_pivot_schema_internal(&sheet, &source_range_a1, sample_size)?;
        serde_wasm_bindgen::to_value(&schema).map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "calculatePivot")]
    pub fn calculate_pivot(
        &self,
        sheet: String,
        source_range_a1: String,
        destination_top_left_a1: String,
        config: JsValue,
    ) -> Result<JsValue, JsValue> {
        ensure_rust_constructors_run();
        let config: formula_model::pivots::PivotConfig =
            serde_wasm_bindgen::from_value(config).map_err(|err| js_err(err.to_string()))?;
        let engine_config = pivot_config_model_to_engine(&config);
        let writes = self.inner.calculate_pivot_writes_internal(
            &sheet,
            &source_range_a1,
            &destination_top_left_a1,
            &engine_config,
        )?;

        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct PivotCalculationResultDto {
            writes: Vec<CellChange>,
        }

        serde_wasm_bindgen::to_value(&PivotCalculationResultDto { writes })
            .map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "getPivotFieldItems")]
    pub fn get_pivot_field_items(
        &self,
        sheet: String,
        source_range_a1: String,
        field: String,
    ) -> Result<JsValue, JsValue> {
        ensure_rust_constructors_run();
        let sheet = self.inner.require_sheet(&sheet)?.to_string();
        let range = WorkbookState::parse_range(&source_range_a1)?;
        let source = self.inner.read_range_values_as_pivot_values(&sheet, &range);
        let cache =
            pivot_engine::PivotCache::from_range(&source).map_err(|err| js_err(err.to_string()))?;

        let Some(values) = cache.unique_values.get(&field) else {
            return Err(js_err(format!("missing field in pivot cache: {field}")));
        };

        use serde::ser::Serialize as _;
        values
            .serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "getPivotFieldItemsPaged")]
    pub fn get_pivot_field_items_paged(
        &self,
        sheet: String,
        source_range_a1: String,
        field: String,
        offset: u32,
        limit: u32,
    ) -> Result<JsValue, JsValue> {
        ensure_rust_constructors_run();
        let sheet = self.inner.require_sheet(&sheet)?.to_string();
        let range = WorkbookState::parse_range(&source_range_a1)?;
        let source = self.inner.read_range_values_as_pivot_values(&sheet, &range);
        let cache =
            pivot_engine::PivotCache::from_range(&source).map_err(|err| js_err(err.to_string()))?;

        let Some(values) = cache.unique_values.get(&field) else {
            return Err(js_err(format!("missing field in pivot cache: {field}")));
        };

        let start = offset as usize;
        let end = start.saturating_add(limit as usize).min(values.len());
        let slice: &[pivot_engine::PivotValue] = if start >= values.len() {
            &[]
        } else {
            &values[start..end]
        };

        use serde::ser::Serialize as _;
        slice
            .serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "recalculate")]
    pub fn recalculate(&mut self, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let changes = self.inner.recalculate_internal(sheet.as_deref())?;
        let out = Array::new();
        for change in changes {
            out.push(&cell_change_to_js(&change)?);
        }
        Ok(out.into())
    }

    #[wasm_bindgen(js_name = "applyOperation")]
    pub fn apply_operation(&mut self, op: JsValue) -> Result<JsValue, JsValue> {
        let op: EditOpDto =
            serde_wasm_bindgen::from_value(op).map_err(|err| js_err(err.to_string()))?;
        let result = self.inner.apply_operation_internal(op)?;
        serde_wasm_bindgen::to_value(&result).map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "defaultSheetName")]
    pub fn default_sheet_name() -> String {
        DEFAULT_SHEET.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn set_cell_rich_entity_roundtrips_and_degrades_in_get_cell() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let entity = CellValue::Entity(formula_model::EntityValue::new("Acme"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, entity);
        assert_eq!(rich.value, rich.input);
        assert_eq!(
            serde_json::to_value(&rich.input).unwrap(),
            json!({"type":"entity","value":{"displayValue":"Acme"}})
        );

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, JsonValue::Null);
        assert_eq!(scalar.value, json!("Acme"));
    }

    #[test]
    fn set_cell_rich_error_field_degrades_in_get_cell() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_rich_internal(
            DEFAULT_SHEET,
            "A1",
            CellValue::Error(formula_model::ErrorValue::Field),
        )
        .unwrap();

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(
            serde_json::to_value(&rich.input).unwrap(),
            json!({"type":"error","value":"#FIELD!"})
        );

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.value, json!("#FIELD!"));
    }

    #[test]
    fn engine_value_to_json_degrades_rich_values_to_display_string() {
        // The JS worker protocol expects scalar-ish JSON values today. Rich values like
        // entities/records should degrade to their display strings so existing callers never have
        // to handle structured JSON objects.
        let entity = EngineValue::Entity(formula_engine::value::EntityValue::new("Apple Inc."));
        assert_eq!(engine_value_to_json(entity), json!("Apple Inc."));

        let record = EngineValue::Record(formula_engine::value::RecordValue::new("My record"));
        assert_eq!(engine_value_to_json(record), json!("My record"));
    }

    #[test]
    fn engine_value_to_json_arrays_use_top_left_value() {
        let arr = formula_engine::value::Array::new(
            2,
            2,
            vec![
                EngineValue::Number(1.0),
                EngineValue::Number(2.0),
                EngineValue::Number(3.0),
                EngineValue::Number(4.0),
            ],
        );
        assert_eq!(engine_value_to_json(EngineValue::Array(arr)), json!(1.0));
    }

    #[test]
    fn cell_value_to_engine_preserves_extended_error_field() {
        let value = CellValue::Error(formula_model::ErrorValue::Field);
        let engine_value = cell_value_to_engine(&value);
        assert_eq!(engine_value, EngineValue::Error(ErrorKind::Field));
        assert_eq!(engine_value_to_json(engine_value), json!("#FIELD!"));
    }

    #[test]
    fn cell_value_to_engine_preserves_extended_error_connect() {
        let value = CellValue::Error(formula_model::ErrorValue::Connect);
        let engine_value = cell_value_to_engine(&value);
        assert_eq!(engine_value, EngineValue::Error(ErrorKind::Connect));
        assert_eq!(engine_value_to_json(engine_value), json!("#CONNECT!"));
    }

    #[test]
    fn set_cell_rich_entity_properties_flow_through_to_field_access_formulas() {
        // Note: the full public WASM interface surface (`setCellRich`/`getCellRich`) is exercised
        // in `tests/wasm.rs` under `wasm32`. Native unit tests cannot construct JS objects via
        // `serde_wasm_bindgen::to_value` because it requires JS host imports.
        let mut wb = WorkbookState::new_with_default_sheet();

        let mut properties = BTreeMap::new();
        properties.insert("Price".to_string(), CellValue::Number(12.5));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1.Price"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(b1.value, json!(12.5));

        let a1_rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(a1_rich.input, entity);
        assert_eq!(a1_rich.value, a1_rich.input);
    }

    #[test]
    fn set_cell_rich_supports_bracketed_field_access_for_special_characters() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let mut properties = BTreeMap::new();
        properties.insert("Change%".to_string(), CellValue::Number(0.0133));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity)
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!(r#"=A1.["Change%"]"#))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(b1.value, json!(0.0133));
    }

    #[test]
    fn set_cell_rich_supports_nested_field_access_through_record_properties() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let mut record_fields = BTreeMap::new();
        record_fields.insert("Name".to_string(), CellValue::String("Alice".to_string()));
        record_fields.insert("Age".to_string(), CellValue::Number(42.0));
        let owner = CellValue::Record(formula_model::RecordValue {
            fields: record_fields,
            display_field: Some("Name".to_string()),
            display_value: String::new(),
        });

        let mut properties = BTreeMap::new();
        properties.insert("Owner".to_string(), owner);
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity)
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1.Owner.Age"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(b1.value, json!(42.0));
    }

    #[test]
    fn set_cell_rich_field_access_returns_field_error_for_missing_properties() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let mut properties = BTreeMap::new();
        properties.insert("Price".to_string(), CellValue::Number(12.5));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity)
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1.Nope"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(b1.value, json!("#FIELD!"));
    }

    #[test]
    fn set_cell_rich_accepts_cell_value_schema_for_entities() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let typed = CellValue::Entity(
            formula_model::EntityValue::new("Apple Inc.")
                .with_entity_type("stock")
                .with_entity_id("AAPL")
                .with_property("Price", 12.5),
        );

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", typed.clone())
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1.Price"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(b1.value, json!(12.5));

        let a1 = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(a1.input, JsonValue::Null);
        assert_eq!(a1.value, json!("Apple Inc."));

        assert_eq!(
            wb.sheets_rich
                .get(DEFAULT_SHEET)
                .and_then(|cells| cells.get("A1")),
            Some(&typed)
        );
    }

    #[test]
    fn set_cell_rich_accepts_cell_value_schema_for_scalars_by_degrading_to_scalar_io() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", CellValue::Number(42.0))
            .unwrap();

        let cell = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(cell.input, json!(42.0));
        assert_eq!(cell.value, json!(42.0));

        // Rich edits preserve the typed schema entry for `getCellRich.input`.
        assert_eq!(
            wb.sheets_rich
                .get(DEFAULT_SHEET)
                .and_then(|cells| cells.get("A1")),
            Some(&CellValue::Number(42.0))
        );
    }

    #[test]
    fn set_cell_rich_rich_text_roundtrips_input_and_degrades_value() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let rich_text = formula_model::RichText::from_segments(vec![(
            "Hello".to_string(),
            formula_model::rich_text::RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        )]);
        let input = CellValue::RichText(rich_text.clone());

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", input.clone())
            .unwrap();

        let cell = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(cell.input, input);
        assert_eq!(cell.value, CellValue::String("Hello".to_string()));

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, json!("Hello"));
        assert_eq!(scalar.value, json!("Hello"));
    }

    #[test]
    fn set_cell_rich_image_roundtrips_and_degrades_in_get_cell() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let image = CellValue::Image(formula_model::ImageValue {
            image_id: formula_model::drawings::ImageId::new("image1.png"),
            alt_text: Some("Logo".to_string()),
            width: None,
            height: None,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", image.clone())
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, JsonValue::Null);
        assert_eq!(scalar.value, json!("Logo"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, image);
        assert_eq!(rich.value, CellValue::String("Logo".to_string()));
    }

    #[test]
    fn set_cell_rich_array_roundtrips_but_engine_degrades_to_spill_error() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let array = CellValue::Array(formula_model::ArrayValue {
            data: vec![vec![CellValue::Number(1.0), CellValue::Number(2.0)]],
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", array.clone())
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, JsonValue::Null);
        assert_eq!(scalar.value, json!("#SPILL!"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, array);
        assert_eq!(
            rich.value,
            CellValue::Error(formula_model::ErrorValue::Spill)
        );
    }

    #[test]
    fn set_cell_rich_spill_marker_roundtrips_but_engine_degrades_to_spill_error() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let spill = CellValue::Spill(formula_model::SpillValue {
            origin: CellRef::new(0, 0),
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", spill.clone())
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, JsonValue::Null);
        assert_eq!(scalar.value, json!("#SPILL!"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, spill);
        assert_eq!(
            rich.value,
            CellValue::Error(formula_model::ErrorValue::Spill)
        );
    }

    #[test]
    fn set_cell_rich_overwrites_existing_scalar_input() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(5.0))
            .unwrap();
        let before = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(before.input, json!(5.0));
        assert_eq!(before.value, json!(5.0));

        let entity = CellValue::Entity(formula_model::EntityValue::new("Acme"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();

        let after = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(after.input, JsonValue::Null);
        assert_eq!(after.value, json!("Acme"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, entity);
        assert_eq!(rich.value, rich.input);
    }

    #[test]
    fn set_cell_overwrites_existing_rich_input() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let entity = CellValue::Entity(formula_model::EntityValue::new("Acme"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity)
            .unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(5.0))
            .unwrap();

        let cell = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(cell.input, json!(5.0));
        assert_eq!(cell.value, json!(5.0));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, CellValue::Number(5.0));
        assert_eq!(rich.value, CellValue::Number(5.0));
    }

    #[test]
    fn set_cell_rich_empty_clears_previous_rich_value() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let entity = CellValue::Entity(formula_model::EntityValue::new("Acme"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity)
            .unwrap();

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", CellValue::Empty)
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, JsonValue::Null);
        assert_eq!(scalar.value, JsonValue::Null);

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, CellValue::Empty);
        assert_eq!(rich.value, CellValue::Empty);
    }

    #[test]
    fn set_cell_rich_string_preserves_error_like_text_via_quote_prefix() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_rich_internal(
            DEFAULT_SHEET,
            "A1",
            CellValue::String("#FIELD!".to_string()),
        )
        .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, json!("'#FIELD!"));
        assert_eq!(scalar.value, json!("#FIELD!"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, CellValue::String("#FIELD!".to_string()));
        assert_eq!(rich.value, rich.input);
    }

    #[test]
    fn set_cell_rich_string_preserves_formula_like_text_via_quote_prefix() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", CellValue::String("=1+1".to_string()))
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, json!("'=1+1"));
        assert_eq!(scalar.value, json!("=1+1"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, CellValue::String("=1+1".to_string()));
        assert_eq!(rich.value, rich.input);
    }

    #[test]
    fn set_cell_rich_string_preserves_leading_apostrophe_by_double_prefixing_input() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", CellValue::String("'hello".to_string()))
            .unwrap();

        let scalar = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(scalar.input, json!("''hello"));
        assert_eq!(scalar.value, json!("'hello"));

        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, CellValue::String("'hello".to_string()));
        assert_eq!(rich.value, rich.input);
    }

    #[test]
    fn cell_value_json_roundtrips_entity_and_record() {
        let mut record_fields = BTreeMap::new();
        record_fields.insert("Name".to_string(), CellValue::String("Alice".to_string()));
        record_fields.insert("Age".to_string(), CellValue::Number(42.0));

        let record = formula_model::RecordValue {
            fields: record_fields,
            display_field: Some("Name".to_string()),
            display_value: String::new(),
        };

        let mut properties = BTreeMap::new();
        properties.insert("Price".to_string(), CellValue::Number(178.5));
        properties.insert("Owner".to_string(), CellValue::Record(record));

        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple Inc.".to_string(),
            properties,
        });

        let json_value = serde_json::to_value(&entity).unwrap();
        let roundtripped: CellValue = serde_json::from_value(json_value).unwrap();
        assert_eq!(roundtripped, entity);
    }

    #[test]
    fn set_cell_rich_does_not_pollute_scalar_workbook_schema() {
        let mut wb = WorkbookState::new_with_default_sheet();
        let mut properties = BTreeMap::new();
        properties.insert("Price".to_string(), CellValue::Number(178.5));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple Inc.".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();

        // Scalar getCell should keep returning scalar inputs/values.
        let cell = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(cell.input, JsonValue::Null);
        assert_eq!(cell.value, json!("Apple Inc."));

        // Rich getter should roundtrip the full payload.
        let rich = wb.get_cell_rich_data(DEFAULT_SHEET, "A1").unwrap();
        assert_eq!(rich.input, entity);

        // Rich inputs are not representable in the scalar workbook JSON schema.
        assert!(wb.sheets.get(DEFAULT_SHEET).unwrap().get("A1").is_none());
    }

    #[test]
    fn parse_formula_partial_uses_utf16_cursor_and_spans() {
        // Emoji (`😀`) is a surrogate pair in UTF-16 (2 code units) but 4 bytes in UTF-8.
        // Ensure cursor positions expressed as UTF-16 offsets do not panic when slicing, and that
        // returned spans are also expressed in UTF-16 code units.
        let formula = "=\"😀";
        let cursor_utf16 = formula.encode_utf16().count() as u32;

        let byte_cursor = utf16_cursor_to_byte_index(formula, cursor_utf16);
        assert_eq!(byte_cursor, formula.len());

        let prefix = &formula[..byte_cursor];
        let parsed =
            formula_engine::parse_formula_partial(prefix, formula_engine::ParseOptions::default());
        let err = parsed
            .error
            .expect("expected unterminated string literal error");
        assert_eq!(err.message, "Unterminated string literal");

        let span_start = byte_index_to_utf16_cursor(prefix, err.span.start);
        let span_end = byte_index_to_utf16_cursor(prefix, err.span.end);
        assert_eq!(span_start, 1);
        assert_eq!(span_end, cursor_utf16 as usize);
    }

    #[test]
    fn utf16_cursor_conversion_clamps_out_of_range_and_surrogate_midpoints() {
        let formula = "=\"😀\"";
        let formula_utf16_len = formula.encode_utf16().count() as u32;

        // Cursor beyond the end clamps to the end.
        let byte_cursor =
            utf16_cursor_to_byte_index(formula, formula_utf16_len.saturating_add(100));
        assert_eq!(byte_cursor, formula.len());

        // Cursor in the middle of a surrogate pair should clamp to a valid UTF-8 boundary.
        // UTF-16 layout: '=' (1), '\"' (1), 😀 (2), '\"' (1)
        // Cursor=3 lands between the two UTF-16 code units for 😀.
        let byte_cursor_mid = utf16_cursor_to_byte_index(formula, 3);
        assert_eq!(&formula[..byte_cursor_mid], "=\"");
    }

    #[test]
    fn lex_formula_emits_utf16_spans_for_emoji() {
        let formula = "=\"😀\"";
        let (expr_src, span_offset) = formula
            .strip_prefix('=')
            .map(|rest| (rest, 1usize))
            .unwrap_or((formula, 0usize));

        let tokens = formula_engine::lex(expr_src, &formula_engine::ParseOptions::default())
            .expect("lexing should succeed");
        let string_token = tokens
            .iter()
            .find(|t| matches!(&t.kind, formula_engine::TokenKind::String(_)))
            .expect("expected a string token");

        let start = byte_index_to_utf16_cursor(formula, string_token.span.start + span_offset);
        let end = byte_index_to_utf16_cursor(formula, string_token.span.end + span_offset);
        assert_eq!(start, 1);
        assert_eq!(end, formula.encode_utf16().count());
    }

    #[test]
    fn fallback_context_scanner_counts_args_in_unterminated_string() {
        let ctx = scan_fallback_function_context(r#"=SUM(1,"hello"#, ',').unwrap();
        assert_eq!(ctx.name, "SUM");
        assert_eq!(ctx.arg_index, 1);
    }

    #[test]
    fn fallback_context_scanner_handles_unterminated_quoted_identifier() {
        let ctx = scan_fallback_function_context("=SUM('My Sheet", ',').unwrap();
        assert_eq!(ctx.name, "SUM");
        assert_eq!(ctx.arg_index, 0);
    }

    #[test]
    fn fallback_context_scanner_ignores_commas_in_brackets_with_escaped_close() {
        let ctx = scan_fallback_function_context("=FOO([a]],b],1", ',').unwrap();
        assert_eq!(ctx.name, "FOO");
        assert_eq!(ctx.arg_index, 1);
    }

    #[test]
    fn get_cell_data_degrades_engine_rich_values_to_display_string_and_chains() {
        use formula_engine::eval::CellAddr;
        use formula_engine::functions::{Reference, SheetId};

        let mut wb = WorkbookState::new_with_default_sheet();

        // Set a rich engine value directly into the engine cell store.
        let rich_value = EngineValue::Reference(Reference {
            sheet_id: SheetId::Local(0),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr { row: 0, col: 0 },
        });
        let expected = rich_value.to_string();
        wb.engine
            .set_cell_value(DEFAULT_SHEET, "A1", rich_value)
            .unwrap();

        // Ensure a formula that references the rich value produces the same degraded display output.
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let a1 = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();

        assert_eq!(a1.value, json!(expected));
        assert_eq!(b1.value, a1.value);
    }

    #[test]
    fn get_cell_data_degrades_model_entity_and_record_values_to_display_string_and_chains() {
        // Ensure we degrade model Entity/Record variants to display strings at the scalar JSON
        // protocol boundary.
        let entity: CellValue = serde_json::from_value(json!({
            "type": "entity",
            "value": {
                "displayValue": "Entity display"
            }
        }))
        .expect("entity CellValue should deserialize");

        let record: CellValue = serde_json::from_value(json!({
            "type": "record",
            "value": {
                "displayField": "name",
                "fields": {
                    "name": { "type": "string", "value": "Alice" },
                    "age": { "type": "number", "value": 42.0 }
                }
            }
        }))
        .expect("record CellValue should deserialize");

        let mut wb = WorkbookState::new_with_default_sheet();

        let entity_engine = cell_value_to_engine(&entity);
        let entity_expected = entity_engine.to_string();
        wb.engine
            .set_cell_value(DEFAULT_SHEET, "A1", entity_engine)
            .unwrap();

        let record_engine = cell_value_to_engine(&record);
        let record_expected = record_engine.to_string();
        wb.engine
            .set_cell_value(DEFAULT_SHEET, "A2", record_engine)
            .unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!("=A2"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        let a1 = wb.get_cell_data(DEFAULT_SHEET, "A1").unwrap();
        let b1 = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert_eq!(a1.value, json!(entity_expected));
        assert_eq!(b1.value, a1.value);

        let a2 = wb.get_cell_data(DEFAULT_SHEET, "A2").unwrap();
        let b2 = wb.get_cell_data(DEFAULT_SHEET, "B2").unwrap();
        assert_eq!(a2.value, json!(record_expected));
        assert_eq!(b2.value, a2.value);
    }

    #[test]
    fn cell_value_to_engine_maps_field_error() {
        let value = CellValue::Error(formula_model::ErrorValue::Field);
        assert_eq!(
            cell_value_to_engine(&value),
            EngineValue::Error(ErrorKind::Field)
        );
        assert_eq!(
            engine_value_to_json(EngineValue::Error(ErrorKind::Field)),
            json!("#FIELD!")
        );
    }

    #[test]
    fn cell_value_to_json_degrades_image_values_deterministically() {
        // The scalar JSON protocol does not support structured rich values yet. Image values
        // should degrade to a stable string for callers (UI, IPC).
        let image: CellValue = match serde_json::from_value(json!({
            "type": "image",
            "value": {
                "imageId": "image1.png",
                "altText": "Logo"
            }
        })) {
            Ok(value) => value,
            // Older versions of `formula-model` won't have the Image variant yet.
            Err(_) => return,
        };

        assert_eq!(
            engine_value_to_json(cell_value_to_engine(&image)),
            json!("Logo")
        );

        let image_no_alt: CellValue = match serde_json::from_value(json!({
            "type": "image",
            "value": {
                "imageId": "image1.png"
            }
        })) {
            Ok(value) => value,
            Err(_) => return,
        };
        assert_eq!(
            engine_value_to_json(cell_value_to_engine(&image_no_alt)),
            json!("[Image]")
        );
    }

    #[test]
    fn recalculate_includes_spill_output_cells() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,2)"))
            .unwrap();

        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "A1".to_string(),
                    value: json!(1.0),
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B1".to_string(),
                    value: json!(2.0),
                },
            ]
        );
    }

    #[test]
    fn recalculate_reports_spill_clears_when_spill_origin_is_edited() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,2)"))
            .unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=1"))
            .unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "A1".to_string(),
                    value: json!(1.0),
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B1".to_string(),
                    value: JsonValue::Null,
                },
            ]
        );
    }

    #[test]
    fn recalculate_reports_spill_clears_when_spill_cell_is_overwritten() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,3)"))
            .unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!(5.0))
            .unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
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

    #[test]
    fn recalculate_reports_formula_edit_to_blank_value() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=1"))
            .unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=A2"))
            .unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: JsonValue::Null,
            }]
        );
    }

    #[test]
    fn recalculate_does_not_filter_changes_by_sheet_argument() {
        let mut wb = WorkbookState::new_with_default_sheet();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("=A1*2"))
            .unwrap();

        wb.set_cell_internal("Sheet2", "A1", json!(10.0)).unwrap();
        wb.set_cell_internal("Sheet2", "A2", json!("=A1*2"))
            .unwrap();

        wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(2.0))
            .unwrap();
        wb.set_cell_internal("Sheet2", "A1", json!(11.0)).unwrap();

        // The wasm API accepts a `sheet` argument for symmetry, but recalc deltas are always
        // workbook-wide. Unknown sheet names should be ignored.
        let changes = wb.recalculate_internal(Some("MissingSheet")).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: "Sheet1".to_string(),
                    address: "A2".to_string(),
                    value: json!(4.0),
                },
                CellChange {
                    sheet: "Sheet2".to_string(),
                    address: "A2".to_string(),
                    value: json!(22.0),
                },
            ]
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn to_json_preserves_engine_workbook_schema() {
        let input = json!({
            "sheets": {
                "Sheet1": {
                    "cells": {
                        "A1": 1.0,
                        "A2": "=A1*2"
                    }
                }
            }
        })
        .to_string();

        let wb = WasmWorkbook::from_json(&input).unwrap();
        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!(1.0));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));

        let wb2 = WasmWorkbook::from_json(&json_str).unwrap();
        let json_str2 = wb2.to_json().unwrap();
        let parsed2: serde_json::Value = serde_json::from_str(&json_str2).unwrap();
        assert_eq!(parsed2["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));
    }

    #[test]
    fn from_xlsx_bytes_imports_tables_for_structured_reference_formulas() {
        let bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsx/tests/fixtures/table.xlsx"
        ));

        let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "D2"),
            EngineValue::Number(6.0)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "E1"),
            EngineValue::Number(20.0)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "F1"),
            EngineValue::Text("Qty".into())
        );
    }

    #[test]
    fn from_xlsx_bytes_preserves_modern_error_values_as_engine_errors() {
        let bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/basic/bool-error.xlsx"
        ));
        let wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Error(ErrorKind::Div0)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Error(ErrorKind::Field)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "D1"),
            EngineValue::Error(ErrorKind::Connect)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "E1"),
            EngineValue::Error(ErrorKind::Blocked)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "F1"),
            EngineValue::Error(ErrorKind::Unknown)
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_encodes_literal_text_inputs_that_look_like_formulas_or_errors() {
        use std::io::Cursor;

        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1").unwrap();
        let sheet = workbook.sheet_mut(sheet_id).unwrap();
        sheet
            .set_value_a1("A1", CellValue::String("=hello".to_string()))
            .unwrap();
        sheet
            .set_value_a1("A2", CellValue::String("'hello".to_string()))
            .unwrap();
        sheet
            .set_value_a1("A3", CellValue::String("#REF!".to_string()))
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text("=hello".to_string())
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Text("'hello".to_string())
        );

        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        // These values must be quote-prefixed in the workbook JSON input map so `fromJson`
        // round-trips preserve them as literal text (not formulas/errors).
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!("'=hello"));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("''hello"));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A3"], json!("'#REF!"));
    }

    #[test]
    fn localized_formula_input_is_canonicalized_and_persisted() {
        let mut wb = WasmWorkbook::new();
        assert!(wb.set_locale("de-DE".to_string()));

        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", json!("=SUMME(1;2)"))
            .unwrap();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A2", json!("=1,5+1"))
            .unwrap();

        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(3.0)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Number(2.5)
        );

        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        assert_eq!(
            parsed["sheets"]["Sheet1"]["cells"]["A1"],
            json!("=SUM(1,2)")
        );
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("=1.5+1"));
    }

    #[test]
    fn canonicalize_and_localize_formula_roundtrip_de_de() {
        let localized = "=SUMME(1,5;2)";
        let canonical = canonicalize_formula(localized, "de-DE", None).unwrap();
        assert_eq!(canonical, "=SUM(1.5,2)");

        let roundtrip = localize_formula(&canonical, "de-DE", None).unwrap();
        assert_eq!(roundtrip, localized);
    }

    #[test]
    fn canonicalize_and_localize_formula_roundtrip_fr_fr() {
        let localized = "=SOMME(1,5;2)";
        let canonical = canonicalize_formula(localized, "fr-FR", None).unwrap();
        assert_eq!(canonical, "=SUM(1.5,2)");

        let roundtrip = localize_formula(&canonical, "fr-FR", None).unwrap();
        assert_eq!(roundtrip, localized);
    }

    #[test]
    fn canonicalize_and_localize_formula_roundtrip_r1c1_reference_style() {
        let localized = "=SUMME(R1C1;R1C2)";
        let canonical = canonicalize_formula(localized, "de-DE", Some("R1C1".to_string())).unwrap();
        assert_eq!(canonical, "=SUM(R1C1,R1C2)");

        let roundtrip = localize_formula(&canonical, "de-DE", Some("R1C1".to_string())).unwrap();
        assert_eq!(roundtrip, localized);
    }

    #[test]
    fn sheet_dimensions_expand_whole_column_references() {
        let mut wb = WasmWorkbook::new();

        // Expand the default sheet to include row 2,000,000.
        wb.set_sheet_dimensions(DEFAULT_SHEET.to_string(), 2_100_000, EXCEL_MAX_COLS)
            .unwrap();

        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A2000000", json!(5.0))
            .unwrap();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "B1", json!("=SUM(A:A)"))
            .unwrap();

        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Number(5.0)
        );
    }

    #[test]
    fn apply_operation_insert_rows_updates_literal_cells_and_formulas() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::InsertRows {
                sheet: DEFAULT_SHEET.to_string(),
                row: 0,
                count: 1,
            })
            .unwrap();

        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Number(1.0)
        );
        assert_eq!(wb.engine.get_cell_formula(DEFAULT_SHEET, "B2"), Some("=A2"));

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A2"), Some(&json!(1.0)));
        assert_eq!(sheet_cells.get("B2"), Some(&json!("=A2")));
        assert!(!sheet_cells.contains_key("A1"));
        assert!(!sheet_cells.contains_key("B1"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B2".to_string(),
                before: "=A1".to_string(),
                after: "=A2".to_string(),
            }),
            "expected formula rewrite for moved formula cell"
        );

        // Workbook JSON should reflect the updated sparse input map.
        let wb = WasmWorkbook { inner: wb };
        let exported = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&exported).unwrap();
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!(1.0));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["B2"], json!("=A2"));
        assert!(parsed["sheets"]["Sheet1"]["cells"].get("A1").is_none());
        assert!(parsed["sheets"]["Sheet1"]["cells"].get("B1").is_none());
    }

    #[test]
    fn apply_operation_delete_cols_updates_inputs_and_formulas() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!(2.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!("=A1+B1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::DeleteCols {
                sheet: DEFAULT_SHEET.to_string(),
                col: 0,
                count: 1,
            })
            .unwrap();

        // B1 shifts left to A1.
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(2.0)
        );
        // Formula cell shifts left to B1 and its A1 reference becomes #REF!.
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "B1"),
            Some("=#REF!+A1")
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A1"), Some(&json!(2.0)));
        assert_eq!(sheet_cells.get("B1"), Some(&json!("=#REF!+A1")));
        assert!(!sheet_cells.contains_key("C1"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                before: "=A1+B1".to_string(),
                after: "=#REF!+A1".to_string(),
            }),
            "expected formula rewrite for shifted formula cell"
        );

        let wb = WasmWorkbook { inner: wb };
        let exported = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&exported).unwrap();
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!(2.0));
        assert_eq!(
            parsed["sheets"]["Sheet1"]["cells"]["B1"],
            json!("=#REF!+A1")
        );
        assert!(parsed["sheets"]["Sheet1"]["cells"].get("C1").is_none());
    }

    #[test]
    fn apply_operation_insert_cells_shift_right_moves_cells_and_rewrites_references() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!(3.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "D1", json!("=A1+C1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::InsertCellsShiftRight {
                sheet: DEFAULT_SHEET.to_string(),
                range: "A1:B1".to_string(),
            })
            .unwrap();

        // A1 moved to C1, and C1 moved to E1.
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Number(1.0)
        );
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "E1"),
            EngineValue::Number(3.0)
        );
        // Formula moved from D1 -> F1 and should track the moved cells.
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "F1"),
            Some("=C1+E1")
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("C1"), Some(&json!(1.0)));
        assert_eq!(sheet_cells.get("E1"), Some(&json!(3.0)));
        assert_eq!(sheet_cells.get("F1"), Some(&json!("=C1+E1")));
        assert!(!sheet_cells.contains_key("A1"));
        assert!(!sheet_cells.contains_key("D1"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "F1".to_string(),
                before: "=A1+C1".to_string(),
                after: "=C1+E1".to_string(),
            }),
            "expected formula rewrite for shifted formula cell"
        );
    }

    #[test]
    fn apply_operation_delete_cells_shift_left_creates_ref_errors_and_updates_shifted_references() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!(2.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!(3.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "D1", json!(4.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "E1", json!("=A1+D1"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("=B1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::DeleteCellsShiftLeft {
                sheet: DEFAULT_SHEET.to_string(),
                range: "B1:C1".to_string(),
            })
            .unwrap();

        // D1 moved into B1.
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Number(4.0)
        );
        // Formula moved from E1 -> C1 and should track the moved cell (D1 -> B1).
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C1"),
            Some("=A1+B1")
        );
        // Reference into deleted region becomes #REF!, even though another cell moved into B1.
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "A2"),
            Some("=#REF!")
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A1"), Some(&json!(1.0)));
        assert_eq!(sheet_cells.get("B1"), Some(&json!(4.0)));
        assert_eq!(sheet_cells.get("C1"), Some(&json!("=A1+B1")));
        assert_eq!(sheet_cells.get("A2"), Some(&json!("=#REF!")));
        assert!(!sheet_cells.contains_key("D1"));
        assert!(!sheet_cells.contains_key("E1"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                before: "=A1+D1".to_string(),
                after: "=A1+B1".to_string(),
            }),
            "expected formula rewrite for shifted formula cell"
        );
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "A2".to_string(),
                before: "=B1".to_string(),
                after: "=#REF!".to_string(),
            }),
            "expected formula rewrite for deleted reference"
        );
    }

    #[test]
    fn apply_operation_insert_cells_shift_down_rewrites_references_into_shifted_region() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(42.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::InsertCellsShiftDown {
                sheet: DEFAULT_SHEET.to_string(),
                range: "A1".to_string(),
            })
            .unwrap();

        // A1 moved down to A2; formula should follow it.
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Number(42.0)
        );
        assert_eq!(wb.engine.get_cell_formula(DEFAULT_SHEET, "B1"), Some("=A2"));

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A2"), Some(&json!(42.0)));
        assert_eq!(sheet_cells.get("B1"), Some(&json!("=A2")));
        assert!(!sheet_cells.contains_key("A1"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                before: "=A1".to_string(),
                after: "=A2".to_string(),
            }),
            "expected formula rewrite for shifted reference"
        );
    }

    #[test]
    fn apply_operation_delete_cells_shift_up_rewrites_moved_references_and_invalidates_deleted_targets(
    ) {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!(3.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A3"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!("=A2"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::DeleteCellsShiftUp {
                sheet: DEFAULT_SHEET.to_string(),
                range: "A1:A2".to_string(),
            })
            .unwrap();

        // A3 moved up to A1; B1 should follow that move.
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(3.0)
        );
        assert_eq!(wb.engine.get_cell_formula(DEFAULT_SHEET, "B1"), Some("=A1"));

        // Reference directly into deleted region becomes #REF!
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "B2"),
            Some("=#REF!")
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A1"), Some(&json!(3.0)));
        assert_eq!(sheet_cells.get("B1"), Some(&json!("=A1")));
        assert_eq!(sheet_cells.get("B2"), Some(&json!("=#REF!")));
        assert!(!sheet_cells.contains_key("A3"));

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                before: "=A3".to_string(),
                after: "=A1".to_string(),
            }),
            "expected formula rewrite for shifted reference"
        );
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B2".to_string(),
                before: "=A2".to_string(),
                after: "=#REF!".to_string(),
            }),
            "expected formula rewrite for deleted reference"
        );
    }

    #[test]
    fn cell_value_to_engine_converts_entity_and_record_values() {
        let mut record_fields = BTreeMap::new();
        record_fields.insert("Name".to_string(), CellValue::String("Alice".to_string()));
        record_fields.insert("Active".to_string(), CellValue::Boolean(true));
        let record = CellValue::Record(formula_model::RecordValue {
            fields: record_fields,
            display_field: Some("Name".to_string()),
            ..formula_model::RecordValue::default()
        });

        let mut properties = BTreeMap::new();
        properties.insert("Person".to_string(), record);
        properties.insert("Score".to_string(), CellValue::Number(10.0));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "user".to_string(),
            entity_id: "alice".to_string(),
            display_value: "Alice".to_string(),
            properties,
        });

        let engine_value = cell_value_to_engine(&entity);
        let entity = match engine_value {
            EngineValue::Entity(entity) => entity,
            other => panic!("expected EngineValue::Entity, got {other:?}"),
        };
        assert_eq!(entity.entity_type.as_deref(), Some("user"));
        assert_eq!(entity.entity_id.as_deref(), Some("alice"));
        assert_eq!(entity.display, "Alice");
        assert!(matches!(
            entity.fields.get("Score"),
            Some(&EngineValue::Number(n)) if n == 10.0
        ));

        let record = match entity.fields.get("Person") {
            Some(EngineValue::Record(record)) => record,
            other => panic!("expected nested EngineValue::Record, got {other:?}"),
        };
        assert_eq!(record.display_field.as_deref(), Some("Name"));
        assert_eq!(
            record.fields.get("Name"),
            Some(&EngineValue::Text("Alice".to_string()))
        );
        assert_eq!(record.fields.get("Active"), Some(&EngineValue::Bool(true)));
    }

    #[test]
    fn apply_operation_preserves_quote_prefixed_text_inputs() {
        let mut wb = WorkbookState::new_with_default_sheet();

        // Literal text that looks like a formula.
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("'=hello"))
            .unwrap();
        // Literal text beginning with an apostrophe (must be double-escaped in inputs).
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("''hello"))
            .unwrap();

        wb.apply_operation_internal(EditOpDto::InsertRows {
            sheet: DEFAULT_SHEET.to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Text("=hello".to_string())
        );
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A3"),
            EngineValue::Text("'hello".to_string())
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A2"), Some(&json!("'=hello")));
        assert_eq!(sheet_cells.get("A3"), Some(&json!("''hello")));
        assert!(!sheet_cells.contains_key("A1"));
    }

    #[test]
    fn apply_operation_move_range_updates_inputs_and_returns_moved_ranges() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(42.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!("=A1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::MoveRange {
                sheet: DEFAULT_SHEET.to_string(),
                src: "A1:B1".to_string(),
                dst_top_left: "A2".to_string(),
            })
            .unwrap();

        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A2"),
            EngineValue::Number(42.0)
        );
        assert_eq!(wb.engine.get_cell_formula(DEFAULT_SHEET, "B2"), Some("=A2"));
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C1"),
            Some("=A2"),
            "formulas outside the moved range should follow the moved cells"
        );
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Blank
        );
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Blank
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("A2"), Some(&json!(42.0)));
        assert_eq!(sheet_cells.get("B2"), Some(&json!("=A2")));
        assert_eq!(sheet_cells.get("C1"), Some(&json!("=A2")));
        assert!(!sheet_cells.contains_key("A1"));
        assert!(!sheet_cells.contains_key("B1"));

        assert_eq!(
            result.moved_ranges,
            vec![EditMovedRangeDto {
                sheet: DEFAULT_SHEET.to_string(),
                from: "A1:B1".to_string(),
                to: "A2:B2".to_string(),
            }]
        );

        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B2".to_string(),
                before: "=A1".to_string(),
                after: "=A2".to_string(),
            }),
            "expected formula rewrite for moved formula cell"
        );
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                before: "=A1".to_string(),
                after: "=A2".to_string(),
            }),
            "expected formula rewrite for external reference"
        );
    }

    #[test]
    fn apply_operation_move_range_remaps_rich_inputs_and_rewrites_field_access_formulas() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let mut properties = BTreeMap::new();
        properties.insert("Price".to_string(), CellValue::Number(12.5));
        let entity = CellValue::Entity(formula_model::EntityValue {
            entity_type: "stock".to_string(),
            entity_id: "AAPL".to_string(),
            display_value: "Apple Inc.".to_string(),
            properties,
        });

        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!("=A1.Price"))
            .unwrap();

        wb.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Number(12.5)
        );

        wb.apply_operation_internal(EditOpDto::MoveRange {
            sheet: DEFAULT_SHEET.to_string(),
            src: "A1".to_string(),
            dst_top_left: "B2".to_string(),
        })
        .unwrap();

        // Rich input should move along with the cell.
        assert_eq!(
            wb.sheets_rich
                .get(DEFAULT_SHEET)
                .and_then(|cells| cells.get("B2")),
            Some(&entity)
        );
        assert!(wb
            .sheets_rich
            .get(DEFAULT_SHEET)
            .and_then(|cells| cells.get("A1"))
            .is_none());

        // Rich values remain absent from the scalar workbook schema.
        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert!(sheet_cells.get("B2").is_none());

        // Formulas outside the moved range should follow the moved rich value.
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C1"),
            Some("=B2.Price")
        );
        assert_eq!(sheet_cells.get("C1"), Some(&json!("=B2.Price")));

        // Rich getter should round-trip the value at the new address.
        let rich_b2 = wb.get_cell_rich_data(DEFAULT_SHEET, "B2").unwrap();
        assert_eq!(rich_b2.input, entity);
        assert_eq!(rich_b2.value, rich_b2.input);

        wb.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Number(12.5)
        );
    }

    #[test]
    fn apply_operation_copy_range_adjusts_relative_references() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::CopyRange {
                sheet: DEFAULT_SHEET.to_string(),
                src: "B1".to_string(),
                dst_top_left: "B2".to_string(),
            })
            .unwrap();

        assert_eq!(wb.engine.get_cell_formula(DEFAULT_SHEET, "B1"), Some("=A1"));
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "B2"),
            Some("=A2"),
            "copied formulas should adjust relative references to the new location"
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("B1"), Some(&json!("=A1")));
        assert_eq!(sheet_cells.get("B2"), Some(&json!("=A2")));

        assert!(result.moved_ranges.is_empty());
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "B2".to_string(),
                before: "=A1".to_string(),
                after: "=A2".to_string(),
            }),
            "expected formula rewrite for copied formula cell"
        );
    }

    #[test]
    fn apply_operation_copy_range_copies_rich_inputs_and_overwrites_destination() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let src_entity = CellValue::Entity(formula_model::EntityValue::new("Source"));
        let dst_entity = CellValue::Entity(formula_model::EntityValue::new("Destination"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", src_entity.clone())
            .unwrap();
        wb.set_cell_rich_internal(DEFAULT_SHEET, "B1", dst_entity)
            .unwrap();

        wb.apply_operation_internal(EditOpDto::CopyRange {
            sheet: DEFAULT_SHEET.to_string(),
            src: "A1".to_string(),
            dst_top_left: "B1".to_string(),
        })
        .unwrap();

        let rich_cells = wb.sheets_rich.get(DEFAULT_SHEET).unwrap();
        assert_eq!(rich_cells.get("A1"), Some(&src_entity));
        assert_eq!(
            rich_cells.get("B1"),
            Some(&src_entity),
            "destination rich input should be overwritten by the copy"
        );
    }

    #[test]
    fn apply_operation_insert_rows_remaps_rich_inputs() {
        let mut wb = WorkbookState::new_with_default_sheet();

        let entity = CellValue::Entity(formula_model::EntityValue::new("Acme"));
        wb.set_cell_rich_internal(DEFAULT_SHEET, "A1", entity.clone())
            .unwrap();

        wb.apply_operation_internal(EditOpDto::InsertRows {
            sheet: DEFAULT_SHEET.to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

        let rich_cells = wb.sheets_rich.get(DEFAULT_SHEET).unwrap();
        assert!(
            rich_cells.get("A1").is_none(),
            "rich input should shift down with inserted rows"
        );
        assert_eq!(rich_cells.get("A2"), Some(&entity));
    }

    #[test]
    fn apply_operation_fill_repeats_formulas_and_updates_relative_references() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "C1", json!("=A1+B1"))
            .unwrap();

        let result = wb
            .apply_operation_internal(EditOpDto::Fill {
                sheet: DEFAULT_SHEET.to_string(),
                src: "C1".to_string(),
                dst: "C1:C3".to_string(),
            })
            .unwrap();

        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C1"),
            Some("=A1+B1")
        );
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C2"),
            Some("=A2+B2")
        );
        assert_eq!(
            wb.engine.get_cell_formula(DEFAULT_SHEET, "C3"),
            Some("=A3+B3")
        );

        let sheet_cells = wb.sheets.get(DEFAULT_SHEET).unwrap();
        assert_eq!(sheet_cells.get("C1"), Some(&json!("=A1+B1")));
        assert_eq!(sheet_cells.get("C2"), Some(&json!("=A2+B2")));
        assert_eq!(sheet_cells.get("C3"), Some(&json!("=A3+B3")));

        assert!(result.moved_ranges.is_empty());
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "C2".to_string(),
                before: "=A1+B1".to_string(),
                after: "=A2+B2".to_string(),
            }),
            "expected formula rewrite for filled cell C2"
        );
        assert!(
            result.formula_rewrites.contains(&EditFormulaRewriteDto {
                sheet: DEFAULT_SHEET.to_string(),
                address: "C3".to_string(),
                before: "=A1+B1".to_string(),
                after: "=A3+B3".to_string(),
            }),
            "expected formula rewrite for filled cell C3"
        );
    }

    #[test]
    fn apply_operation_clears_stale_spill_outputs_on_next_recalc() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,2)"))
            .unwrap();
        wb.recalculate_internal(None).unwrap();

        // Ensure the spill output cell exists as a cached value (not an input).
        let b1_before = wb.get_cell_data(DEFAULT_SHEET, "B1").unwrap();
        assert!(b1_before.input.is_null());
        assert_eq!(b1_before.value, json!(2.0));

        wb.apply_operation_internal(EditOpDto::InsertRows {
            sheet: DEFAULT_SHEET.to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

        // The spill output at B1 should be cleared even though spill metadata was reset during the
        // edit and the next recalc will spill into B2.
        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B1".to_string(),
                    value: JsonValue::Null,
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "A2".to_string(),
                    value: json!(1.0),
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B2".to_string(),
                    value: json!(2.0),
                },
            ]
        );
    }

    #[test]
    fn calculate_pivot_returns_cell_writes_for_basic_row_sum() {
        let mut wb = WorkbookState::new_with_default_sheet();

        // Source data (headers + records).
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("Category"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("Amount"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("A")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(10.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!("A")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B3", json!(5.0)).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A4", json!("B")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B4", json!(7.0)).unwrap();

        // No formulas, but run a recalc to mirror typical usage where pivots reflect calculated
        // values.
        wb.recalculate_internal(None).unwrap();

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
            // Match Excel: no "Grand Total" column when there are no column fields.
            grand_totals: formula_model::pivots::GrandTotals {
                rows: true,
                columns: false,
            },
        };

        let engine_config = pivot_config_model_to_engine(&config);
        let writes = wb
            .calculate_pivot_writes_internal(DEFAULT_SHEET, "A1:B4", "D1", &engine_config)
            .unwrap();

        let expected = vec![
            ("D1", JsonValue::String("Category".to_string())),
            ("E1", JsonValue::String("Sum of Amount".to_string())),
            ("D2", JsonValue::String("A".to_string())),
            ("E2", json!(15.0)),
            ("D3", JsonValue::String("B".to_string())),
            ("E3", json!(7.0)),
            ("D4", JsonValue::String("Grand Total".to_string())),
            ("E4", json!(22.0)),
        ];

        assert_eq!(
            writes.len(),
            expected.len(),
            "expected {expected:?}, got {writes:?}"
        );

        let mut got_by_address: HashMap<String, JsonValue> = HashMap::new();
        for w in writes {
            assert_eq!(w.sheet, DEFAULT_SHEET);
            got_by_address.insert(w.address, w.value);
        }

        for (addr, expected_value) in expected {
            let got = got_by_address
                .get(addr)
                .unwrap_or_else(|| panic!("missing write for {addr}, got {got_by_address:?}"));
            assert_eq!(
                got, &expected_value,
                "unexpected value for {addr}: got {got:?}, expected {expected_value:?}"
            );
        }
    }
}
