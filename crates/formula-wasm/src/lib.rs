use std::collections::{BTreeMap, BTreeSet, HashMap};

use formula_engine::calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
use formula_engine::{
    metadata::FormatRun as EngineFormatRun,
    CellAddr, Coord, EditError as EngineEditError, EditOp as EngineEditOp,
    EditResult as EngineEditResult, Engine, EngineInfo, ErrorKind, NameDefinition,
    NameScope, ParseOptions, Span as EngineSpan, Token, TokenKind, Value as EngineValue,
};
use formula_engine::editing::rewrite::rewrite_formula_for_copy_delta;
use formula_engine::locale::{
    canonicalize_formula_with_style, get_locale, localize_formula_with_style, FormulaLocale,
    iter_locales, text_codepage_for_locale_id, ValueLocaleConfig, EN_US,
};
use formula_engine::pivot as pivot_engine;
use formula_engine::what_if::{
    goal_seek::{GoalSeek, GoalSeekParams, GoalSeekResult},
    CellRef as WhatIfCellRef, CellValue as WhatIfCellValue, WhatIfError, WhatIfModel,
};
use formula_model::{
    display_formula_text, Alignment, CellRef, CellValue, Color, DateSystem, DefinedNameScope, Font,
    HorizontalAlignment, Protection, Range, SheetVisibility, Style, TabColor, VerticalAlignment,
    EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};
use js_sys::{Array, Object, Reflect};
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use unicode_normalization::UnicodeNormalization;
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

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PivotCellWrite {
    pub sheet: String,
    pub address: String,
    pub value: JsonValue,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct GoalSeekRequestDto {
    target_cell: String,
    target_value: f64,
    changing_cell: String,
    #[serde(default)]
    sheet: Option<String>,
    #[serde(default)]
    max_iterations: Option<u32>,
    #[serde(default)]
    tolerance: Option<f64>,
    #[serde(default)]
    derivative_step: Option<f64>,
    #[serde(default)]
    min_derivative: Option<f64>,
    #[serde(default)]
    max_bracket_expansions: Option<u32>,
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct GoalSeekResponseDto {
    result: GoalSeekResult,
    changes: Vec<CellChange>,
}

#[derive(Clone, Debug, Default)]
struct GoalSeekTuning {
    max_iterations: Option<usize>,
    tolerance: Option<f64>,
    derivative_step: Option<f64>,
    min_derivative: Option<f64>,
    max_bracket_expansions: Option<usize>,
}

fn js_err(message: impl ToString) -> JsValue {
    JsValue::from_str(&message.to_string())
}

const OFFICE_CRYPTO_ERROR_PREFIX: &str = "OFFICE_CRYPTO_ERROR:";

fn office_crypto_kind_and_message(
    err: &formula_office_crypto::OfficeCryptoError,
) -> (&'static str, String) {
    // The TS worker RPC surface currently transports errors as strings only. Encode Office
    // encryption errors as a tagged JSON payload so callers can distinguish common cases
    // programmatically (password required vs invalid password vs unsupported encryption).
    let kind = match &err {
        formula_office_crypto::OfficeCryptoError::PasswordRequired => "PasswordRequired",
        formula_office_crypto::OfficeCryptoError::InvalidPassword
        | formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed => "InvalidPassword",
        formula_office_crypto::OfficeCryptoError::SpinCountTooLarge { .. } => "SpinCountTooLarge",
        formula_office_crypto::OfficeCryptoError::UnsupportedEncryption(_) => "UnsupportedEncryption",
        formula_office_crypto::OfficeCryptoError::InvalidOptions(_) => "InvalidOptions",
        formula_office_crypto::OfficeCryptoError::InvalidFormat(_) => "InvalidFormat",
        formula_office_crypto::OfficeCryptoError::SizeLimitExceeded { .. } => "SizeLimitExceeded",
        formula_office_crypto::OfficeCryptoError::SizeLimitExceededU64 { .. } => "SizeLimitExceeded",
        formula_office_crypto::OfficeCryptoError::EncryptedPackageSizeOverflow { .. } => {
            "EncryptedPackageSizeOverflow"
        }
        formula_office_crypto::OfficeCryptoError::EncryptedPackageAllocationFailed { .. } => {
            "EncryptedPackageAllocationFailed"
        }
        formula_office_crypto::OfficeCryptoError::Io(_) => "Io",
    };
    let message = match &err {
        // Treat integrity mismatches as invalid passwords so callers can re-prompt. This matches the
        // `formula-io` and desktop semantics (integrity mismatch = retryable invalid password).
        formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed => {
            formula_office_crypto::OfficeCryptoError::InvalidPassword.to_string()
        }
        _ => err.to_string(),
    };
    (kind, message)
}

fn office_crypto_err(err: formula_office_crypto::OfficeCryptoError) -> JsValue {
    let (kind, message) = office_crypto_kind_and_message(&err);
    let payload = serde_json::json!({
        "kind": kind,
        "message": message,
    });
    JsValue::from_str(&format!("{OFFICE_CRYPTO_ERROR_PREFIX}{payload}"))
}

#[cfg(test)]
mod office_crypto_err_tests {
    use super::*;

    #[test]
    fn integrity_check_failed_is_invalid_password() {
        let (kind, message) = office_crypto_kind_and_message(
            &formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed,
        );
        assert_eq!(kind, "InvalidPassword");
        assert_eq!(message, "invalid password");
    }
}

fn js_value_to_object(value: &JsValue) -> Option<Object> {
    if value.is_null() || value.is_undefined() {
        return None;
    }
    value.clone().dyn_into::<Object>().ok()
}

fn has_js_prop(obj: &Object, key: &str) -> bool {
    Reflect::has(obj, &JsValue::from_str(key)).unwrap_or(false)
}

fn get_js_prop(obj: &Object, key: &str) -> Option<JsValue> {
    Reflect::get(obj, &JsValue::from_str(key))
        .ok()
        .filter(|v| !v.is_null() && !v.is_undefined())
}

fn get_js_prop_raw(obj: &Object, key: &str) -> Option<JsValue> {
    Reflect::get(obj, &JsValue::from_str(key)).ok()
}

fn get_js_bool(obj: &Object, keys: &[&str]) -> Option<bool> {
    for key in keys {
        let value = match get_js_prop(obj, key) {
            Some(v) => v,
            None => continue,
        };
        if let Some(b) = value.as_bool() {
            return Some(b);
        }
    }
    None
}

fn get_js_string(obj: &Object, keys: &[&str]) -> Option<String> {
    for key in keys {
        let value = match get_js_prop(obj, key) {
            Some(v) => v,
            None => continue,
        };
        if let Some(s) = value.as_string() {
            return Some(s);
        }
    }
    None
}

fn parse_tint_thousandths(value: f64) -> Option<i16> {
    if !value.is_finite() {
        return None;
    }

    let thousandths = if value.abs() <= 1.0 {
        // OOXML `tint` is typically a float in [-1.0, 1.0].
        (value.clamp(-1.0, 1.0) * 1000.0).round()
    } else {
        // Some UI payloads may already express tint in thousandths.
        value.clamp(-1000.0, 1000.0).round()
    };

    Some(thousandths as i16)
}

fn parse_color_string(raw: &str) -> Option<Color> {
    let s = raw.trim();
    let hex = s.strip_prefix('#').unwrap_or(s);
    match hex.len() {
        6 => u32::from_str_radix(hex, 16)
            .ok()
            .map(|rgb| Color::new_argb(0xFF00_0000 | rgb)),
        8 => u32::from_str_radix(hex, 16).ok().map(Color::new_argb),
        _ => None,
    }
}

fn parse_color_from_js(value: &JsValue) -> Option<Color> {
    if let Some(s) = value.as_string() {
        return parse_color_string(&s);
    }

    let obj = js_value_to_object(value)?;

    if let Some(rgb) = get_js_string(&obj, &["rgb", "argb"]) {
        return parse_color_string(&rgb);
    }

    let auto = get_js_prop(&obj, "auto")
        .and_then(|v| {
            v.as_bool()
                .or_else(|| v.as_f64().map(|n| n.is_finite() && n != 0.0))
        })
        .unwrap_or(false);
    if auto {
        return Some(Color::Auto);
    }

    if let Some(theme) = get_js_prop(&obj, "theme")
        .and_then(|v| v.as_f64())
        .and_then(|n| {
            if n.is_finite() && n >= 0.0 && n <= u16::MAX as f64 {
                Some(n.trunc() as u16)
            } else {
                None
            }
        })
    {
        let tint = get_js_prop(&obj, "tint")
            .and_then(|v| v.as_f64())
            .and_then(parse_tint_thousandths);
        return Some(Color::Theme { theme, tint });
    }

    if let Some(indexed) = get_js_prop(&obj, "indexed")
        .and_then(|v| v.as_f64())
        .and_then(|n| {
            if n.is_finite() && n >= 0.0 && n <= u16::MAX as f64 {
                Some(n.trunc() as u16)
            } else {
                None
            }
        })
    {
        return Some(Color::Indexed(indexed));
    }

    None
}

fn parse_alignment_from_js(value: &JsValue) -> Option<Alignment> {
    let obj = js_value_to_object(value)?;

    // For patch semantics, we need to distinguish:
    // - missing key: no override
    // - key present with null/undefined: explicit clear back to General
    let horizontal = if has_js_prop(&obj, "horizontal") {
        let raw = get_js_prop_raw(&obj, "horizontal").unwrap_or(JsValue::UNDEFINED);
        if raw.is_null() || raw.is_undefined() {
            Some(HorizontalAlignment::General)
        } else if let Some(raw) = raw.as_string() {
            match raw.trim().to_lowercase().as_str() {
                "general" => Some(HorizontalAlignment::General),
                "left" => Some(HorizontalAlignment::Left),
                "center" | "centre" => Some(HorizontalAlignment::Center),
                "right" => Some(HorizontalAlignment::Right),
                "fill" => Some(HorizontalAlignment::Fill),
                "justify" => Some(HorizontalAlignment::Justify),
                _ => None,
            }
        } else {
            None
        }
    } else {
        None
    };

    let vertical = get_js_string(&obj, &["vertical"]).and_then(|raw| {
        match raw.trim().to_lowercase().as_str() {
            "top" => Some(VerticalAlignment::Top),
            "center" => Some(VerticalAlignment::Center),
            "bottom" => Some(VerticalAlignment::Bottom),
            _ => None,
        }
    });

    let wrap_text = get_js_bool(&obj, &["wrapText", "wrap_text"]).unwrap_or(false);

    let rotation = get_js_prop(&obj, "rotation")
        .and_then(|v| v.as_f64())
        .map(|n| n.trunc() as i16);

    let indent = get_js_prop(&obj, "indent")
        .and_then(|v| v.as_f64())
        .map(|n| n.trunc() as u16);

    let out = Alignment {
        horizontal,
        vertical,
        wrap_text,
        rotation,
        indent,
    };

    let is_default = out.horizontal.is_none()
        && out.vertical.is_none()
        && !out.wrap_text
        && out.rotation.is_none()
        && out.indent.is_none();
    if is_default {
        None
    } else {
        Some(out)
    }
}

fn parse_protection_from_js(value: &JsValue) -> Option<Protection> {
    let obj = js_value_to_object(value)?;
    let has_locked = has_js_prop(&obj, "locked");
    let has_hidden = has_js_prop(&obj, "hidden");
    if !has_locked && !has_hidden {
        return None;
    }

    let mut explicit_default = false;

    let locked = if has_locked {
        let raw = get_js_prop_raw(&obj, "locked").unwrap_or(JsValue::UNDEFINED);
        if raw.is_null() || raw.is_undefined() {
            explicit_default = true;
            true
        } else {
            raw.as_bool().unwrap_or(true)
        }
    } else {
        true
    };

    let hidden = if has_hidden {
        let raw = get_js_prop_raw(&obj, "hidden").unwrap_or(JsValue::UNDEFINED);
        if raw.is_null() || raw.is_undefined() {
            explicit_default = true;
            false
        } else {
            raw.as_bool().unwrap_or(false)
        }
    } else {
        false
    };

    // Avoid interning redundant "default" protection structs; Excel default is locked.
    if locked && !hidden && !explicit_default {
        return None;
    }

    Some(Protection { locked, hidden })
}

fn parse_font_from_js(value: &JsValue, top_level_strike: Option<bool>) -> Option<Font> {
    let obj = js_value_to_object(value)?;

    let name = get_js_string(&obj, &["name"]);
    let size_100pt = get_js_prop(&obj, "size_100pt")
        .or_else(|| get_js_prop(&obj, "size100pt"))
        .and_then(|v| v.as_f64())
        .and_then(|n| {
            if n.is_finite() && n >= 0.0 && n <= u16::MAX as f64 {
                Some(n.trunc() as u16)
            } else {
                None
            }
        });

    let bold = get_js_bool(&obj, &["bold"]).unwrap_or(false);
    let italic = get_js_bool(&obj, &["italic"]).unwrap_or(false);
    let underline = get_js_bool(&obj, &["underline"]).unwrap_or(false);
    let strike = get_js_bool(&obj, &["strike"]).unwrap_or(top_level_strike.unwrap_or(false));

    let color = get_js_prop(&obj, "color")
        .as_ref()
        .and_then(parse_color_from_js);

    let out = Font {
        name,
        size_100pt,
        bold,
        italic,
        underline,
        strike,
        color,
    };

    let is_default = out.name.is_none()
        && out.size_100pt.is_none()
        && !out.bold
        && !out.italic
        && !out.underline
        && !out.strike
        && out.color.is_none();

    if is_default {
        None
    } else {
        Some(out)
    }
}

fn parse_style_from_js(style: JsValue) -> Result<Style, JsValue> {
    if style.is_null() || style.is_undefined() {
        return Ok(Style::default());
    }

    let obj = style
        .dyn_into::<Object>()
        .map_err(|_| js_err("internStyle: style must be an object"))?;

    let number_format = if has_js_prop(&obj, "numberFormat") || has_js_prop(&obj, "number_format") {
        let key = if has_js_prop(&obj, "numberFormat") {
            "numberFormat"
        } else {
            "number_format"
        };
        let raw = get_js_prop_raw(&obj, key).unwrap_or(JsValue::UNDEFINED);
        if raw.is_null() || raw.is_undefined() {
            Some("General".to_string())
        } else if let Some(raw) = raw.as_string() {
            let trimmed = raw.trim();
            if trimmed.is_empty() {
                None
            } else if trimmed.eq_ignore_ascii_case("general") {
                // Excel treats "General" as the default.
                None
            } else {
                Some(raw)
            }
        } else {
            None
        }
    } else {
        None
    };

    let top_level_strike = get_js_bool(&obj, &["strike"]);
    let mut font = get_js_prop(&obj, "font")
        .as_ref()
        .and_then(|value| parse_font_from_js(value, top_level_strike));

    // Some UI payloads flatten `strike` and/or font color at the top level.
    let top_level_color = get_js_prop(&obj, "fontColor")
        .or_else(|| get_js_prop(&obj, "font_color"))
        .as_ref()
        .and_then(parse_color_from_js);

    if font.is_none() && top_level_strike.unwrap_or(false) {
        font = Some(Font {
            strike: true,
            ..Default::default()
        });
    }

    if let Some(color) = top_level_color {
        font = Some(match font {
            Some(mut existing) => {
                if existing.color.is_none() {
                    existing.color = Some(color);
                }
                existing
            }
            None => Font {
                color: Some(color),
                strike: top_level_strike.unwrap_or(false),
                ..Default::default()
            },
        });
    }

    let alignment = get_js_prop(&obj, "alignment")
        .as_ref()
        .and_then(parse_alignment_from_js);

    let protection = get_js_prop(&obj, "protection")
        .as_ref()
        .and_then(parse_protection_from_js)
        .or_else(|| {
            // Some UI payloads flatten protection flags at the top level (`{ locked: false }`).
            // Treat those as an alias for `{ protection: { ... } }`.
            let has_locked = has_js_prop(&obj, "locked");
            let has_hidden = has_js_prop(&obj, "hidden");
            if !has_locked && !has_hidden {
                return None;
            }

            let mut explicit_default = false;
            let locked = if has_locked {
                let raw = get_js_prop_raw(&obj, "locked").unwrap_or(JsValue::UNDEFINED);
                if raw.is_null() || raw.is_undefined() {
                    explicit_default = true;
                    true
                } else {
                    raw.as_bool().unwrap_or(true)
                }
            } else {
                true
            };

            let hidden = if has_hidden {
                let raw = get_js_prop_raw(&obj, "hidden").unwrap_or(JsValue::UNDEFINED);
                if raw.is_null() || raw.is_undefined() {
                    explicit_default = true;
                    false
                } else {
                    raw.as_bool().unwrap_or(false)
                }
            } else {
                false
            };

            if locked && !hidden && !explicit_default {
                None
            } else {
                Some(Protection { locked, hidden })
            }
        });

    Ok(Style {
        font,
        fill: None,
        border: None,
        alignment,
        protection,
        number_format,
    })
}

/// Best-effort conversion from UI formatting JSON (typically camelCase) into a `formula_model::Style`.
///
/// This helper is used by native (non-wasm) unit tests because `js_sys` object construction is
/// unavailable on non-wasm targets.
#[cfg(test)]
fn style_json_to_model_style(value: &JsonValue) -> Style {
    fn parse_color_from_json(value: &JsonValue) -> Option<Color> {
        match value {
            JsonValue::String(s) => parse_color_string(s),
            JsonValue::Object(obj) => {
                if let Some(rgb) = obj.get("rgb").and_then(|v| v.as_str()) {
                    return parse_color_string(rgb);
                }
                if let Some(argb) = obj.get("argb").and_then(|v| v.as_str()) {
                    return parse_color_string(argb);
                }
                if obj
                    .get("auto")
                    .and_then(|v| v.as_bool().or_else(|| v.as_f64().map(|n| n.is_finite() && n != 0.0)))
                    .unwrap_or(false)
                {
                    return Some(Color::Auto);
                }
                if let Some(theme) = obj.get("theme").and_then(|v| v.as_f64()).and_then(|n| {
                    if n.is_finite() && n >= 0.0 && n <= u16::MAX as f64 {
                        Some(n.trunc() as u16)
                    } else {
                        None
                    }
                }) {
                    let tint = obj
                        .get("tint")
                        .and_then(|v| v.as_f64())
                        .and_then(parse_tint_thousandths);
                    return Some(Color::Theme { theme, tint });
                }
                if let Some(indexed) = obj
                    .get("indexed")
                    .and_then(|v| v.as_f64())
                    .and_then(|n| {
                        if n.is_finite() && n >= 0.0 && n <= u16::MAX as f64 {
                            Some(n.trunc() as u16)
                        } else {
                            None
                        }
                    })
                {
                    return Some(Color::Indexed(indexed));
                }
                None
            }
            _ => None,
        }
    }

    fn normalize_color_json(value: &JsonValue) -> Option<JsonValue> {
        let color = parse_color_from_json(value)?;
        serde_json::to_value(color).ok()
    }

    // Prefer parsing as a full `Style` first (preserves font/fill/border when the caller supplies
    // the snake_case schema), but overlay UI-friendly camelCase mappings so callers can provide
    // `numberFormat`, `{ protection: { locked } }`, etc.
    let mut sanitized = value.clone();
    if let Some(obj) = sanitized.as_object_mut() {
        // Normalize UI/OOXML-like font color payloads into `formula_model::Color`-compatible shapes
        // (e.g. `{ rgb: "FFRRGGBB" }` -> "#FFRRGGBB") so `serde_json` parsing succeeds.
        if let Some(font) = obj.get_mut("font").and_then(|v| v.as_object_mut()) {
            if let Some(color_value) = font.get("color").cloned() {
                if let Some(normalized) = normalize_color_json(&color_value) {
                    font.insert("color".to_string(), normalized);
                } else {
                    font.remove("color");
                }
            }
        }

        // Some UI payloads use `{ fontColor: ... }` at the top level instead of `font.color`.
        let top_level_color = obj
            .get("fontColor")
            .or_else(|| obj.get("font_color"))
            .cloned()
            .and_then(|v| normalize_color_json(&v));
        if let Some(color) = top_level_color {
            let font = obj
                .entry("font".to_string())
                .or_insert_with(|| JsonValue::Object(Default::default()));
            if let Some(font_obj) = font.as_object_mut() {
                font_obj.entry("color".to_string()).or_insert(color);
            }
        }
    }

    let mut out: Style = serde_json::from_value(sanitized).unwrap_or_default();

    let Some(obj) = value.as_object() else {
        return out;
    };

    // --- number_format ---
    if obj.contains_key("numberFormat") || obj.contains_key("number_format") {
        let raw = obj
            .get("numberFormat")
            .or_else(|| obj.get("number_format"))
            .unwrap_or(&JsonValue::Null);
        match raw {
            JsonValue::Null => {
                // Explicit `null` means "clear to General" (even if a lower layer sets a format).
                out.number_format = Some("General".to_string());
            }
            JsonValue::String(s) => {
                let trimmed = s.trim();
                if trimmed.is_empty() {
                    out.number_format = None;
                } else if trimmed.eq_ignore_ascii_case("general") {
                    // Excel treats "General" as the default.
                    out.number_format = None;
                } else {
                    out.number_format = Some(trimmed.to_string());
                }
            }
            _ => {}
        }
    }

    // --- protection ---
    let protection = obj.get("protection").and_then(|v| v.as_object());
    let protection_has_locked = protection.is_some_and(|p| p.contains_key("locked"));
    let protection_has_hidden = protection.is_some_and(|p| p.contains_key("hidden"));
    let top_level_has_locked = obj.contains_key("locked");
    let top_level_has_hidden = obj.contains_key("hidden");

    if protection_has_locked || protection_has_hidden || top_level_has_locked || top_level_has_hidden {
        let mut explicit_default = false;

        let locked_raw = protection
            .and_then(|p| p.get("locked"))
            .or_else(|| obj.get("locked"));
        let locked = match locked_raw {
            Some(JsonValue::Null) => {
                explicit_default = true;
                true
            }
            Some(JsonValue::Bool(b)) => *b,
            _ => true,
        };

        let hidden_raw = protection
            .and_then(|p| p.get("hidden"))
            .or_else(|| obj.get("hidden"));
        let hidden = match hidden_raw {
            Some(JsonValue::Null) => {
                explicit_default = true;
                false
            }
            Some(JsonValue::Bool(b)) => *b,
            _ => false,
        };

        if locked && !hidden && !explicit_default {
            out.protection = None;
        } else {
            out.protection = Some(Protection { locked, hidden });
        }
    }

    // --- alignment.horizontal ---
    if let Some(alignment) = obj.get("alignment").and_then(|v| v.as_object()) {
        if alignment.contains_key("horizontal") {
            let raw = alignment.get("horizontal").unwrap_or(&JsonValue::Null);
            let horizontal = match raw {
                JsonValue::Null => Some(HorizontalAlignment::General),
                JsonValue::String(s) => match s.trim().to_ascii_lowercase().as_str() {
                    "general" => Some(HorizontalAlignment::General),
                    "left" => Some(HorizontalAlignment::Left),
                    "center" | "centre" => Some(HorizontalAlignment::Center),
                    "right" => Some(HorizontalAlignment::Right),
                    "fill" => Some(HorizontalAlignment::Fill),
                    "justify" => Some(HorizontalAlignment::Justify),
                    _ => None,
                },
                _ => None,
            };
            if horizontal.is_some() {
                let mut next = out.alignment.unwrap_or_default();
                next.horizontal = horizontal;
                out.alignment = Some(next);
            }
        }
    }

    out
}

fn supported_locale_ids_sorted() -> Vec<&'static str> {
    let mut ids: Vec<&'static str> = iter_locales().map(|locale| locale.id).collect();
    ids.sort_unstable();
    ids
}

#[wasm_bindgen(js_name = "supportedLocaleIds")]
pub fn supported_locale_ids() -> JsValue {
    ensure_rust_constructors_run();
    let ids = supported_locale_ids_sorted();
    let out = Array::new();
    for id in ids {
        out.push(&JsValue::from_str(id));
    }
    out.into()
}

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct LocaleInfoDto {
    locale_id: &'static str,
    decimal_separator: String,
    arg_separator: String,
    array_row_separator: String,
    array_col_separator: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    thousands_separator: Option<String>,
    is_rtl: bool,
    boolean_true: &'static str,
    boolean_false: &'static str,
}

#[wasm_bindgen(js_name = "getLocaleInfo")]
pub fn get_locale_info(locale_id: &str) -> Result<JsValue, JsValue> {
    ensure_rust_constructors_run();
    let locale = require_formula_locale(locale_id)?;
    let info = LocaleInfoDto {
        locale_id: locale.id,
        decimal_separator: locale.config.decimal_separator.to_string(),
        arg_separator: locale.config.arg_separator.to_string(),
        array_row_separator: locale.config.array_row_separator.to_string(),
        array_col_separator: locale.config.array_col_separator.to_string(),
        thousands_separator: locale.config.thousands_separator.map(|ch| ch.to_string()),
        is_rtl: locale.is_rtl,
        boolean_true: locale.boolean_true,
        boolean_false: locale.boolean_false,
    };
    serde_wasm_bindgen::to_value(&info).map_err(|err| js_err(err.to_string()))
}

fn require_formula_locale(locale_id: &str) -> Result<&'static FormulaLocale, JsValue> {
    get_locale(locale_id).ok_or_else(|| {
        let supported = supported_locale_ids_sorted().join(", ");
        js_err(format!(
            "unknown localeId: {locale_id}. Supported locale ids: {supported}",
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

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
enum CalcModeDto {
    Automatic,
    AutomaticNoTable,
    Manual,
}

impl From<CalculationMode> for CalcModeDto {
    fn from(mode: CalculationMode) -> Self {
        match mode {
            CalculationMode::Automatic => CalcModeDto::Automatic,
            CalculationMode::AutomaticNoTable => CalcModeDto::AutomaticNoTable,
            CalculationMode::Manual => CalcModeDto::Manual,
        }
    }
}

impl From<CalcModeDto> for CalculationMode {
    fn from(mode: CalcModeDto) -> Self {
        match mode {
            CalcModeDto::Automatic => CalculationMode::Automatic,
            CalcModeDto::AutomaticNoTable => CalculationMode::AutomaticNoTable,
            CalcModeDto::Manual => CalculationMode::Manual,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct IterativeCalcSettingsDto {
    enabled: bool,
    max_iterations: u32,
    max_change: f64,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct IterativeCalcSettingsInputDto {
    enabled: bool,
    max_iterations: f64,
    max_change: f64,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
struct CalcSettingsDto {
    calculation_mode: CalcModeDto,
    calculate_before_save: bool,
    full_precision: bool,
    full_calc_on_load: bool,
    iterative: IterativeCalcSettingsDto,
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct CalcSettingsInputDto {
    calculation_mode: CalcModeDto,
    calculate_before_save: bool,
    full_precision: bool,
    full_calc_on_load: bool,
    iterative: IterativeCalcSettingsInputDto,
}

/// Indicates whether formula strings in the workbook JSON payload are in canonical (en-US) syntax
/// or localized according to `localeId`.
///
/// This is an additive field in the workbook JSON schema consumed/emitted by `WasmWorkbook`
/// (`fromJson`/`toJson`). When absent, `fromJson` preserves legacy behavior: if `localeId` is a
/// non-en-US locale, formula strings are treated as localized and canonicalized during import.
#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
enum WorkbookFormulaLanguageDto {
    /// Canonical (en-US) formula text, using comma argument separators and `.` decimals.
    Canonical,
    /// Locale-dependent formula text, parsed according to the workbook `localeId`.
    Localized,
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
    parse_options_and_locale_from_js(options).map(|(opts, _)| opts)
}

fn parse_options_and_locale_from_js(
    options: Option<JsValue>,
) -> Result<(ParseOptions, Option<&'static FormulaLocale>), JsValue> {
    let Some(value) = options else {
        return Ok((ParseOptions::default(), None));
    };
    if value.is_undefined() || value.is_null() {
        return Ok((ParseOptions::default(), None));
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
        return Ok((ParseOptions::default(), None));
    }

    let has_locale_id = Reflect::has(&obj, &JsValue::from_str("localeId")).unwrap_or(false);
    let has_ref_style = Reflect::has(&obj, &JsValue::from_str("referenceStyle")).unwrap_or(false);
    if has_locale_id || has_ref_style {
        let dto: ParseOptionsJsDto =
            serde_wasm_bindgen::from_value(obj.into()).map_err(|err| js_err(err.to_string()))?;
        let mut opts = ParseOptions::default();
        let mut locale: Option<&'static FormulaLocale> = None;
        if let Some(locale_id) = dto.locale_id {
            let formula_locale = require_formula_locale(&locale_id)?;
            opts.locale = formula_locale.config.clone();
            locale = Some(formula_locale);
        }
        if let Some(style) = dto.reference_style {
            opts.reference_style = style;
        }
        return Ok((opts, locale));
    }

    let looks_like_parse_options = Reflect::has(&obj, &JsValue::from_str("locale"))
        .unwrap_or(false)
        || Reflect::has(&obj, &JsValue::from_str("reference_style")).unwrap_or(false)
        || Reflect::has(&obj, &JsValue::from_str("normalize_relative_to")).unwrap_or(false);
    if looks_like_parse_options {
        // Fall back to the full ParseOptions struct for advanced callers.
        return serde_wasm_bindgen::from_value(obj.into())
            .map(|opts| (opts, None))
            .map_err(|err| js_err(err.to_string()));
    }

    Err(js_err(
        "options must be { localeId?: string, referenceStyle?: \"A1\" | \"R1C1\" } or a ParseOptions object",
    ))
}

fn normalize_function_context_name(name: &str, locale: Option<&FormulaLocale>) -> String {
    let canonical = match locale {
        Some(locale) => locale.canonical_function_name(name),
        None => name.to_ascii_uppercase(),
    };

    const XL_FN_PREFIX: &str = "_xlfn.";
    canonical
        .get(..XL_FN_PREFIX.len())
        .filter(|prefix| prefix.eq_ignore_ascii_case(XL_FN_PREFIX))
        .map(|_| canonical[XL_FN_PREFIX.len()..].to_string())
        .unwrap_or(canonical)
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
    // Match `formula_engine::Workbook::sheet_key` / `formula_model::sheet_name_eq_case_insensitive`.
    //
    // Excel compares sheet names case-insensitively across Unicode and applies compatibility
    // normalization (NFKC). We approximate this by normalizing with Unicode NFKC and then applying
    // Unicode uppercasing (locale-independent).
    //
    // Fast-path ASCII sheet names to avoid the cost of Unicode normalization on the common case.
    if name.is_ascii() {
        return name.to_ascii_uppercase();
    }
    name.nfkc().flat_map(|c| c.to_uppercase()).collect()
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

fn pivot_value_to_json(value: pivot_engine::PivotValue, date_system: formula_engine::date::ExcelDateSystem) -> JsonValue {
    match value {
        pivot_engine::PivotValue::Blank => JsonValue::Null,
        pivot_engine::PivotValue::Bool(b) => JsonValue::Bool(b),
        pivot_engine::PivotValue::Text(s) => JsonValue::String(s),
        pivot_engine::PivotValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        // Cell values in the engine are stored using Excel serial numbers (with styling carrying the
        // number format). Represent pivot dates the same way so callers can apply the accompanying
        // date number format and formulas see Excel-like underlying values.
        pivot_engine::PivotValue::Date(d) => {
            // Avoid pulling in `chrono` as a direct dependency: `NaiveDate`'s component accessors
            // live on the `Datelike` trait. The ISO-8601 `Display` form is stable (`YYYY-MM-DD`),
            // so parse that into an `ExcelDate`.
            let s = d.to_string();
            let mut parts = s.split('-');
            let year = parts.next().and_then(|v| v.parse::<i32>().ok());
            let month = parts.next().and_then(|v| v.parse::<u8>().ok());
            let day = parts.next().and_then(|v| v.parse::<u8>().ok());
            let excel_date = match (year, month, day, parts.next()) {
                (Some(year), Some(month), Some(day), None) => {
                    formula_engine::date::ExcelDate::new(year, month, day)
                }
                _ => return JsonValue::Null,
            };
            match formula_engine::date::ymd_to_serial(excel_date, date_system) {
                Ok(serial) => serde_json::Number::from_f64(serial as f64)
                    .map(JsonValue::Number)
                    .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
                Err(_) => JsonValue::Null,
            }
        }
    }
}

fn pivot_key_part_model_to_engine(
    part: &formula_model::pivots::PivotKeyPart,
) -> pivot_engine::PivotKeyPart {
    match part {
        formula_model::pivots::PivotKeyPart::Blank => pivot_engine::PivotKeyPart::Blank,
        formula_model::pivots::PivotKeyPart::Number(bits) => {
            pivot_engine::PivotKeyPart::Number(*bits)
        }
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

fn pivot_field_model_to_engine(
    field: &formula_model::pivots::PivotField,
) -> pivot_engine::PivotField {
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

fn pivot_value_field_model_to_engine(
    field: &formula_model::pivots::ValueField,
) -> pivot_engine::ValueField {
    pivot_engine::ValueField {
        source_field: field.source_field.clone(),
        name: field.name.clone(),
        aggregation: pivot_aggregation_model_to_engine(field.aggregation),
        number_format: field.number_format.clone(),
        show_as: field.show_as,
        base_field: field.base_field.clone(),
        base_item: field.base_item.clone(),
    }
}

fn pivot_filter_field_model_to_engine(
    field: &formula_model::pivots::FilterField,
) -> pivot_engine::FilterField {
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
    // `SubtotalPosition` is re-exported by the pivot engine from `formula-model`, so this is
    // currently a 1:1 conversion.
    match position {
        formula_model::pivots::SubtotalPosition::None => pivot_engine::SubtotalPosition::None,
        formula_model::pivots::SubtotalPosition::Top => pivot_engine::SubtotalPosition::Top,
        formula_model::pivots::SubtotalPosition::Bottom => pivot_engine::SubtotalPosition::Bottom,
    }
}

fn pivot_config_model_to_engine(
    cfg: &formula_model::pivots::PivotConfig,
) -> pivot_engine::PivotConfig {
    pivot_engine::PivotConfig {
        row_fields: cfg
            .row_fields
            .iter()
            .map(pivot_field_model_to_engine)
            .collect(),
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
    /// Optional sheet visibility metadata (Excel-compatible).
    ///
    /// This is not currently modeled by the calc engine, but we preserve it for UI/workbook
    /// metadata consumers (e.g. `WorkbookInfo.sheets[*].visibility`).
    sheet_visibility: HashMap<String, SheetVisibility>,
    /// Optional sheet tab color metadata (`<sheetPr><tabColor ...>`).
    ///
    /// This is not currently modeled by the calc engine, but we preserve it for UI/workbook
    /// metadata consumers (e.g. `WorkbookInfo.sheets[*].tabColor`).
    sheet_tab_colors: HashMap<String, TabColor>,
    /// Per-sheet per-column width overrides in Excel "character" units (OOXML `col/@width`).
    ///
    /// This is separate from the calc engine's grid state today; it exists to support worksheet
    /// information functions like `CELL("width")` and to preserve imported column widths.
    col_widths_chars: BTreeMap<String, BTreeMap<u32, f32>>,
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

#[derive(Clone, Debug)]
struct GoalSeekModelError(String);

impl std::fmt::Display for GoalSeekModelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(f)
    }
}

impl From<JsValue> for GoalSeekModelError {
    fn from(value: JsValue) -> Self {
        let message = value
            .as_string()
            .unwrap_or_else(|| format!("unexpected JS error value: {value:?}"));
        Self(message)
    }
}

struct WorkbookGoalSeekModel<'a> {
    wb: &'a mut WorkbookState,
    sheet: String,
    changes: BTreeMap<FormulaCellKey, JsonValue>,
}

impl<'a> WorkbookGoalSeekModel<'a> {
    fn new(wb: &'a mut WorkbookState, sheet: String) -> Self {
        Self {
            wb,
            sheet,
            changes: BTreeMap::new(),
        }
    }

    fn push_changes(&mut self, changes: Vec<CellChange>) {
        for change in changes {
            // Engine-driven changes should always have valid A1 addresses; treat parsing failures
            // as a no-op rather than panicking.
            let Ok(cell_ref) = CellRef::from_a1(&change.address) else {
                continue;
            };
            self.changes.insert(
                FormulaCellKey {
                    sheet: change.sheet,
                    row: cell_ref.row,
                    col: cell_ref.col,
                },
                change.value,
            );
        }
    }
}

fn engine_value_to_what_if_value(value: EngineValue) -> WhatIfCellValue {
    match value {
        EngineValue::Number(n) => WhatIfCellValue::Number(n),
        EngineValue::Bool(b) => WhatIfCellValue::Bool(b),
        EngineValue::Text(s) => WhatIfCellValue::Text(s),
        EngineValue::Blank => WhatIfCellValue::Blank,
        EngineValue::Error(err) => WhatIfCellValue::Text(err.as_code().to_string()),
        EngineValue::Entity(entity) => WhatIfCellValue::Text(entity.display),
        EngineValue::Record(record) => WhatIfCellValue::Text(record.display),
        EngineValue::Array(arr) => engine_value_to_what_if_value(arr.top_left()),
        other => WhatIfCellValue::Text(other.to_string()),
    }
}

fn what_if_value_to_json(value: WhatIfCellValue) -> JsonValue {
    match value {
        WhatIfCellValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        WhatIfCellValue::Bool(b) => JsonValue::Bool(b),
        WhatIfCellValue::Text(s) => JsonValue::String(encode_scalar_text_input(&s)),
        WhatIfCellValue::Blank => JsonValue::Null,
    }
}

impl WhatIfModel for WorkbookGoalSeekModel<'_> {
    type Error = GoalSeekModelError;

    fn get_cell_value(&self, cell: &WhatIfCellRef) -> Result<WhatIfCellValue, Self::Error> {
        Ok(engine_value_to_what_if_value(
            self.wb.engine.get_cell_value(&self.sheet, cell.as_str()),
        ))
    }

    fn set_cell_value(&mut self, cell: &WhatIfCellRef, value: WhatIfCellValue) -> Result<(), Self::Error> {
        self.wb
            .set_cell_internal(&self.sheet, cell.as_str(), what_if_value_to_json(value))
            .map_err(GoalSeekModelError::from)
    }

    fn recalculate(&mut self) -> Result<(), Self::Error> {
        let changes = self
            .wb
            .recalculate_internal(None)
            .map_err(GoalSeekModelError::from)?;
        self.push_changes(changes);
        Ok(())
    }
}

#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct FormatRunDto {
    start_row: u32,
    end_row_exclusive: u32,
    style_id: u32,
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
            sheet_visibility: HashMap::new(),
            sheet_tab_colors: HashMap::new(),
            col_widths_chars: BTreeMap::new(),
            pending_spill_clears: BTreeSet::new(),
            pending_formula_baselines: BTreeMap::new(),
        }
    }

    fn new_with_default_sheet() -> Self {
        let mut wb = Self::new_empty();
        wb.ensure_sheet(DEFAULT_SHEET);
        wb
    }

    /// Run `f` with the engine forced into manual calculation mode, restoring the original workbook
    /// calc settings afterwards.
    ///
    /// The WASM worker protocol relies on explicit `recalculate()` calls to produce value-change
    /// deltas. Excel workbooks often default to automatic calculation mode (`calcMode="auto"`), so
    /// we must prevent the engine from performing automatic recalculations during edits or the JS
    /// layer would miss those notifications.
    fn with_manual_calc_mode<T>(
        &mut self,
        f: impl FnOnce(&mut WorkbookState) -> Result<T, JsValue>,
    ) -> Result<T, JsValue> {
        let previous = self.engine.calc_settings().clone();
        if previous.calculation_mode != CalculationMode::Manual {
            let mut manual = previous.clone();
            manual.calculation_mode = CalculationMode::Manual;
            self.engine.set_calc_settings(manual);
        }

        let result = f(self);
        self.engine.set_calc_settings(previous);
        result
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
        self.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(name);
            this.engine
                .set_sheet_dimensions(&sheet, rows, cols)
                .map_err(|err| js_err(err.to_string()))
        })
    }

    fn set_col_width_chars_internal(
        &mut self,
        name: &str,
        col: u32,
        width_chars: Option<f32>,
    ) -> Result<(), JsValue> {
        if col >= EXCEL_MAX_COLS {
            return Err(js_err(format!("col out of Excel bounds: {col}")));
        }

        if let Some(width) = width_chars {
            if !width.is_finite() || width < 0.0 {
                return Err(js_err(
                    "width must be a non-negative finite number".to_string(),
                ));
            }
        }

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(name);

            match width_chars {
                Some(width) => {
                    this.col_widths_chars
                        .entry(sheet.clone())
                        .or_default()
                        .insert(col, width);
                }
                None => {
                    if let Some(cols) = this.col_widths_chars.get_mut(&sheet) {
                        cols.remove(&col);
                        if cols.is_empty() {
                            this.col_widths_chars.remove(&sheet);
                        }
                    }
                }
            }

            // Keep the underlying engine's workbook metadata in sync so worksheet information
            // functions (e.g. `CELL("width")`) can consult column properties.
            this.engine.set_col_width(&sheet, col, width_chars);
            Ok(())
        })
    }

    fn set_workbook_file_metadata_internal(
        &mut self,
        directory: Option<&str>,
        filename: Option<&str>,
    ) -> Result<(), JsValue> {
        // Prevent automatic recalculation when the workbook calc mode is `Automatic`.
        //
        // The WASM worker protocol expects callers to invoke `recalculate()` explicitly so value
        // change deltas can be surfaced over RPC; if we allow an automatic recalc here, JS would
        // miss those notifications.
        self.with_manual_calc_mode(|this| {
            this.engine.set_workbook_file_metadata(directory, filename);
            Ok(())
        })
    }

    fn get_sheet_dimensions_internal(&self, name: &str) -> Result<(u32, u32), JsValue> {
        let sheet = self.require_sheet(name)?;
        self.engine
            .sheet_dimensions(sheet)
            .ok_or_else(|| js_err(format!("missing sheet: {name}")))
    }

    fn set_sheet_display_name_internal(
        &mut self,
        sheet_key: &str,
        display_name: &str,
    ) -> Result<(), JsValue> {
        // Avoid holding an immutable borrow of `sheet_lookup` across the mutable engine call.
        let sheet = self.require_sheet(sheet_key)?.to_string();
        self.engine.set_sheet_display_name(&sheet, display_name);
        Ok(())
    }

    fn set_col_format_runs_internal(
        &mut self,
        sheet: &str,
        col: u32,
        runs: Vec<EngineFormatRun>,
    ) -> Result<(), JsValue> {
        let sheet = self.ensure_sheet(sheet);
        self.engine
            .set_col_format_runs(&sheet, col, runs)
            .map_err(|err| js_err(err.to_string()))
    }

    fn resolve_sheet(&self, name: &str) -> Option<&str> {
        let key = normalize_sheet_key(name);
        self.sheet_lookup.get(&key).map(String::as_str)
    }

    fn require_sheet(&self, name: &str) -> Result<&str, JsValue> {
        self.resolve_sheet(name)
            .ok_or_else(|| js_err(format!("missing sheet: {name}")))
    }

    fn rename_sheet_internal(&mut self, old_name: &str, new_name: &str) -> bool {
        let old_display = match self.resolve_sheet(old_name) {
            Some(name) => name.to_string(),
            None => return false,
        };
        let new_display = new_name.trim();
        if new_display.is_empty() {
            return false;
        }
        if old_display == new_display {
            return true;
        }

        if !self.engine.rename_sheet(&old_display, new_display) {
            return false;
        }
        let new_display = new_display.to_string();

        // Update the case-insensitive sheet name mapping.
        let old_key = normalize_sheet_key(&old_display);
        let new_key = normalize_sheet_key(&new_display);
        if old_key != new_key {
            self.sheet_lookup.remove(&old_key);
        }
        self.sheet_lookup.insert(new_key, new_display.clone());

        // Rename sheet-scoped input maps used by `toJson` / `getCell.input`.
        if let Some(cells) = self.sheets.remove(&old_display) {
            self.sheets.insert(new_display.clone(), cells);
        } else {
            self.sheets.entry(new_display.clone()).or_default();
        }
        if let Some(cells) = self.sheets_rich.remove(&old_display) {
            self.sheets_rich.insert(new_display.clone(), cells);
        } else {
            self.sheets_rich.entry(new_display.clone()).or_default();
        }
        if let Some(cols) = self.col_widths_chars.remove(&old_display) {
            self.col_widths_chars.insert(new_display.clone(), cols);
        }
        if let Some(visibility) = self.sheet_visibility.remove(&old_display) {
            self.sheet_visibility.insert(new_display.clone(), visibility);
        }
        if let Some(color) = self.sheet_tab_colors.remove(&old_display) {
            self.sheet_tab_colors.insert(new_display.clone(), color);
        }

        // Rename pending spill/formula bookkeeping entries so the next recalc tick stays coherent.
        if !self.pending_spill_clears.is_empty() {
            let pending = std::mem::take(&mut self.pending_spill_clears);
            self.pending_spill_clears = pending
                .into_iter()
                .map(|mut key| {
                    if key.sheet == old_display {
                        key.sheet = new_display.clone();
                    }
                    key
                })
                .collect();
        }

        if !self.pending_formula_baselines.is_empty() {
            let pending = std::mem::take(&mut self.pending_formula_baselines);
            let mut next = BTreeMap::new();
            for (mut key, value) in pending {
                if key.sheet == old_display {
                    key.sheet = new_display.clone();
                }
                next.insert(key, value);
            }
            self.pending_formula_baselines = next;
        }

        // Rewrite stored formula inputs so `toJson()` / `getCell.input` match Excel-like rename
        // semantics (and stay consistent with `Engine::rename_sheet`).
        for sheet_cells in self.sheets.values_mut() {
            for input in sheet_cells.values_mut() {
                if !is_formula_input(input) {
                    continue;
                }
                let Some(formula) = input.as_str() else {
                    continue;
                };
                let rewritten = formula_model::rewrite_sheet_names_in_formula(
                    formula,
                    &old_display,
                    &new_display,
                );
                if rewritten != formula {
                    *input = JsonValue::String(rewritten);
                }
            }
        }

        true
    }

    fn parse_address(address: &str) -> Result<CellRef, JsValue> {
        CellRef::from_a1(address).map_err(|_| js_err(format!("invalid cell address: {address}")))
    }

    fn parse_range(range: &str) -> Result<Range, JsValue> {
        Range::from_a1(range).map_err(|_| js_err(format!("invalid range: {range}")))
    }

    fn get_pivot_schema_internal(
        &self,
        sheet: &str,
        source_range_a1: &str,
        sample_size: usize,
    ) -> Result<pivot_engine::PivotSchema, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let range = Self::parse_range(source_range_a1)?;
        let cache = self
            .engine
            .pivot_cache_from_range(&sheet, range)
            .map_err(|err| js_err(err.to_string()))?;
        Ok(cache.schema(sample_size))
    }

    fn calculate_pivot_writes_internal(
        &self,
        sheet: &str,
        source_range_a1: &str,
        destination_top_left_a1: &str,
        config: &pivot_engine::PivotConfig,
    ) -> Result<Vec<PivotCellWrite>, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let range = Self::parse_range(source_range_a1)?;
        let destination = Self::parse_address(destination_top_left_a1)?;

        let result = self
            .engine
            .calculate_pivot_from_range(&sheet, range, config)
            .map_err(|err| js_err(err.to_string()))?;

        let writes = result.to_cell_writes_with_formats(
            pivot_engine::CellRef {
            row: destination.row,
            col: destination.col,
            },
            config,
            &pivot_engine::PivotApplyOptions::default(),
        );

        let date_system = self.engine.date_system();
        let mut out = Vec::with_capacity(writes.len());
        for write in writes {
            out.push(PivotCellWrite {
                sheet: sheet.clone(),
                address: CellRef::new(write.row, write.col).to_a1(),
                value: pivot_value_to_json(write.value, date_system),
                number_format: write.number_format,
            });
        }
        Ok(out)
    }
    fn set_cell_style_id_internal(
        &mut self,
        sheet: &str,
        address: &str,
        style_id: u32,
    ) -> Result<(), JsValue> {
        self.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet);
            let cell_ref = Self::parse_address(address)?;
            let address = cell_ref.to_a1();
            this.engine
                .set_cell_style_id(&sheet, &address, style_id)
                .map_err(|err| js_err(err.to_string()))
        })
    }

    fn get_cell_style_id_internal(&self, sheet: &str, address: &str) -> Result<u32, JsValue> {
        let sheet = self.require_sheet(sheet)?;
        let cell_ref = Self::parse_address(address)?;
        let address = cell_ref.to_a1();
        let style_id = self
            .engine
            .get_cell_style_id(sheet, &address)
            .map_err(|err| js_err(err.to_string()))?;
        Ok(style_id.unwrap_or(0))
    }
    fn set_cell_internal(
        &mut self,
        sheet: &str,
        address: &str,
        input: JsonValue,
    ) -> Result<(), JsValue> {
        self.with_manual_calc_mode(|this| {
            if !is_scalar_json(&input) {
                return Err(js_err(format!("invalid cell value: {address}")));
            }

            let sheet = this.ensure_sheet(sheet);
            let cell_ref = Self::parse_address(address)?;
            let address = cell_ref.to_a1();

            // Legacy scalar edits overwrite any previous rich input for this cell.
            if let Some(rich_cells) = this.sheets_rich.get_mut(&sheet) {
                rich_cells.remove(&address);
            }

            if let Some((origin, end)) = this.engine.spill_range(&sheet, &address) {
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
                        this.pending_spill_clears
                            .insert(FormulaCellKey::new(sheet.clone(), CellRef::new(row, col)));
                    }
                }
            }

            let sheet_cells = this
                .sheets
                .get_mut(&sheet)
                .expect("sheet just ensured must exist");

            // `null` represents an empty cell in the JS protocol. Preserve sparse semantics in the
            // JSON input map by removing the stored entry instead of storing an explicit blank.
            //
            // In the engine, treat this as "clear contents" (value/formula -> blank) so formatting can
            // be preserved when a cell has a non-default style.
            if input.is_null() {
                this.engine
                    .set_cell_value(&sheet, &address, EngineValue::Blank)
                    .map_err(|err| js_err(err.to_string()))?;

                sheet_cells.remove(&address);
                // If this cell was previously tracked as part of a spill-clear batch, drop it so we
                // don't report direct input edits as recalc changes.
                this.pending_spill_clears
                    .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
                this.pending_formula_baselines
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
                    this.engine
                        .set_cell_value(&sheet, &address, EngineValue::Blank)
                        .map_err(|err| js_err(err.to_string()))?;
                    sheet_cells.remove(&address);
                    this.pending_spill_clears
                        .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
                    this.pending_formula_baselines
                        .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
                    return Ok(());
                }

                let canonical = if this.formula_locale.id == EN_US.id {
                    normalized
                } else {
                    canonicalize_formula_with_style(
                        &normalized,
                        this.formula_locale,
                        formula_engine::ReferenceStyle::A1,
                    )
                    .map_err(|err| js_err(err.to_string()))?
                };

                let key = FormulaCellKey::new(sheet.clone(), cell_ref);
                this.pending_formula_baselines
                    .entry(key)
                    .or_insert_with(|| {
                        engine_value_to_json(this.engine.get_cell_value(&sheet, &address))
                    });

                // Reset the stored value to blank so `getCell` returns null until the next recalc,
                // matching the existing worker semantics.
                this.engine
                    .set_cell_value(&sheet, &address, EngineValue::Blank)
                    .map_err(|err| js_err(err.to_string()))?;
                this.engine
                    .set_cell_formula(&sheet, &address, &canonical)
                    .map_err(|err| js_err(err.to_string()))?;

                sheet_cells.insert(address.clone(), JsonValue::String(canonical));
                return Ok(());
            }

            // Non-formula scalar value.
            this.engine
                .set_cell_value(&sheet, &address, json_to_engine_value(&input))
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.insert(address.clone(), input);
            // If this cell was previously tracked as part of a spill-clear batch (e.g. a multi-cell
            // paste over a spill range), drop it so we don't report direct input edits as recalc
            // changes.
            this.pending_spill_clears
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            this.pending_formula_baselines
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            Ok(())
        })
    }

    fn set_cell_rich_internal(
        &mut self,
        sheet: &str,
        address: &str,
        input: CellValue,
    ) -> Result<(), JsValue> {
        self.with_manual_calc_mode(|this| {
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
                this.set_cell_internal(sheet, address, scalar_input)?;

                // Preserve the typed representation for `getCellRich.input`.
                //
                // Note: For rich text values, the engine currently only stores the plain string value.
                // Persisting the input here allows callers to round-trip rich text styling even though
                // `getCellRich.value` will still reflect the scalar engine value.
                if !input.is_empty() {
                    let sheet = this.ensure_sheet(sheet);
                    let address = Self::parse_address(address)?.to_a1();
                    this.sheets_rich
                        .entry(sheet)
                        .or_default()
                        .insert(address, input);
                }

                return Ok(());
            }

            let sheet = this.ensure_sheet(sheet);
            let cell_ref = Self::parse_address(address)?;
            let address = cell_ref.to_a1();

            if let Some((origin, end)) = this.engine.spill_range(&sheet, &address) {
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
                        this.pending_spill_clears
                            .insert(FormulaCellKey::new(sheet.clone(), CellRef::new(row, col)));
                    }
                }
            }

            let sheet_cells = this
                .sheets
                .get_mut(&sheet)
                .expect("sheet just ensured must exist");
            let sheet_cells_rich = this
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
            this.engine
                .set_cell_value(&sheet, &address, engine_value)
                .map_err(|err| js_err(err.to_string()))?;

            // Rich values are not representable in the scalar workbook input schema; preserve scalar
            // compatibility by removing any stored scalar input for this cell.
            sheet_cells.remove(&address);

            // Store the full rich input for `getCellRich.input`.
            sheet_cells_rich.insert(address.clone(), input);

            this.pending_spill_clears
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            this.pending_formula_baselines
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            Ok(())
        })
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

    fn goal_seek_internal(
        &mut self,
        sheet: &str,
        target_cell: &str,
        target_value: f64,
        changing_cell: &str,
        tuning: GoalSeekTuning,
    ) -> Result<(GoalSeekResult, Vec<CellChange>), JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let target_cell_ref = Self::parse_address(target_cell)?;
        let changing_cell_ref = Self::parse_address(changing_cell)?;
        let target_cell = target_cell_ref.to_a1();
        let changing_cell = changing_cell_ref.to_a1();

        let mut params =
            GoalSeekParams::new(target_cell.as_str(), target_value, changing_cell.as_str());
        if let Some(max_iterations) = tuning.max_iterations {
            params.max_iterations = max_iterations;
        }
        if let Some(tolerance) = tuning.tolerance {
            params.tolerance = tolerance;
        }
        if tuning.derivative_step.is_some() {
            params.derivative_step = tuning.derivative_step;
        }
        if let Some(min_derivative) = tuning.min_derivative {
            params.min_derivative = min_derivative;
        }
        if let Some(max_bracket_expansions) = tuning.max_bracket_expansions {
            params.max_bracket_expansions = max_bracket_expansions;
        }

        let mut model = WorkbookGoalSeekModel::new(self, sheet.clone());
        let result = GoalSeek::solve(&mut model, params).map_err(|err| {
            let message = match err {
                WhatIfError::Model(err) => err.to_string(),
                WhatIfError::NonNumericCell { cell, value } => {
                    let value_desc = match value {
                        WhatIfCellValue::Number(n) => n.to_string(),
                        WhatIfCellValue::Text(s) => s,
                        WhatIfCellValue::Bool(b) => b.to_string(),
                        WhatIfCellValue::Blank => "blank".to_string(),
                    };
                    format!("cell {sheet}!{cell} is not numeric: {value_desc}")
                }
                WhatIfError::InvalidParams(msg) => format!("invalid goal seek parameters: {msg}"),
                WhatIfError::NoBracketFound => "goal seek: could not bracket a solution".to_string(),
                WhatIfError::NumericalFailure(msg) => format!("goal seek numerical failure: {msg}"),
            };
            js_err(message)
        })?;

        // Ensure the final workbook state matches the returned solution. Some `GoalSeek` exit paths
        // (notably `NoBracketFound`) can leave the changing cell at the last attempted value rather
        // than the returned `result.solution`.
        match model.wb.engine.get_cell_value(&sheet, &changing_cell) {
            EngineValue::Number(n) if n == result.solution => {}
            _ => {
                let json_solution = serde_json::Number::from_f64(result.solution)
                    .map(JsonValue::Number)
                    .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string()));
                model
                    .wb
                    .set_cell_internal(&sheet, &changing_cell, json_solution)?;
                model.recalculate().map_err(|err| js_err(err.to_string()))?;
            }
        }

        // Extract accumulated changes and add an explicit delta for the changing cell's final
        // value (since callers did not invoke `setCell` directly).
        let mut by_cell = std::mem::take(&mut model.changes);
        drop(model);

        by_cell.insert(
            FormulaCellKey::new(sheet.clone(), changing_cell_ref),
            engine_value_to_json(self.engine.get_cell_value(&sheet, &changing_cell)),
        );

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

        Ok((result, changes))
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
        let previous = self.engine.calc_settings().clone();
        if previous.calculation_mode != CalculationMode::Manual {
            let mut manual = previous.clone();
            manual.calculation_mode = CalculationMode::Manual;
            self.engine.set_calc_settings(manual);
        }

        let out = (|| {
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
                            sheet_cells
                                .insert(address.clone(), JsonValue::String(formula.to_string()));
                        } else {
                            let Some(value) =
                                engine_value_to_scalar_json_input(after.value.clone())
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

                    let phonetic = self
                        .engine
                        .get_cell_phonetic(&sheet, &address)
                        .map(|s| s.to_string());

                    // Reset stored value to blank while preserving the formula. This matches the
                    // `setCell` behavior where formula results are treated as unknown until recalc.
                    self.engine
                        .set_cell_value(&sheet, &address, EngineValue::Blank)
                        .map_err(|err| js_err(err.to_string()))?;
                    self.engine
                        .set_cell_formula(&sheet, &address, formula)
                        .map_err(|err| js_err(err.to_string()))?;
                    if let Some(phonetic) = phonetic {
                        // `Engine::set_cell_value`/`set_cell_formula` clear phonetic metadata, but
                        // structural edits should preserve it.
                        self.engine
                            .set_cell_phonetic(&sheet, &address, Some(phonetic))
                            .map_err(|err| js_err(err.to_string()))?;
                    }
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
        })();

        self.engine.set_calc_settings(previous);
        out
    }

    fn set_locale_id(&mut self, locale_id: &str) -> bool {
        let Some(formula_locale) = get_locale(locale_id) else {
            return false;
        };
        let Some(value_locale) = ValueLocaleConfig::for_locale_id(locale_id) else {
            return false;
        };
        let text_codepage = text_codepage_for_locale_id(locale_id);

        let previous = self.engine.calc_settings().clone();
        if previous.calculation_mode != CalculationMode::Manual {
            let mut manual = previous.clone();
            manual.calculation_mode = CalculationMode::Manual;
            self.engine.set_calc_settings(manual);
        }
        self.formula_locale = formula_locale;
        self.engine.set_locale_config(formula_locale.config.clone());
        self.engine.set_value_locale(value_locale);
        self.engine.set_text_codepage(text_codepage);
        self.engine.set_calc_settings(previous);
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

fn engine_value_to_js_scalar(value: EngineValue) -> JsValue {
    match value {
        EngineValue::Blank => JsValue::NULL,
        EngineValue::Bool(b) => JsValue::from_bool(b),
        EngineValue::Text(s) => JsValue::from_str(&s),
        EngineValue::Number(n) => {
            if n.is_finite() {
                JsValue::from_f64(n)
            } else {
                // Preserve the existing scalar protocol semantics used by `engine_value_to_json`:
                // JSON cannot represent NaN/Infinity, so degrade to a #NUM! error code.
                JsValue::from_str(ErrorKind::Num.as_code())
            }
        }
        EngineValue::Entity(entity) => JsValue::from_str(&entity.display),
        EngineValue::Record(record) => JsValue::from_str(&record.display),
        EngineValue::Error(kind) => JsValue::from_str(kind.as_code()),
        // Arrays should generally be spilled into grid cells. If one reaches the JS boundary,
        // degrade to its top-left value so callers still get a scalar.
        EngineValue::Array(arr) => engine_value_to_js_scalar(arr.top_left()),
        // The JS worker protocol only supports scalar-ish values today.
        //
        // Degrade any rich/non-scalar value (references, lambdas, entities, records, etc.) to its
        // display string so existing `getCell` / `recalculate` callers keep receiving scalars.
        other => JsValue::from_str(&other.to_string()),
    }
}

fn push_a1_col_name(col: u32, out: &mut String) {
    // Excel columns are 1-based in A1 notation. We store 0-based internally.
    let mut n = u64::from(col) + 1;
    // A u32 column index fits in at most 7 A1 letters (26^7 > u32::MAX).
    let mut buf = [0u8; 8];
    let mut len = 0usize;
    while n > 0 {
        let rem = (n - 1) % 26;
        buf[len] = b'A' + rem as u8;
        len += 1;
        n = (n - 1) / 26;
    }
    for b in buf[..len].iter().rev() {
        out.push(*b as char);
    }
}

fn push_u64_decimal(mut n: u64, out: &mut String) {
    let mut buf = [0u8; 20];
    let mut len = 0usize;
    loop {
        let digit = (n % 10) as u8;
        buf[len] = b'0' + digit;
        len += 1;
        n /= 10;
        if n == 0 {
            break;
        }
    }
    for b in buf[..len].iter().rev() {
        out.push(*b as char);
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

fn find_workbook_prefix_end(src: &str, start: usize) -> Option<usize> {
    // External workbook prefixes escape literal `]` characters by doubling them: `]]` -> `]`.
    //
    // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
    let bytes = src.as_bytes();
    if bytes.get(start) != Some(&b'[') {
        return None;
    }

    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == b']' {
            if bytes.get(i + 1) == Some(&b']') {
                i += 2;
                continue;
            }
            return Some(i + 1);
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside a multi-byte sequence as actual bracket characters.
        let ch = src[i..].chars().next()?;
        i += ch.len_utf8();
    }

    None
}

fn skip_ws(src: &str, mut i: usize) -> usize {
    while i < src.len() {
        let Some(ch) = src[i..].chars().next() else {
            break;
        };
        if !ch.is_whitespace() {
            break;
        }
        i += ch.len_utf8();
    }
    i
}

fn scan_quoted_sheet_name(src: &str, start: usize) -> Option<usize> {
    // Quoted sheet names escape apostrophes by doubling them: `''` -> `'`.
    let bytes = src.as_bytes();
    if bytes.get(start) != Some(&b'\'') {
        return None;
    }

    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == b'\'' {
            if bytes.get(i + 1) == Some(&b'\'') {
                i += 2;
                continue;
            }
            return Some(i + 1);
        }
        let ch = src[i..].chars().next()?;
        i += ch.len_utf8();
    }
    None
}

fn scan_unquoted_name(src: &str, start: usize) -> Option<usize> {
    // Match the engine's identifier lexer rules for unquoted sheet names / defined names.
    let first = src[start..].chars().next()?;
    if !is_ident_start_char(first) {
        return None;
    }
    let mut i = start + first.len_utf8();
    while i < src.len() {
        let ch = src[i..].chars().next()?;
        if is_ident_cont_char(ch) {
            i += ch.len_utf8();
            continue;
        }
        break;
    }
    Some(i)
}

fn scan_sheet_name_token(src: &str, start: usize) -> Option<usize> {
    let i = skip_ws(src, start);
    if i >= src.len() {
        return None;
    }
    match src[i..].chars().next()? {
        '\'' => scan_quoted_sheet_name(src, i),
        _ => scan_unquoted_name(src, i),
    }
}

fn find_workbook_prefix_end_if_valid(src: &str, start: usize) -> Option<usize> {
    let end = find_workbook_prefix_end(src, start)?;

    // Heuristic: only treat this as an external workbook prefix if it is immediately followed by:
    // - a sheet spec and `!` (e.g. `[Book.xlsx]Sheet1!A1`), OR
    // - a defined name identifier (e.g. `[Book.xlsx]MyName`).
    //
    // This avoids incorrectly treating nested structured references (which *are* nested) as
    // workbook prefixes while still supporting workbook names that contain `[` characters (Excel
    // treats `[` as plain text within workbook ids).
    let i = skip_ws(src, end);
    if let Some(mut sheet_end) = scan_sheet_name_token(src, i) {
        sheet_end = skip_ws(src, sheet_end);

        // `[Book.xlsx]Sheet1:Sheet3!A1` (external 3D span)
        if sheet_end < src.len() && src[sheet_end..].starts_with(':') {
            sheet_end += 1;
            sheet_end = skip_ws(src, sheet_end);
            sheet_end = scan_sheet_name_token(src, sheet_end)?;
            sheet_end = skip_ws(src, sheet_end);
        }

        if sheet_end < src.len() && src[sheet_end..].starts_with('!') {
            return Some(end);
        }
    }

    // Workbook-scoped external defined name `[Book.xlsx]MyName`.
    // Note: defined names are not quoted with `'` in formula text, so we only scan the unquoted
    // identifier form here.
    let name_start = skip_ws(src, end);
    if scan_unquoted_name(src, name_start).is_some() {
        return Some(end);
    }

    None
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
                        // Workbook prefixes are *not* nesting, even if the workbook name contains `[` characters
                        // (e.g. `=[A1[Name.xlsx]Sheet1!A1`). Prefer a non-nesting scan when the bracket segment
                        // is followed by a sheet name and `!` or a defined name.
                        if let Some(end) = find_workbook_prefix_end_if_valid(formula_prefix, i) {
                            i = end;
                            continue;
                        }

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

    let (opts, locale) = parse_options_and_locale_from_js(opts)?;

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
            name: normalize_function_context_name(&ctx.name, locale),
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

    /// Set the workbook's locale id.
    ///
    /// This updates both the formula locale (tokenization/localization) and value locale (parsing
    /// numbers/dates). For DBCS locales such as `ja-JP`, this may also update the workbook's legacy
    /// text codepage so `LENB`/`LEFTB`/... and `ASC`/`DBCS` behave like Excel.
    #[wasm_bindgen(js_name = "setLocale")]
    pub fn set_locale(&mut self, locale_id: String) -> bool {
        self.inner.set_locale_id(&locale_id)
    }

    /// Get the workbook's legacy Windows text codepage (used for DBCS `*B` text functions like
    /// `LENB`, `ASC`, and `DBCS`).
    #[wasm_bindgen(js_name = "getTextCodepage")]
    pub fn get_text_codepage(&self) -> u16 {
        self.inner.engine.text_codepage()
    }

    /// Set the workbook's legacy Windows text codepage (used for DBCS `*B` text functions like
    /// `LENB`, `ASC`, and `DBCS`).
    ///
    /// This marks compiled formulas dirty; callers should invoke `recalculate()` to observe changes.
    #[wasm_bindgen(js_name = "setTextCodepage")]
    pub fn set_text_codepage(&mut self, codepage: u16) -> Result<(), JsValue> {
        self.inner.with_manual_calc_mode(|this| {
            this.engine.set_text_codepage(codepage);
            Ok(())
        })
    }

    /// Intern (deduplicate) a style object into the workbook's style table, returning its style id.
    ///
    /// The input uses a JS-friendly shape (best-effort). Unknown keys are ignored.
    ///
    /// Style id `0` is always the default (empty) style.
    #[wasm_bindgen(js_name = "internStyle")]
    pub fn intern_style(&mut self, style: JsValue) -> Result<u32, JsValue> {
        let style = parse_style_from_js(style)?;
        Ok(self.inner.engine.intern_style(style))
    }

    /// Set (or clear) the default style id for all cells in a row.
    ///
    /// `row` is 0-indexed (engine coordinates). `style_id=0` (or `null`/`undefined`) clears the override.
    #[wasm_bindgen(js_name = "setRowStyleId")]
    pub fn set_row_style_id(&mut self, sheet: String, row: u32, style_id: Option<u32>) {
        // Backward compatibility: treat `0` as "clear override" (style id `0` is always the
        // default style in the workbook style table).
        let style_id = style_id.filter(|id| *id != 0);
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        let _ = self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(&sheet);
            this.engine.set_row_style_id(&sheet, row, style_id);
            Ok(())
        });
    }

    /// Set (or clear) the default style id for all cells in a column.
    ///
    /// `col` is 0-indexed (engine coordinates). `style_id=0` (or `null`/`undefined`) clears the override.
    #[wasm_bindgen(js_name = "setColStyleId")]
    pub fn set_col_style_id(&mut self, sheet: String, col: u32, style_id: Option<u32>) {
        // Backward compatibility: treat `0` as "clear override" (style id `0` is always the
        // default style in the workbook style table).
        let style_id = style_id.filter(|id| *id != 0);
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        let _ = self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(&sheet);
            this.engine.set_col_style_id(&sheet, col, style_id);
            Ok(())
        });
    }

    /// Replace the compressed format-run layer for a column (DocumentController `formatRunsByCol`).
    ///
    /// `runs` must be an array of `{ startRow, endRowExclusive, styleId }` objects.
    ///
    /// Row indices are 0-based and runs use half-open intervals `[startRow, endRowExclusive)`.
    #[wasm_bindgen(js_name = "setFormatRunsByCol")]
    pub fn set_format_runs_by_col(
        &mut self,
        sheet: String,
        col: u32,
        runs: JsValue,
    ) -> Result<(), JsValue> {
        let sheet = self.inner.ensure_sheet(&sheet);
        if col >= EXCEL_MAX_COLS {
            return Err(js_err(format!("col out of Excel bounds: {col}")));
        }

        fn parse_u32_field(obj: &Object, key: &str, context: &str) -> Result<u32, JsValue> {
            let value = Reflect::get(obj, &JsValue::from_str(key))
                .map_err(|_| js_err(format!("{context}: failed to read {key}")))?;
            let n = value
                .as_f64()
                .ok_or_else(|| js_err(format!("{context} must be a non-negative integer")))?;
            if !n.is_finite() || n < 0.0 || n > u32::MAX as f64 || n.fract() != 0.0 {
                return Err(js_err(format!("{context} must be a non-negative integer")));
            }
            Ok(n as u32)
        }

        let mut parsed: Vec<EngineFormatRun> = Vec::new();

        // Accept null/undefined as clearing the column's runs.
        if !(runs.is_null() || runs.is_undefined()) {
            let arr = runs
                .dyn_into::<Array>()
                .map_err(|_| js_err("setFormatRunsByCol: runs must be an array"))?;

            for (idx, item) in arr.iter().enumerate() {
                let obj = item.dyn_into::<Object>().map_err(|_| {
                    js_err(format!(
                        "setFormatRunsByCol: runs[{idx}] must be an object"
                    ))
                })?;

                let start_row = parse_u32_field(
                    &obj,
                    "startRow",
                    &format!("setFormatRunsByCol: runs[{idx}].startRow"),
                )?;
                let end_row_exclusive = parse_u32_field(
                    &obj,
                    "endRowExclusive",
                    &format!("setFormatRunsByCol: runs[{idx}].endRowExclusive"),
                )?;
                let style_id = parse_u32_field(
                    &obj,
                    "styleId",
                    &format!("setFormatRunsByCol: runs[{idx}].styleId"),
                )?;

                if start_row >= EXCEL_MAX_ROWS {
                    return Err(js_err(format!(
                        "setFormatRunsByCol: runs[{idx}].startRow out of Excel bounds: {start_row}"
                    )));
                }
                if end_row_exclusive > EXCEL_MAX_ROWS {
                    return Err(js_err(format!(
                        "setFormatRunsByCol: runs[{idx}].endRowExclusive out of Excel bounds: {end_row_exclusive}"
                    )));
                }
                if end_row_exclusive <= start_row {
                    return Err(js_err(format!(
                        "setFormatRunsByCol: runs[{idx}].endRowExclusive must be greater than startRow"
                    )));
                }

                parsed.push(EngineFormatRun {
                    start_row,
                    end_row_exclusive,
                    style_id,
                });
            }
        }

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            this.engine
                .set_format_runs_by_col(&sheet, col, parsed)
                .map_err(|err| js_err(err.to_string()))
        })
    }

    /// Set (or clear) the default style id for an entire sheet.
    ///
    /// Pass `null`/`undefined` (or `0` for backward compatibility) to clear the override.
    #[wasm_bindgen(js_name = "setSheetDefaultStyleId")]
    pub fn set_sheet_default_style_id(&mut self, sheet: String, style_id: Option<u32>) {
        // Backward compatibility: treat `0` as "clear override" (style id `0` is always the
        // default style in the workbook style table).
        let style_id = style_id.filter(|id| *id != 0);
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        let _ = self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(&sheet);
            this.engine.set_sheet_default_style_id(&sheet, style_id);
            Ok(())
        });
    }

    /// Replace the range-run formatting runs for a column.
    ///
    /// `runs` is an array of `{ startRow, endRowExclusive, styleId }`.
    #[wasm_bindgen(js_name = "setColFormatRuns")]
    pub fn set_col_format_runs(
        &mut self,
        sheet: String,
        col: u32,
        runs: JsValue,
    ) -> Result<(), JsValue> {
        let parsed: Vec<FormatRunDto> = if runs.is_null() || runs.is_undefined() {
            Vec::new()
        } else {
            serde_wasm_bindgen::from_value(runs).map_err(|err| js_err(err.to_string()))?
        };
        let runs: Vec<EngineFormatRun> = parsed
            .into_iter()
            .map(|r| EngineFormatRun {
                start_row: r.start_row,
                end_row_exclusive: r.end_row_exclusive,
                style_id: r.style_id,
            })
            .collect();
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner
            .with_manual_calc_mode(|this| this.set_col_format_runs_internal(&sheet, col, runs))
    }
    #[wasm_bindgen(js_name = "getCalcSettings")]
    pub fn get_calc_settings(&self) -> Result<JsValue, JsValue> {
        let settings = self.inner.engine.calc_settings();
        let dto = CalcSettingsDto {
            calculation_mode: settings.calculation_mode.into(),
            calculate_before_save: settings.calculate_before_save,
            full_precision: settings.full_precision,
            full_calc_on_load: settings.full_calc_on_load,
            iterative: IterativeCalcSettingsDto {
                enabled: settings.iterative.enabled,
                max_iterations: settings.iterative.max_iterations,
                max_change: settings.iterative.max_change,
            },
        };

        use serde::ser::Serialize as _;
        dto.serialize(&serde_wasm_bindgen::Serializer::json_compatible())
            .map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "setCalcSettings")]
    pub fn set_calc_settings(&mut self, settings: JsValue) -> Result<(), JsValue> {
        if settings.is_null() || settings.is_undefined() {
            return Err(js_err("settings must be an object"));
        }

        let dto: CalcSettingsInputDto = serde_wasm_bindgen::from_value(settings)
            .map_err(|err| js_err(format!("invalid calc settings: {err}")))?;

        if !dto.iterative.max_iterations.is_finite()
            || dto.iterative.max_iterations < 0.0
            || dto.iterative.max_iterations > u32::MAX as f64
            || dto.iterative.max_iterations.fract() != 0.0
        {
            return Err(js_err(
                "iterative.maxIterations must be a non-negative integer",
            ));
        }
        let max_iterations = dto.iterative.max_iterations as u32;
        if !dto.iterative.max_change.is_finite() || dto.iterative.max_change < 0.0 {
            return Err(js_err(
                "iterative.maxChange must be a finite number greater than or equal to 0",
            ));
        }

        self.inner.engine.set_calc_settings(CalcSettings {
            calculation_mode: dto.calculation_mode.into(),
            calculate_before_save: dto.calculate_before_save,
            iterative: IterativeCalculationSettings {
                enabled: dto.iterative.enabled,
                max_iterations,
                max_change: dto.iterative.max_change,
            },
            full_precision: dto.full_precision,
            full_calc_on_load: dto.full_calc_on_load,
        });
        Ok(())
    }

    #[wasm_bindgen(js_name = "setEngineInfo")]
    pub fn set_engine_info(&mut self, info: JsValue) -> Result<(), JsValue> {
        if info.is_null() || info.is_undefined() {
            return Err(js_err("setEngineInfo: info must be an object"));
        }

        let obj = info
            .dyn_into::<Object>()
            .map_err(|_| js_err("setEngineInfo: info must be an object"))?;

        let mut next: EngineInfo = self.inner.engine.engine_info().clone();

        fn update_string(obj: &Object, key: &str, out: &mut Option<String>) -> Result<(), JsValue> {
            let has = Reflect::has(obj, &JsValue::from_str(key)).unwrap_or(false);
            if !has {
                return Ok(());
            }
            let raw = Reflect::get(obj, &JsValue::from_str(key))
                .map_err(|_| js_err(format!("setEngineInfo: failed to read {key}")))?;
            if raw.is_null() || raw.is_undefined() {
                *out = None;
                return Ok(());
            }
            let s = raw
                .as_string()
                .ok_or_else(|| js_err(format!("setEngineInfo: {key} must be a string")))?;
            let trimmed = s.trim();
            *out = (!trimmed.is_empty()).then_some(trimmed.to_string());
            Ok(())
        }

        fn update_number(obj: &Object, key: &str, out: &mut Option<f64>) -> Result<(), JsValue> {
            let has = Reflect::has(obj, &JsValue::from_str(key)).unwrap_or(false);
            if !has {
                return Ok(());
            }
            let raw = Reflect::get(obj, &JsValue::from_str(key))
                .map_err(|_| js_err(format!("setEngineInfo: failed to read {key}")))?;
            if raw.is_null() || raw.is_undefined() {
                *out = None;
                return Ok(());
            }
            let n = raw
                .as_f64()
                .ok_or_else(|| js_err(format!("setEngineInfo: {key} must be a finite number")))?;
            if !n.is_finite() {
                return Err(js_err(format!(
                    "setEngineInfo: {key} must be a finite number"
                )));
            }
            *out = Some(n);
            Ok(())
        }

        update_string(&obj, "system", &mut next.system)?;
        update_string(&obj, "directory", &mut next.directory)?;
        update_string(&obj, "osversion", &mut next.osversion)?;
        update_string(&obj, "release", &mut next.release)?;
        update_string(&obj, "version", &mut next.version)?;
        update_number(&obj, "memavail", &mut next.memavail)?;
        update_number(&obj, "totmem", &mut next.totmem)?;

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(next);
            Ok(())
        })
    }

    #[wasm_bindgen(js_name = "setInfoOrigin")]
    pub fn set_info_origin(&mut self, origin: Option<String>) -> Result<(), JsValue> {
        let origin = origin.and_then(|s| {
            let s = s.trim();
            (!s.is_empty()).then_some(s.to_string())
        });
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            this.engine.set_info_origin(origin);
            Ok(())
        })
    }

    #[wasm_bindgen(js_name = "setInfoOriginForSheet")]
    pub fn set_info_origin_for_sheet(
        &mut self,
        sheet_name: String,
        origin_a1: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet_name = sheet_name.trim();
        // For parity with `setSheetOrigin`, treat empty/whitespace sheet names as the default
        // worksheet.
        let sheet_name = if sheet_name.is_empty() {
            DEFAULT_SHEET
        } else {
            sheet_name
        };
        let origin_trimmed = origin_a1
            .as_deref()
            .map(str::trim)
            .filter(|s| !s.is_empty());

        // `INFO("origin")` is tied to host-provided worksheet view state. The core engine validates
        // and stores the origin as a parsed A1 cell reference, so we pass the host string through
        // verbatim (after trimming) and let the engine handle parsing.
        //
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet_name);
            this.engine
                .set_sheet_origin(&sheet, origin_trimmed)
                .map_err(|err| js_err(err.to_string()))
        })
    }

    #[wasm_bindgen(js_name = "setInfoSystem")]
    pub fn set_info_system(&mut self, system: Option<String>) {
        let mut info = self.inner.engine.engine_info().clone();
        info.system = system.and_then(|s| {
            let s = s.trim();
            (!s.is_empty()).then_some(s.to_string())
        });
        let _ = self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        });
    }

    #[wasm_bindgen(js_name = "setInfoOSVersion")]
    pub fn set_info_os_version(&mut self, os_version: Option<String>) {
        let mut info = self.inner.engine.engine_info().clone();
        info.osversion = os_version.and_then(|s| {
            let s = s.trim();
            (!s.is_empty()).then_some(s.to_string())
        });
        let _ = self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        });
    }

    #[wasm_bindgen(js_name = "setInfoRelease")]
    pub fn set_info_release(&mut self, release: Option<String>) {
        let mut info = self.inner.engine.engine_info().clone();
        info.release = release.and_then(|s| {
            let s = s.trim();
            (!s.is_empty()).then_some(s.to_string())
        });
        let _ = self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        });
    }

    #[wasm_bindgen(js_name = "setInfoVersion")]
    pub fn set_info_version(&mut self, version: Option<String>) {
        let mut info = self.inner.engine.engine_info().clone();
        info.version = version.and_then(|s| {
            let s = s.trim();
            (!s.is_empty()).then_some(s.to_string())
        });
        let _ = self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        });
    }

    #[wasm_bindgen(js_name = "setInfoMemAvail")]
    pub fn set_info_mem_avail(&mut self, mem_avail: Option<f64>) -> Result<(), JsValue> {
        if let Some(n) = mem_avail {
            if !n.is_finite() {
                return Err(js_err("memAvail must be a finite number"));
            }
        }
        let mut info = self.inner.engine.engine_info().clone();
        info.memavail = mem_avail;
        self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        })?;
        Ok(())
    }

    #[wasm_bindgen(js_name = "setInfoTotMem")]
    pub fn set_info_tot_mem(&mut self, tot_mem: Option<f64>) -> Result<(), JsValue> {
        if let Some(n) = tot_mem {
            if !n.is_finite() {
                return Err(js_err("totMem must be a finite number"));
            }
        }
        let mut info = self.inner.engine.engine_info().clone();
        info.totmem = tot_mem;
        self.inner.with_manual_calc_mode(|this| {
            this.engine.set_engine_info(info);
            Ok(())
        })?;
        Ok(())
    }

    #[wasm_bindgen(js_name = "fromJson")]
    pub fn from_json(json: &str) -> Result<WasmWorkbook, JsValue> {
        #[derive(Debug, Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct WorkbookJson {
            #[serde(default, rename = "localeId")]
            locale_id: Option<String>,
            #[serde(default, rename = "formulaLanguage")]
            formula_language: Option<WorkbookFormulaLanguageDto>,
            #[serde(default, rename = "sheetOrder")]
            sheet_order: Option<Vec<String>>,
            /// Optional workbook text codepage (Windows codepage number).
            ///
            /// This powers Excel's legacy DBCS (`*B`) text functions (e.g. `LENB`) which behave
            /// differently under Japanese codepages (e.g. 932 / Shift-JIS).
            #[serde(default, rename = "textCodepage", alias = "codepage", alias = "text_codepage")]
            text_codepage: Option<u16>,
            sheets: BTreeMap<String, SheetJson>,
            #[serde(default)]
            style_table: BTreeMap<u32, formula_engine::style_patch::StylePatch>,
        }

        #[derive(Debug, Deserialize)]
        #[serde(untagged)]
        enum SheetVisibilityJson {
            String(String),
            Other(JsonValue),
        }

        #[derive(Debug, Deserialize)]
        #[serde(untagged)]
        enum TabColorJson {
            Color(TabColor),
            String(String),
            Other(JsonValue),
        }

        #[derive(Debug, Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct SheetJson {
            #[serde(default)]
            row_count: Option<u32>,
            #[serde(default)]
            col_count: Option<u32>,
            #[serde(default)]
            visibility: Option<SheetVisibilityJson>,
            #[serde(default, rename = "tabColor")]
            tab_color: Option<TabColorJson>,
            #[serde(default, rename = "cellPhonetics")]
            cell_phonetics: Option<BTreeMap<String, JsonValue>>,
            cells: BTreeMap<String, JsonValue>,
            #[serde(default)]
            default_style_id: Option<u32>,
            #[serde(default)]
            row_style_ids: BTreeMap<u32, u32>,
            #[serde(default)]
            col_style_ids: BTreeMap<u32, u32>,
            #[serde(default)]
            format_runs_by_col: BTreeMap<u32, Vec<formula_engine::style_patch::FormatRun>>,
            #[serde(default)]
            cell_style_ids: BTreeMap<String, u32>,
        }

        let parsed: WorkbookJson = serde_json::from_str(json)
            .map_err(|err| js_err(format!("invalid workbook json: {err}")))?;
        let WorkbookJson {
            locale_id,
            formula_language,
            sheet_order,
            text_codepage,
            sheets,
            style_table,
        } = parsed;
        let formula_language = formula_language.unwrap_or(WorkbookFormulaLanguageDto::Localized);

        let mut wb = WorkbookState::new_empty();

        // Best-effort: set workbook locale before importing cells so localized formulas and
        // locale-aware parsing semantics (argument separators, decimal commas, boolean keywords)
        // are handled correctly during JSON hydration.
        //
        // Unknown locale ids are ignored for backwards compatibility (treat as en-US).
        //
        // Note: `toJson()` currently always emits canonical (en-US) formula strings even when
        // `localeId` is non-en-US. The `formulaLanguage` field disambiguates this: when it is set
        // to `"canonical"`, we delay applying the workbook locale until after formula import so the
        // canonical text is not reinterpreted using localized parsing rules.
        if formula_language != WorkbookFormulaLanguageDto::Canonical {
            if let Some(locale_id) = locale_id.as_deref() {
                wb.set_locale_id(locale_id);
            }
        }

        if let Some(codepage) = text_codepage {
            wb.engine.set_text_codepage(codepage);
        }

        // Import the style table up-front so per-layer style ids can be resolved by the engine.
        for (style_id, patch) in style_table {
            wb.engine.set_style_patch(style_id, patch);
        }

        // Create all sheets up-front so cross-sheet formula references resolve correctly.
        //
        // When `sheetOrder` is provided, preserve the tab ordering by creating sheets in that order
        // before importing cells. This is important for 3D references (e.g. `Sheet1:Sheet3!A1`) and
        // worksheet functions that consult sheet indices like `SHEET()`.
        let mut ensured: BTreeSet<String> = BTreeSet::new();
        if let Some(order) = sheet_order.as_ref() {
            for name in order {
                if ensured.contains(name) {
                    continue;
                }
                if !sheets.contains_key(name) {
                    continue;
                }
                wb.ensure_sheet(name);
                ensured.insert(name.clone());
            }
        }
        for sheet_name in sheets.keys() {
            if ensured.contains(sheet_name) {
                continue;
            }
            wb.ensure_sheet(sheet_name);
            ensured.insert(sheet_name.clone());
        }

        for (sheet_name, sheet) in sheets {
            let SheetJson {
                row_count,
                col_count,
                visibility,
                tab_color,
                cell_phonetics,
                cells,
                default_style_id,
                row_style_ids,
                col_style_ids,
                format_runs_by_col,
                cell_style_ids,
            } = sheet;
            let display_name = wb.ensure_sheet(&sheet_name);

            if let Some(raw) = visibility.as_ref().and_then(|v| match v {
                SheetVisibilityJson::String(s) => Some(s.as_str()),
                SheetVisibilityJson::Other(_other) => None,
            }) {
                let trimmed = raw.trim();
                if trimmed.is_empty() || trimmed == "visible" {
                    wb.sheet_visibility.remove(&display_name);
                } else {
                    let vis = match trimmed {
                        "hidden" => Some(SheetVisibility::Hidden),
                        "veryHidden" | "very_hidden" | "veryhidden" => Some(SheetVisibility::VeryHidden),
                        _ => None,
                    };
                    match vis {
                        Some(vis) => {
                            wb.sheet_visibility.insert(display_name.clone(), vis);
                        }
                        None => {
                            // Backwards/forwards compatible behavior: unknown visibility values are
                            // treated as the default ("visible") instead of failing the entire
                            // workbook hydration.
                            wb.sheet_visibility.remove(&display_name);
                        }
                    }
                }
            }

            if let Some(color) = tab_color {
                match color {
                    TabColorJson::String(raw) => {
                        let mut trimmed = raw.trim();
                        if let Some(stripped) = trimmed.strip_prefix('#') {
                            trimmed = stripped;
                        }
                        if trimmed.is_empty() {
                            wb.sheet_tab_colors.remove(&display_name);
                        } else {
                            wb.sheet_tab_colors.insert(
                                display_name.clone(),
                                TabColor::rgb(trimmed.to_uppercase()),
                            );
                        }
                    }
                    TabColorJson::Color(color) => {
                        let is_empty = color.rgb.is_none()
                            && color.theme.is_none()
                            && color.indexed.is_none()
                            && color.tint.is_none()
                            && color.auto.is_none();
                        if is_empty {
                            wb.sheet_tab_colors.remove(&display_name);
                        } else {
                            wb.sheet_tab_colors
                                .insert(display_name.clone(), color.clone());
                        }
                    }
                    TabColorJson::Other(_other) => {
                        // Ignore unknown tabColor types for forwards compatibility.
                    }
                }
            }
            // Apply sheet dimensions (when provided) before importing cells so large addresses
            // can be set without pre-populating the full grid.
            if row_count.is_some() || col_count.is_some() {
                let rows = row_count.unwrap_or(EXCEL_MAX_ROWS);
                let cols = col_count.unwrap_or(EXCEL_MAX_COLS);
                if rows != EXCEL_MAX_ROWS || cols != EXCEL_MAX_COLS {
                    wb.set_sheet_dimensions_internal(&display_name, rows, cols)?;
                }
            }

            // Apply layered formatting metadata.
            if let Some(style_id) = default_style_id.filter(|id| *id != 0) {
                wb.engine
                    .set_sheet_default_patch_style_id(&display_name, style_id);
            }
            for (row, style_id) in row_style_ids {
                if style_id != 0 {
                    wb.engine.set_row_patch_style_id(&display_name, row, style_id);
                }
            }
            for (col, style_id) in col_style_ids {
                if style_id != 0 {
                    wb.engine.set_col_patch_style_id(&display_name, col, style_id);
                }
            }
            for (col, runs) in format_runs_by_col {
                if !runs.is_empty() {
                    wb.engine
                        .set_patch_format_runs_by_col(&display_name, col, runs);
                }
            }
            for (addr, style_id) in cell_style_ids {
                if style_id != 0 {
                    wb.engine
                        .set_cell_patch_style_id(&display_name, &addr, style_id)
                        .map_err(|err| js_err(err.to_string()))?;
                }
            }

            for (address, input) in cells {
                if !is_scalar_json(&input) {
                    return Err(js_err(format!("invalid cell value: {address}")));
                }
                if input.is_null() {
                    // `null` cells are treated as absent (sparse semantics).
                    continue;
                }
                wb.set_cell_internal(&display_name, &address, input)?;
            }

            // Apply per-cell phonetic guide metadata (furigana) after cell inputs have been set.
            // Setting cell values/formulas clears phonetic metadata in the engine, so we must apply
            // it after importing `cells`.
            if let Some(phonetics) = cell_phonetics {
                // Best-effort: ignore invalid addresses or other errors so optional metadata
                // doesn't prevent opening the workbook.
                let _ = wb.with_manual_calc_mode(|this| {
                    for (address, value) in phonetics {
                        let Some(phonetic) = value.as_str() else {
                            continue;
                        };
                        let _ = this.engine.set_cell_phonetic(
                            &display_name,
                            &address,
                            Some(phonetic.to_string()),
                        );
                    }
                    Ok(())
                });
            }
        }

        // Ensure the workbook locale is applied for subsequent edits/value coercion.
        if formula_language == WorkbookFormulaLanguageDto::Canonical {
            if let Some(locale_id) = locale_id.as_deref() {
                wb.set_locale_id(locale_id);
            }
        }

        if wb.sheets.is_empty() {
            wb.ensure_sheet(DEFAULT_SHEET);
        }

        Ok(WasmWorkbook { inner: wb })
    }

    #[wasm_bindgen(js_name = "fromXlsxBytes")]
    pub fn from_xlsx_bytes(bytes: &[u8]) -> Result<WasmWorkbook, JsValue> {
        // Ensure the function registry is populated before parsing any workbook formulas.
        ensure_rust_constructors_run();

        if formula_office_crypto::is_encrypted_ooxml_ole(bytes) {
            return Err(js_err(
                "workbook is encrypted/password-protected; use `fromEncryptedXlsxBytes(bytes, password)`",
            ));
        }

        let model = formula_xlsx::read_workbook_model_from_bytes(bytes)
            .map_err(|err| js_err(err.to_string()))?;
        Self::from_workbook_model(model)
    }

    #[wasm_bindgen(js_name = "fromModelJson")]
    pub fn from_model_json(model_json: String) -> Result<WasmWorkbook, JsValue> {
        let model: formula_model::Workbook = serde_json::from_str(&model_json)
            .map_err(|err| js_err(format!("invalid workbook model json: {err}")))?;
        Self::from_workbook_model(model)
    }

    fn from_workbook_model(model: formula_model::Workbook) -> Result<WasmWorkbook, JsValue> {
        let mut wb = WorkbookState::new_empty();

        // Import workbook calculation settings before seeding any values/formulas so features like
        // "precision as displayed" (`full_precision = false`) can affect how cached values are
        // stored at load time.
        //
        // Note: The WASM worker protocol expects manual recalc (callers invoke `recalculate()`
        // explicitly), so preserve manual calculation mode regardless of what the XLSX requests.
        let mut calc_settings = model.calc_settings.clone();
        calc_settings.calculation_mode = formula_engine::calc_settings::CalculationMode::Manual;
        wb.engine.set_calc_settings(calc_settings);

        // Date system influences date serials for NOW/TODAY/DATE, etc.
        wb.engine.set_date_system(match model.date_system {
            DateSystem::Excel1900 => formula_engine::date::ExcelDateSystem::EXCEL_1900,
            DateSystem::Excel1904 => formula_engine::date::ExcelDateSystem::Excel1904,
        });

        // Import the workbook style table so style ids used by row/column formatting layers can be
        // resolved by worksheet information functions like `CELL("protect")`.
        wb.engine.set_style_table(model.styles.clone());
        // DBCS / byte-count text functions (LENB, etc) depend on the workbook codepage.
        wb.engine.set_text_codepage(model.codepage);

        // Create all sheets up-front so formulas can resolve cross-sheet references.
        for sheet in &model.sheets {
            let sheet_name = wb.ensure_sheet(&sheet.name);
            if sheet.visibility != SheetVisibility::Visible {
                wb.sheet_visibility.insert(sheet_name.clone(), sheet.visibility);
            }
            if let Some(color) = sheet.tab_color.as_ref() {
                let is_empty = color.rgb.is_none()
                    && color.theme.is_none()
                    && color.indexed.is_none()
                    && color.tint.is_none()
                    && color.auto.is_none();
                if !is_empty {
                    wb.sheet_tab_colors.insert(sheet_name, color.clone());
                }
            }
        }

        // Apply per-sheet dimensions (logical grid size) before importing cells/formulas so
        // whole-column/row semantics (`A:A`, `1:1`) resolve correctly for large sheets.
        for sheet in &model.sheets {
            if sheet.row_count != EXCEL_MAX_ROWS || sheet.col_count != EXCEL_MAX_COLS {
                wb.set_sheet_dimensions_internal(&sheet.name, sheet.row_count, sheet.col_count)?;
            }
        }

        // Map workbook model style ids through the engine's style interner so row/column/cell
        // formatting layers reference the engine's canonical style ids. This keeps style ids
        // consistent even when the incoming style table contains duplicate entries.
        let mut style_id_map: Vec<u32> = Vec::with_capacity(model.styles.styles.len());
        style_id_map.push(0);
        for style in model.styles.styles.iter().skip(1) {
            style_id_map.push(wb.engine.intern_style(style.clone()));
        }

        // Best-effort: import the persisted worksheet view origin (`pane/@topLeftCell`) so
        // `INFO("origin")` returns an Excel-like value immediately after XLSX import.
        //
        // Hosts may still override this later via `setSheetOrigin`.
        for sheet in &model.sheets {
            let Some(origin) = sheet.view.pane.top_left_cell else {
                continue;
            };
            let sheet_name = wb
                .resolve_sheet(&sheet.name)
                .expect("sheet just ensured must resolve")
                .to_string();
            let origin = origin.to_a1();
            let _ = wb.engine.set_sheet_origin(&sheet_name, Some(&origin));
        }

        // Import worksheet column/row properties (width/hidden/default style) and default column
        // width.
        //
        // These are persisted in OOXML (`col/@width`, `col/@hidden`, `col/@style`, `row/@s`, and
        // `<sheetFormatPr defaultColWidth="...">`) and are needed for workbook info functions like
        // `CELL("format")` and `CELL("width")`.
        for sheet in &model.sheets {
            let sheet_name = wb
                .resolve_sheet(&sheet.name)
                .expect("sheet just ensured must resolve")
                .to_string();

            wb.engine
                .set_sheet_default_col_width(&sheet_name, sheet.default_col_width);

            for (&col, props) in &sheet.col_properties {
                if col >= EXCEL_MAX_COLS {
                    continue;
                }
                if let Some(width) = props.width {
                    wb.set_col_width_chars_internal(&sheet_name, col, Some(width))?;
                }
                wb.engine.set_col_hidden(&sheet_name, col, props.hidden);
                let mapped_style_id = props
                    .style_id
                    .and_then(|id| style_id_map.get(id as usize).copied())
                    .filter(|id| *id != 0);
                wb.engine
                    .set_col_style_id(&sheet_name, col, mapped_style_id);
            }

            // Outline indices are 1-based (Excel / OOXML).
            //
            // Some workbooks persist hidden state as outline/group collapse flags rather than the
            // user-hidden bit (`col_properties[*].hidden`), so include outline-based hidden state as
            // well.
            for (index, entry) in sheet.outline.cols.iter() {
                if entry.hidden.is_hidden() && index > 0 {
                    let col = index - 1;
                    if col < EXCEL_MAX_COLS {
                        wb.engine.set_col_hidden(&sheet_name, col, true);
                    }
                }
            }

            for (&row, props) in &sheet.row_properties {
                let Some(style_id) = props.style_id else {
                    continue;
                };
                let mapped = style_id_map
                    .get(style_id as usize)
                    .copied()
                    .unwrap_or(0);
                if mapped != 0 {
                    wb.engine.set_row_style_id(&sheet_name, row, Some(mapped));
                }
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
                let phonetic = cell.phonetic.as_deref().map(|s| s.to_string());

                // Apply formatting metadata first (including style-only cells) so cached values can
                // be rounded correctly when the workbook uses "precision as displayed".
                if cell.style_id != 0 {
                    let mapped = style_id_map
                        .get(cell.style_id as usize)
                        .copied()
                        .unwrap_or(0);
                    if mapped != 0 {
                        wb.engine
                            .set_cell_style_id(&sheet_name, &address, mapped)
                            .map_err(|err| js_err(err.to_string()))?;
                    }
                }

                // Skip style-only cells (no value/formula/phonetic). These are not represented in
                // the sparse JS input map (`toJson()`), but their formatting must still be present
                // in the calc engine so worksheet information functions like `CELL("format")` /
                // `CELL("protect")` observe the correct metadata.
                let has_formula = cell.formula.is_some();
                let has_value = !cell.value.is_empty();
                let has_phonetic = cell.phonetic.is_some();
                if !has_formula && !has_value && !has_phonetic {
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
                        if let Some(phonetic) = &phonetic {
                            // `Engine::set_cell_formula` clears phonetic metadata, so re-apply it after
                            // setting the formula.
                            wb.engine
                                .set_cell_phonetic(&sheet_name, &address, Some(phonetic.clone()))
                                .map_err(|err| js_err(err.to_string()))?;
                        }

                        let sheet_cells = wb
                            .sheets
                            .get_mut(&sheet_name)
                            .expect("sheet just ensured must exist");
                        sheet_cells.insert(address.clone(), JsonValue::String(display));
                        continue;
                    }
                }

                if let Some(phonetic) = &phonetic {
                    wb.engine
                        .set_cell_phonetic(&sheet_name, &address, Some(phonetic.clone()))
                        .map_err(|err| js_err(err.to_string()))?;
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

    #[wasm_bindgen(js_name = "fromEncryptedXlsxBytes")]
    pub fn from_encrypted_xlsx_bytes(
        bytes: &[u8],
        password: String,
    ) -> Result<WasmWorkbook, JsValue> {
        // Ensure the function registry is populated before importing any workbook formulas.
        ensure_rust_constructors_run();

        if !formula_office_crypto::is_encrypted_ooxml_ole(bytes) {
            // Not an Office-encrypted OLE container; fall back to the plaintext XLSX/XLSM loader.
            return Self::from_xlsx_bytes(bytes);
        }

        let decrypted =
            formula_office_crypto::decrypt_encrypted_package_ole(bytes, &password).map_err(|err| match err {
                // Special-case errors that imply we decrypted successfully but didn't end up with a
                // workbook ZIP package.
                formula_office_crypto::OfficeCryptoError::InvalidFormat(message)
                    if message.contains("ZIP archive") =>
                {
                    js_err(
                        "decrypted payload is not an `.xlsx`/`.xlsm`/`.xlsb` ZIP package; only encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks are supported for now",
                    )
                }
                other => office_crypto_err(other),
            })?;

        // Office-encrypted containers can wrap arbitrary payloads (e.g. XLS, DOCX). We support
        // encrypted OOXML workbooks stored as ZIP packages:
        // - `.xlsx`/`.xlsm`: `xl/workbook.xml`
        // - `.xlsb`: `xl/workbook.bin`
        if decrypted.len() < 2 || &decrypted[..2] != b"PK" {
            return Err(js_err(
                "decrypted payload is not an `.xlsx`/`.xlsm`/`.xlsb` ZIP package; only encrypted `.xlsx`/`.xlsm`/`.xlsb` workbooks are supported for now",
            ));
        }

        let cursor = std::io::Cursor::new(decrypted.as_slice());
        let archive = zip::ZipArchive::new(cursor).map_err(|err| {
            js_err(format!(
                "decrypted payload is not a valid ZIP archive: {err}"
            ))
        })?;

        let mut has_workbook_xml = false;
        let mut has_workbook_bin = false;
        for name in archive.file_names() {
            // Normalize ZIP entry names to forward slashes (Excel uses `/`, but tolerate `\`).
            let mut normalized = name.trim_start_matches('/');
            let replaced;
            if normalized.contains('\\') {
                replaced = normalized.replace('\\', "/");
                normalized = &replaced;
            }
            if normalized.eq_ignore_ascii_case("xl/workbook.xml") {
                has_workbook_xml = true;
            } else if normalized.eq_ignore_ascii_case("xl/workbook.bin") {
                has_workbook_bin = true;
            }
            if has_workbook_xml || has_workbook_bin {
                break;
            }
        }

        if has_workbook_xml {
            return Self::from_xlsx_bytes(&decrypted);
        }
        if has_workbook_bin {
            let options = formula_xlsb::OpenOptions {
                preserve_unknown_parts: false,
                preserve_parsed_parts: false,
                preserve_worksheets: false,
                decode_formulas: true,
            };
            let wb = formula_xlsb::XlsbWorkbook::open_from_bytes_with_options(&decrypted, options)
                .map_err(|err| js_err(format!("invalid .xlsb workbook: {err}")))?;
            let model = xlsb_to_model_workbook(&wb)
                .map_err(|err| js_err(format!("invalid .xlsb workbook: {err}")))?;
            return Self::from_workbook_model(model);
        }
        Err(js_err(
            "decrypted payload is a ZIP file but does not appear to be an `.xlsx`/`.xlsm`/`.xlsb` workbook (missing `xl/workbook.xml` and `xl/workbook.bin`)",
        ))
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

    /// Rename a worksheet and rewrite formulas that reference it (Excel-like).
    ///
    /// Returns `false` when `old_name` does not exist or `new_name` conflicts with another sheet.
    #[wasm_bindgen(js_name = "renameSheet")]
    pub fn rename_sheet(&mut self, old_name: String, new_name: String) -> bool {
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner
            .with_manual_calc_mode(|this| Ok(this.rename_sheet_internal(&old_name, &new_name)))
            .unwrap_or(false)
    }

    #[wasm_bindgen(js_name = "setSheetDisplayName")]
    pub fn set_sheet_display_name(
        &mut self,
        sheet_key: String,
        display_name: String,
    ) -> Result<(), JsValue> {
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            this.set_sheet_display_name_internal(&sheet_key, &display_name)
        })
    }
    /// Set (or clear) a per-column width override for a sheet.
    ///
    /// `width` is expressed in Excel "character" units (OOXML `col/@width`), **not pixels**.
    ///
    /// Prefer [`WasmWorkbook::set_col_width_chars`] for an explicit unit name.
    ///
    /// Pass `null`/`undefined` to clear the override.
    #[wasm_bindgen(js_name = "setColWidth")]
    pub fn set_col_width(
        &mut self,
        sheet_name: String,
        col: u32,
        width: JsValue,
    ) -> Result<(), JsValue> {
        let width = if width.is_null() || width.is_undefined() {
            None
        } else {
            let raw = width
                .as_f64()
                .ok_or_else(|| js_err("width must be a number or null".to_string()))?;
            if !raw.is_finite() || raw < 0.0 {
                return Err(js_err(
                    "width must be a non-negative finite number".to_string(),
                ));
            }
            Some(raw as f32)
        };

        self.inner
            .set_col_width_chars_internal(&sheet_name, col, width)
    }

    /// Set (or clear) a per-column width override for a sheet.
    ///
    /// `width_chars` is expressed in Excel "character" units (OOXML `col/@width`), **not pixels**.
    ///
    /// Pass `null`/`undefined` to clear the override.
    #[wasm_bindgen(js_name = "setColWidthChars")]
    pub fn set_col_width_chars(
        &mut self,
        sheet_name: String,
        col: u32,
        width_chars: JsValue,
    ) -> Result<(), JsValue> {
        let width_chars = if width_chars.is_null() || width_chars.is_undefined() {
            None
        } else {
            let raw = width_chars
                .as_f64()
                .ok_or_else(|| js_err("widthChars must be a number or null".to_string()))?;
            if !raw.is_finite() || raw < 0.0 {
                return Err(js_err(
                    "widthChars must be a non-negative finite number".to_string(),
                ));
            }
            Some(raw as f32)
        };

        self.inner
            .set_col_width_chars_internal(&sheet_name, col, width_chars)
    }

    /// Set whether a column is user-hidden.
    ///
    /// `col` is 0-based (A=0).
    #[wasm_bindgen(js_name = "setColHidden")]
    pub fn set_col_hidden(
        &mut self,
        sheet_name: String,
        col: u32,
        hidden: bool,
    ) -> Result<(), JsValue> {
        if col >= EXCEL_MAX_COLS {
            return Err(js_err(format!("col out of Excel bounds: {col}")));
        }
        let sheet_name = sheet_name.trim();
        let sheet_name = if sheet_name.is_empty() {
            DEFAULT_SHEET
        } else {
            sheet_name
        };

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet_name);
            this.engine.set_col_hidden(&sheet, col, hidden);
            Ok(())
        })
    }

    /// Set (or clear) the sheet's default column width in Excel "character" units.
    ///
    /// This corresponds to the worksheet's OOXML `<sheetFormatPr defaultColWidth="...">` attribute.
    ///
    /// Pass `null`/`undefined` to clear the override back to Excel's standard default width.
    #[wasm_bindgen(js_name = "setSheetDefaultColWidth")]
    pub fn set_sheet_default_col_width(
        &mut self,
        sheet_name: String,
        width_chars: JsValue,
    ) -> Result<(), JsValue> {
        let width_chars = if width_chars.is_null() || width_chars.is_undefined() {
            None
        } else {
            let raw = width_chars
                .as_f64()
                .ok_or_else(|| js_err("widthChars must be a number or null".to_string()))?;
            if !raw.is_finite() || raw < 0.0 {
                return Err(js_err(
                    "widthChars must be a non-negative finite number".to_string(),
                ));
            }
            Some(raw as f32)
        };

        let sheet_name = sheet_name.trim();
        let sheet_name = if sheet_name.is_empty() {
            DEFAULT_SHEET
        } else {
            sheet_name
        };

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet_name);
            this.engine.set_sheet_default_col_width(&sheet, width_chars);
            Ok(())
        })
    }

    /// Update workbook file metadata used by Excel-compatible functions like `CELL("filename")`
    /// and `INFO("directory")`.
    #[wasm_bindgen(js_name = "setWorkbookFileMetadata")]
    pub fn set_workbook_file_metadata(
        &mut self,
        directory: JsValue,
        filename: JsValue,
    ) -> Result<(), JsValue> {
        let directory = if directory.is_null() || directory.is_undefined() {
            None
        } else {
            Some(
                directory
                    .as_string()
                    .ok_or_else(|| js_err("directory must be a string or null".to_string()))?,
            )
        };

        let filename = if filename.is_null() || filename.is_undefined() {
            None
        } else {
            Some(
                filename
                    .as_string()
                    .ok_or_else(|| js_err("filename must be a string or null".to_string()))?,
            )
        };

        self.inner
            .set_workbook_file_metadata_internal(directory.as_deref(), filename.as_deref())
    }

    /// Set the style id for a cell.
    ///
    /// Note: unlike `setCell`, this does not modify a cell's value/formula.
    #[wasm_bindgen(js_name = "setCellStyleId")]
    pub fn set_cell_style_id(
        &mut self,
        sheet: String,
        address: String,
        style_id: u32,
    ) -> Result<(), JsValue> {
        let sheet = sheet.trim();
        let sheet = if sheet.is_empty() { DEFAULT_SHEET } else { sheet };
        self.inner
            .set_cell_style_id_internal(sheet, &address, style_id)
    }

    #[wasm_bindgen(js_name = "setSheetOrigin")]
    pub fn set_sheet_origin(
        &mut self,
        sheet_name: String,
        origin: JsValue,
    ) -> Result<(), JsValue> {
        let sheet_name = sheet_name.trim();
        let sheet_name = if sheet_name.is_empty() {
            DEFAULT_SHEET
        } else {
            sheet_name
        };

        let origin_opt: Option<String> = if origin.is_null() || origin.is_undefined() {
            None
        } else {
            Some(
                origin
                    .as_string()
                    .ok_or_else(|| js_err("origin must be a string or null".to_string()))?,
            )
        };

        let origin_trimmed = origin_opt
            .as_deref()
            .map(str::trim)
            .filter(|s| !s.is_empty());

        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet_name);
            this.engine
                .set_sheet_origin(&sheet, origin_trimmed)
                .map_err(|err| js_err(err.to_string()))
        })
    }
    #[wasm_bindgen(js_name = "toJson")]
    pub fn to_json(&self) -> Result<String, JsValue> {
        #[derive(Serialize)]
        struct WorkbookJson<'a> {
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "localeId")]
            locale_id: Option<&'a str>,
            #[serde(rename = "formulaLanguage")]
            formula_language: WorkbookFormulaLanguageDto,
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "textCodepage")]
            text_codepage: Option<u16>,
            #[serde(default, skip_serializing_if = "Vec::is_empty", rename = "sheetOrder")]
            sheet_order: Vec<String>,
            sheets: BTreeMap<String, SheetJson>,
        }

        #[derive(Serialize)]
        struct SheetJson {
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "rowCount")]
            row_count: Option<u32>,
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "colCount")]
            col_count: Option<u32>,
            #[serde(default, skip_serializing_if = "Option::is_none")]
            visibility: Option<&'static str>,
            #[serde(default, skip_serializing_if = "Option::is_none", rename = "tabColor")]
            tab_color: Option<TabColor>,
            #[serde(default, skip_serializing_if = "BTreeMap::is_empty", rename = "cellPhonetics")]
            cell_phonetics: BTreeMap<String, String>,
            cells: BTreeMap<String, JsonValue>,
        }

        let mut sheets = BTreeMap::new();
        for (sheet_name, cells) in &self.inner.sheets {
            let mut out_cells = BTreeMap::new();
            let mut cell_phonetics: BTreeMap<String, String> = BTreeMap::new();
            for (address, input) in cells {
                // Ensure we never serialize explicit `null` cells; empty cells are
                // omitted from the sparse workbook representation.
                if input.is_null() {
                    continue;
                }
                out_cells.insert(address.clone(), input.clone());

                if let Some(phonetic) = self.inner.engine.get_cell_phonetic(sheet_name, address) {
                    // Preserve phonetic guide metadata used by Excel's `PHONETIC()` function.
                    // Note: the stored metadata may be an empty string (presence without text).
                    cell_phonetics.insert(address.clone(), phonetic.to_string());
                }
            }
            let (rows, cols) = self
                .inner
                .engine
                .sheet_dimensions(sheet_name)
                .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS));
            let row_count = (rows != EXCEL_MAX_ROWS).then_some(rows);
            let col_count = (cols != EXCEL_MAX_COLS).then_some(cols);

            let visibility = self.inner.sheet_visibility.get(sheet_name).and_then(|v| match v {
                SheetVisibility::Hidden => Some("hidden"),
                SheetVisibility::VeryHidden => Some("veryHidden"),
                SheetVisibility::Visible => None,
            });
            let tab_color = self.inner.sheet_tab_colors.get(sheet_name).cloned();
            sheets.insert(
                sheet_name.clone(),
                SheetJson {
                    row_count,
                    col_count,
                    visibility,
                    tab_color,
                    cell_phonetics,
                    cells: out_cells,
                },
            );
        }

        // Preserve the workbook formula locale id so round-tripping through the JSON workbook
        // schema does not lose locale-aware formula input semantics.
        //
        // Note: `toJson()` always emits canonical (en-US) formulas today. The `formulaLanguage`
        // field disambiguates this for `fromJson`, especially for comma-decimal locales like
        // `de-DE` where canonical `,` argument separators could be misinterpreted as decimal commas.
        let locale_id = if self.inner.formula_locale.id == EN_US.id {
            None
        } else {
            Some(self.inner.formula_locale.id)
        };

        let text_codepage = {
            let codepage = self.inner.engine.text_codepage();
            (codepage != 1252).then_some(codepage)
        };

        // Preserve sheet tab order so clients can round-trip through `toJson`/`fromJson` without
        // changing 3D reference semantics (`Sheet1:Sheet3!A1`) or worksheet functions like `SHEET()`.
        //
        // Note: `sheetOrder` must reference the same identifiers used as keys in `sheets` (stable
        // sheet ids), not user-visible display names. Display names may differ when hosts call
        // `setSheetDisplayName` (e.g. DocumentController stable sheet ids).
        let sheet_order = self.inner.engine.sheet_keys_in_order();

        serde_json::to_string(&WorkbookJson {
            locale_id,
            formula_language: WorkbookFormulaLanguageDto::Canonical,
            text_codepage,
            sheet_order,
            sheets,
        })
            .map_err(|err| js_err(format!("invalid workbook json: {err}")))
    }

    /// Return a lightweight workbook metadata payload (sheet list + dimensions + best-effort used ranges)
    /// without materializing the full workbook JSON string returned by `toJson()`.
    ///
    /// This is intended for web clients that need to open `.xlsx` bytes and quickly determine the
    /// available sheets / their used ranges without scanning every cell key in JS.
    #[wasm_bindgen(js_name = "getWorkbookInfo")]
    pub fn get_workbook_info(&self) -> Result<JsValue, JsValue> {
        let obj = Object::new();
        object_set(&obj, "path", &JsValue::NULL)?;
        object_set(&obj, "origin_path", &JsValue::NULL)?;

        let sheets_out = Array::new();

        // Prefer the engine's sheet tab order instead of the `BTreeMap` ordering of the sparse input
        // maps so UI clients (and sheet-indexed functions) observe Excel-like semantics.
        //
        // Use stable sheet keys (the identifiers used as keys in `toJson()`/`fromJson()`), not
        // display names, so we can look up persisted inputs and metadata maps keyed by sheet id.
        let keys_in_order = self.inner.engine.sheet_keys_in_order();
        let empty_cells: BTreeMap<String, JsonValue> = BTreeMap::new();

        let push_sheet = |sheet_key: &str, cells: &BTreeMap<String, JsonValue>| -> Result<(), JsValue> {
            let sheet_obj = Object::new();
            object_set(&sheet_obj, "id", &JsValue::from_str(sheet_key))?;
            let display_name = self
                .inner
                .engine
                .sheet_id(sheet_key)
                .and_then(|id| self.inner.engine.sheet_name(id))
                .unwrap_or(sheet_key);
            object_set(&sheet_obj, "name", &JsValue::from_str(display_name))?;

            if let Some(visibility) = self.inner.sheet_visibility.get(sheet_key).copied() {
                let value = match visibility {
                    SheetVisibility::Visible => "visible",
                    SheetVisibility::Hidden => "hidden",
                    SheetVisibility::VeryHidden => "veryHidden",
                };
                object_set(&sheet_obj, "visibility", &JsValue::from_str(value))?;
            }

            if let Some(color) = self.inner.sheet_tab_colors.get(sheet_key) {
                use serde::ser::Serialize as _;
                let js = color
                    .serialize(&serde_wasm_bindgen::Serializer::json_compatible())
                    .map_err(|err| js_err(err.to_string()))?;
                object_set(&sheet_obj, "tabColor", &js)?;
            }

            // Include sheet dimensions when they differ from Excel defaults (to match `toJson()`).
            let (rows, cols) = self
                .inner
                .engine
                .sheet_dimensions(sheet_key)
                .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS));
            if rows != EXCEL_MAX_ROWS {
                object_set(&sheet_obj, "rowCount", &JsValue::from_f64(rows as f64))?;
            }
            if cols != EXCEL_MAX_COLS {
                object_set(&sheet_obj, "colCount", &JsValue::from_f64(cols as f64))?;
            }

            // Best-effort used range derived from the sparse input maps (scalar + rich).
            let mut used_start_row: Option<u32> = None;
            let mut used_end_row: u32 = 0;
            let mut used_start_col: u32 = 0;
            let mut used_end_col: u32 = 0;

            for (address, input) in cells {
                // Explicit nulls should not affect used range tracking (sparse semantics).
                if input.is_null() {
                    continue;
                }
                let Ok(cell_ref) = CellRef::from_a1(address) else {
                    continue;
                };

                match used_start_row {
                    None => {
                        used_start_row = Some(cell_ref.row);
                        used_end_row = cell_ref.row;
                        used_start_col = cell_ref.col;
                        used_end_col = cell_ref.col;
                    }
                    Some(start_row) => {
                        used_start_row = Some(start_row.min(cell_ref.row));
                        used_end_row = used_end_row.max(cell_ref.row);
                        used_start_col = used_start_col.min(cell_ref.col);
                        used_end_col = used_end_col.max(cell_ref.col);
                    }
                }
            }

            if let Some(rich_cells) = self.inner.sheets_rich.get(sheet_key) {
                for (address, input) in rich_cells {
                    if input.is_empty() {
                        continue;
                    }
                    let Ok(cell_ref) = CellRef::from_a1(address) else {
                        continue;
                    };
                    match used_start_row {
                        None => {
                            used_start_row = Some(cell_ref.row);
                            used_end_row = cell_ref.row;
                            used_start_col = cell_ref.col;
                            used_end_col = cell_ref.col;
                        }
                        Some(start_row) => {
                            used_start_row = Some(start_row.min(cell_ref.row));
                            used_end_row = used_end_row.max(cell_ref.row);
                            used_start_col = used_start_col.min(cell_ref.col);
                            used_end_col = used_end_col.max(cell_ref.col);
                        }
                    }
                }
            }

            if let Some(start_row) = used_start_row {
                let used_obj = Object::new();
                object_set(&used_obj, "start_row", &JsValue::from_f64(start_row as f64))?;
                object_set(
                    &used_obj,
                    "end_row",
                    &JsValue::from_f64(used_end_row as f64),
                )?;
                object_set(
                    &used_obj,
                    "start_col",
                    &JsValue::from_f64(used_start_col as f64),
                )?;
                object_set(
                    &used_obj,
                    "end_col",
                    &JsValue::from_f64(used_end_col as f64),
                )?;
                object_set(&sheet_obj, "usedRange", &used_obj.into())?;
            }

            sheets_out.push(&sheet_obj);
            Ok(())
        };

        if keys_in_order.is_empty() {
            for (sheet_name, cells) in &self.inner.sheets {
                push_sheet(sheet_name, cells)?;
            }
        } else {
            for sheet_key in &keys_in_order {
                let cells = self.inner.sheets.get(sheet_key).unwrap_or(&empty_cells);
                push_sheet(sheet_key, cells)?;
            }
        }

        object_set(&obj, "sheets", &sheets_out.into())?;
        Ok(obj.into())
    }

    #[wasm_bindgen(js_name = "getCell")]
    pub fn get_cell(&self, address: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let cell = self.inner.get_cell_data(sheet, &address)?;
        cell_data_to_js(&cell)
    }

    /// Returns the per-cell style id, or `0` if the cell has the default style.
    ///
    /// Note: This is currently a narrow interop hook so JS callers can preserve formatting when
    /// clearing cell contents.
    #[wasm_bindgen(js_name = "getCellStyleId")]
    pub fn get_cell_style_id(
        &self,
        address: String,
        sheet: Option<String>,
    ) -> Result<u32, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        self.inner.get_cell_style_id_internal(sheet, &address)
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

    #[wasm_bindgen(js_name = "setCellPhonetic")]
    pub fn set_cell_phonetic(
        &mut self,
        address: String,
        phonetic: Option<String>,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        // Preserve explicit-recalc semantics even when the workbook's calcMode is automatic.
        self.inner.with_manual_calc_mode(|this| {
            let sheet = this.ensure_sheet(sheet);
            this.engine
                .set_cell_phonetic(&sheet, &address, phonetic)
                .map_err(|err| js_err(err.to_string()))
        })
    }

    #[wasm_bindgen(js_name = "getCellPhonetic")]
    pub fn get_cell_phonetic(
        &self,
        address: String,
        sheet: Option<String>,
    ) -> Result<Option<String>, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet)?.to_string();
        let address = WorkbookState::parse_address(&address)?.to_a1();
        Ok(self
            .inner
            .engine
            .get_cell_phonetic(&sheet, &address)
            .map(|s| s.to_string()))
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
        let start_row = range.start.row;
        let start_col = range.start.col;

        let values = self
            .inner
            .engine
            .get_range_values(&sheet, range)
            .map_err(|err| js_err(err.to_string()))?;

        let sheet_cells = self.inner.sheets.get(&sheet);
        let sheet_js = JsValue::from_str(&sheet);
        let key_sheet = JsValue::from_str("sheet");
        let key_address = JsValue::from_str("address");
        let key_input = JsValue::from_str("input");
        let key_value = JsValue::from_str("value");

        let outer = Array::new_with_length(values.len() as u32);
        // Reuse buffers to avoid per-cell string allocations (both for input lookup and
        // for emitting the `address` string field).
        let mut addr_buf = String::with_capacity(16);
        let mut row_buf = String::with_capacity(16);
        for (row_off, row_values) in values.into_iter().enumerate() {
            let row = start_row + row_off as u32;
            row_buf.clear();
            push_u64_decimal(u64::from(row).saturating_add(1), &mut row_buf);
            let inner = Array::new_with_length(row_values.len() as u32);
            for (col_off, engine_value) in row_values.into_iter().enumerate() {
                let col = start_col + col_off as u32;
                addr_buf.clear();
                push_a1_col_name(col, &mut addr_buf);
                addr_buf.push_str(&row_buf);

                let input = if let Some(cells) = sheet_cells {
                    cells
                        .get(addr_buf.as_str())
                        .map(json_scalar_to_js)
                        .unwrap_or(JsValue::NULL)
                } else {
                    JsValue::NULL
                };
                let value = engine_value_to_js_scalar(engine_value);

                let obj = Object::new();
                Reflect::set(&obj, &key_sheet, &sheet_js)?;
                Reflect::set(&obj, &key_address, &JsValue::from_str(&addr_buf))?;
                Reflect::set(&obj, &key_input, &input)?;
                Reflect::set(&obj, &key_value, &value)?;
                inner.set(col_off as u32, obj.into());
            }
            outer.set(row_off as u32, inner.into());
        }

        Ok(outer.into())
    }

    #[wasm_bindgen(js_name = "getRangeCompact")]
    pub fn get_range_compact(
        &self,
        range: String,
        sheet: Option<String>,
    ) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet)?;
        let range = WorkbookState::parse_range(&range)?;
        let start_row = range.start.row;
        let start_col = range.start.col;

        // Return a nested JS array (rows -> columns) with a compact per-cell payload:
        //   [input, value]
        // This avoids allocating redundant `{sheet,address}` strings per cell, which the
        // TS backend discards anyway.
        let sheet_cells = self.inner.sheets.get(sheet);
        let values = self
            .inner
            .engine
            .get_range_values(sheet, range)
            .map_err(|err| js_err(err.to_string()))?;

        let outer = Array::new_with_length(values.len() as u32);
        // Reuse buffers to avoid per-cell string allocations while looking up sparse inputs.
        let mut addr_buf = String::with_capacity(16);
        let mut row_buf = String::with_capacity(16);
        for (row_off, row_values) in values.into_iter().enumerate() {
            let row = start_row + row_off as u32;
            row_buf.clear();
            push_u64_decimal(u64::from(row).saturating_add(1), &mut row_buf);
            let inner = Array::new_with_length(row_values.len() as u32);
            for (col_off, engine_value) in row_values.into_iter().enumerate() {
                let col = start_col + col_off as u32;
                let input = if let Some(cells) = sheet_cells {
                    addr_buf.clear();
                    push_a1_col_name(col, &mut addr_buf);
                    addr_buf.push_str(&row_buf);
                    cells
                        .get(addr_buf.as_str())
                        .map(json_scalar_to_js)
                        .unwrap_or(JsValue::NULL)
                } else {
                    JsValue::NULL
                };
                let value = engine_value_to_js_scalar(engine_value);

                let cell = Array::new_with_length(2);
                cell.set(0, input);
                cell.set(1, value);
                inner.set(col_off as u32, cell.into());
            }
            outer.set(row_off as u32, inner.into());
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
    pub fn goal_seek(&mut self, params: JsValue) -> Result<JsValue, JsValue> {
        ensure_rust_constructors_run();

        let params: GoalSeekRequestDto =
            serde_wasm_bindgen::from_value(params).map_err(|err| js_err(err.to_string()))?;
        let sheet = params.sheet.as_deref().unwrap_or(DEFAULT_SHEET).trim();
        let sheet = if sheet.is_empty() { DEFAULT_SHEET } else { sheet };

        let target_cell = params.target_cell.trim();
        if target_cell.is_empty() {
            return Err(js_err("targetCell must be a non-empty string"));
        }
        let changing_cell = params.changing_cell.trim();
        if changing_cell.is_empty() {
            return Err(js_err("changingCell must be a non-empty string"));
        }

        if !params.target_value.is_finite() {
            return Err(js_err("targetValue must be a finite number"));
        }

        if let Some(tol) = params.tolerance {
            if !tol.is_finite() {
                return Err(js_err("tolerance must be a finite number"));
            }
            if !(tol > 0.0) {
                return Err(js_err("tolerance must be > 0"));
            }
        }
        if let Some(step) = params.derivative_step {
            if !step.is_finite() {
                return Err(js_err("derivativeStep must be a finite number"));
            }
            if !(step > 0.0) {
                return Err(js_err("derivativeStep must be > 0"));
            }
        }
        if let Some(min) = params.min_derivative {
            if !min.is_finite() {
                return Err(js_err("minDerivative must be a finite number"));
            }
            if !(min > 0.0) {
                return Err(js_err("minDerivative must be > 0"));
            }
        }
        if let Some(max) = params.max_iterations {
            if max == 0 {
                return Err(js_err("maxIterations must be > 0"));
            }
        }
        if let Some(max) = params.max_bracket_expansions {
            if max == 0 {
                return Err(js_err("maxBracketExpansions must be > 0"));
            }
        }

        let tuning = GoalSeekTuning {
            max_iterations: params.max_iterations.map(|v| v as usize),
            tolerance: params.tolerance,
            derivative_step: params.derivative_step,
            min_derivative: params.min_derivative,
            max_bracket_expansions: params.max_bracket_expansions.map(|v| v as usize),
        };

        let (result, changes) = self.inner.goal_seek_internal(
            sheet,
            target_cell,
            params.target_value,
            changing_cell,
            tuning,
        )?;

        let out = GoalSeekResponseDto { result, changes };
        serde_wasm_bindgen::to_value(&out).map_err(|err| js_err(err.to_string()))
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
            writes: Vec<PivotCellWrite>,
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
        let cache = self
            .inner
            .engine
            .pivot_cache_from_range(&sheet, range)
            .map_err(|err| js_err(err.to_string()))?;

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
        let cache = self
            .inner
            .engine
            .pivot_cache_from_range(&sheet, range)
            .map_err(|err| js_err(err.to_string()))?;

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

fn xlsb_error_code_to_model_error(code: u8) -> formula_model::ErrorValue {
    use core::str::FromStr as _;

    formula_xlsb::errors::xlsb_error_literal(code)
        .and_then(|lit| formula_model::ErrorValue::from_str(lit).ok())
        .unwrap_or(formula_model::ErrorValue::Unknown)
}

fn xlsb_to_model_workbook(
    wb: &formula_xlsb::XlsbWorkbook,
) -> Result<formula_model::Workbook, formula_xlsb::Error> {
    use formula_model::{
        normalize_formula_text, CalculationMode as ModelCalculationMode, CellRef,
        CellValue as ModelCellValue, DateSystem, DefinedNameScope, SheetVisibility as ModelSheetVisibility,
        Style, Workbook as ModelWorkbook,
    };

    let mut out = ModelWorkbook::new();
    out.date_system = if wb.workbook_properties().date_system_1904 {
        DateSystem::Excel1904
    } else {
        DateSystem::Excel1900
    };
    if let Some(calc_mode) = wb.workbook_properties().calc_mode {
        out.calc_settings.calculation_mode = match calc_mode {
            formula_xlsb::CalcMode::Auto => ModelCalculationMode::Automatic,
            formula_xlsb::CalcMode::Manual => ModelCalculationMode::Manual,
            formula_xlsb::CalcMode::AutoExceptTables => ModelCalculationMode::AutomaticNoTable,
        };
    }
    if let Some(full_calc_on_load) = wb.workbook_properties().full_calc_on_load {
        out.calc_settings.full_calc_on_load = full_calc_on_load;
    }

    // Best-effort style mapping: XLSB cell records reference an XF index.
    //
    // `formula-xlsb` currently only exposes number formats (`numFmtId`/`ifmt`) via `Styles`.
    // Preserve those for downstream consumers like date inference in pivot tables.
    let mut xf_to_style_id: Vec<u32> = Vec::with_capacity(wb.styles().len());
    for xf_idx in 0..wb.styles().len() {
        let info = wb
            .styles()
            .get(xf_idx as u32)
            .expect("xf index within wb.styles().len()");
        if info.num_fmt_id == 0 {
            xf_to_style_id.push(0);
            continue;
        }
        let style_id = info
            .number_format
            .as_deref()
            .filter(|fmt| !fmt.is_empty())
            .map(|fmt| {
                out.intern_style(Style {
                    number_format: Some(fmt.to_string()),
                    ..Default::default()
                })
            })
            .unwrap_or(0);
        xf_to_style_id.push(style_id);
    }

    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> =
        Vec::with_capacity(wb.sheet_metas().len());

    for (sheet_index, meta) in wb.sheet_metas().iter().enumerate() {
        let sheet_id = out
            .add_sheet(meta.name.clone())
            .map_err(|err| formula_xlsb::Error::InvalidSheetName(format!("{}: {err}", meta.name)))?;
        worksheet_ids_by_index.push(sheet_id);

        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");
        sheet.visibility = match meta.visibility {
            formula_xlsb::SheetVisibility::Visible => ModelSheetVisibility::Visible,
            formula_xlsb::SheetVisibility::Hidden => ModelSheetVisibility::Hidden,
            formula_xlsb::SheetVisibility::VeryHidden => ModelSheetVisibility::VeryHidden,
        };

        wb.for_each_cell(sheet_index, |cell| {
            let cell_ref = CellRef::new(cell.row, cell.col);
            let style_id = xf_to_style_id
                .get(cell.style as usize)
                .copied()
                .unwrap_or(0);

            match cell.value {
                formula_xlsb::CellValue::Blank => {}
                formula_xlsb::CellValue::Number(v) => sheet.set_value(cell_ref, ModelCellValue::Number(v)),
                formula_xlsb::CellValue::Bool(v) => sheet.set_value(cell_ref, ModelCellValue::Boolean(v)),
                formula_xlsb::CellValue::Text(s) => sheet.set_value(cell_ref, ModelCellValue::String(s)),
                formula_xlsb::CellValue::Error(code) => sheet.set_value(
                    cell_ref,
                    ModelCellValue::Error(xlsb_error_code_to_model_error(code)),
                ),
            };

            // Cells with non-zero style ids must be stored, even if blank, matching Excel's ability
            // to format empty cells.
            if style_id != 0 {
                sheet.set_style_id(cell_ref, style_id);
            }

            if let Some(formula) = cell.formula.and_then(|f| f.text) {
                if let Some(normalized) = normalize_formula_text(&formula) {
                    sheet.set_formula(cell_ref, Some(normalized));
                }
            }

            // Best-effort phonetic guide (furigana) extraction.
            if let Some(phonetic) = cell
                .preserved_string
                .as_ref()
                .and_then(|s| s.phonetic_text())
            {
                let mut model_cell = sheet.cell(cell_ref).cloned().unwrap_or_default();
                model_cell.phonetic = Some(phonetic);
                sheet.set_cell(cell_ref, model_cell);
            }
        })?;
    }

    // Defined names: parsed from `xl/workbook.bin` `BrtName` records.
    for name in wb.defined_names() {
        let Some(formula) = name.formula.as_ref().and_then(|f| f.text.as_deref()) else {
            continue;
        };
        let Some(refers_to) = normalize_formula_text(formula) else {
            continue;
        };

        let (scope, local_sheet_id) = match name.scope_sheet.and_then(|idx| {
            worksheet_ids_by_index
                .get(idx as usize)
                .copied()
                .map(|id| (idx, id))
        }) {
            Some((local_sheet_id, sheet_id)) => {
                (DefinedNameScope::Sheet(sheet_id), Some(local_sheet_id))
            }
            None => (DefinedNameScope::Workbook, None),
        };

        // Best-effort: ignore invalid/duplicate names so we can still import the workbook.
        let _ = out.create_defined_name(
            scope,
            name.name.clone(),
            refers_to,
            name.comment.clone(),
            name.hidden,
            local_sheet_id,
        );
    }

    Ok(out)
}

// Native-only helpers for integration tests and tooling.
//
// These are intentionally not exported to JS/WASM. The JS worker protocol uses the regular
// `getCell`/`recalculate` APIs; native tests need direct access to the engine value surface so they
// don't depend on `js_sys` shims.
#[cfg(not(target_arch = "wasm32"))]
impl WasmWorkbook {
    #[doc(hidden)]
    pub fn debug_get_engine_value(&self, sheet: &str, address: &str) -> EngineValue {
        self.inner.engine.get_cell_value(sheet, address)
    }

    #[doc(hidden)]
    pub fn debug_recalculate(&mut self) -> Vec<CellChange> {
        self.inner
            .recalculate_internal(None)
            .expect("recalculate should succeed")
    }

    #[doc(hidden)]
    pub fn debug_set_workbook_file_metadata(
        &mut self,
        directory: Option<&str>,
        filename: Option<&str>,
    ) {
        self.inner
            .set_workbook_file_metadata_internal(directory, filename)
            .expect("set_workbook_file_metadata should succeed");
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn supported_locale_ids_are_sorted_and_contain_known_ids() {
        let ids = supported_locale_ids_sorted();

        for expected in ["de-DE", "en-US", "es-ES", "fr-FR"] {
            assert!(
                ids.contains(&expected),
                "expected locale list to contain {expected:?}, got {ids:?}"
            );
        }

        assert!(
            ids.windows(2).all(|pair| pair[0] <= pair[1]),
            "expected locale list to be sorted, got {ids:?}"
        );
    }

    #[test]
    fn from_json_sheet_order_controls_3d_reference_semantics() {
        // 3D references (`Sheet1:Sheet3!A1`) depend on sheet tab order. The JSON workbook schema is
        // map-based, so without an explicit order hint the key iteration order is lost during parse
        // (BTreeMap). Verify that `sheetOrder` preserves the intended semantics.

        let workbook = json!({
            "sheets": {
                "A": { "cells": { "A1": 10 } },
                "B": { "cells": { "A1": 1, "A2": "=SUM(B:C!A1)" } },
                "C": { "cells": { "A1": 100 } },
            }
        })
        .to_string();

        // Default behavior: sheets are created in sorted-key order (A, B, C), so `B:C` includes
        // only B and C.
        let mut wb = WasmWorkbook::from_json(&workbook).unwrap();
        wb.inner.recalculate_internal(None).unwrap();
        let value = wb.inner.get_cell_data("B", "A2").unwrap().value;
        assert_eq!(value.as_f64().unwrap(), 101.0);

        // With `sheetOrder`, sheets are created in the specified tab order (B, A, C), so `B:C`
        // includes B, A, and C.
        let workbook_with_order = json!({
            "sheetOrder": ["B", "A", "C"],
            "sheets": {
                "A": { "cells": { "A1": 10 } },
                "B": { "cells": { "A1": 1, "A2": "=SUM(B:C!A1)" } },
                "C": { "cells": { "A1": 100 } },
            }
        })
        .to_string();

        let mut wb_ordered = WasmWorkbook::from_json(&workbook_with_order).unwrap();
        wb_ordered.inner.recalculate_internal(None).unwrap();
        let value_ordered = wb_ordered.inner.get_cell_data("B", "A2").unwrap().value;
        assert_eq!(value_ordered.as_f64().unwrap(), 111.0);
    }

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
    fn parse_formula_partial_normalizes_xlfn_prefix_in_fallback_contexts() {
        // Unterminated string literals cause `formula-engine::parse_formula_partial` to fail
        // during lexing, which means we fall back to `scan_fallback_function_context`. Ensure
        // those contexts are still normalized to match the canonical function catalog.
        let ctx = scan_fallback_function_context(r#"=_xlfn.SEQUENCE(1,"hello"#, ',').unwrap();
        assert_eq!(ctx.arg_index, 1);
        assert_eq!(normalize_function_context_name(&ctx.name, None), "SEQUENCE");

        // Locale-aware canonicalization should run before stripping the `_xlfn.` prefix.
        let localized =
            scan_fallback_function_context(r#"=_xlfn.SEQUENZ(1;"hallo"#, ';').unwrap();
        assert_eq!(localized.arg_index, 1);
        let de_de = get_locale("de-DE").expect("expected de-DE locale to be registered");
        assert_eq!(
            normalize_function_context_name(&localized.name, Some(de_de)),
            "SEQUENCE"
        );
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
    fn fallback_context_scanner_ignores_commas_in_external_workbook_prefixes_with_brackets_in_workbook_name() {
        // Workbook names may contain literal `[` characters, but workbook prefixes are not nested.
        // The scanner should still treat the comma after the external name reference as the argument
        // separator.
        let ctx = scan_fallback_function_context("=SUM([A1[Name.xlsx]MyName,1", ',').unwrap();
        assert_eq!(ctx.name, "SUM");
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

        // `toJson()` emits canonical (en-US) formula syntax today; `formulaLanguage` disambiguates
        // this from localized formulas for round-trips in comma-decimal locales like `de-DE`.
        assert_eq!(parsed["formulaLanguage"], json!("canonical"));
        // `toJson()` should include sheet tab order so `fromJson` can preserve 3D reference semantics.
        assert_eq!(parsed["sheetOrder"], json!(["Sheet1"]));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!(1.0));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));

        let wb2 = WasmWorkbook::from_json(&json_str).unwrap();
        let json_str2 = wb2.to_json().unwrap();
        let parsed2: serde_json::Value = serde_json::from_str(&json_str2).unwrap();
        assert_eq!(parsed2["formulaLanguage"], json!("canonical"));
        assert_eq!(parsed2["sheetOrder"], json!(["Sheet1"]));
        assert_eq!(parsed2["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_json_honors_formula_language_canonical_for_comma_decimal_locales() {
        // In comma-decimal locales like de-DE, a canonical formula like `=LOG(8,2)` must not be
        // interpreted as the localized decimal literal `8,2` (8.2). `formulaLanguage=canonical`
        // disambiguates the intended syntax for round-trips.
        let input = json!({
            "localeId": "de-DE",
            "formulaLanguage": "canonical",
            "sheets": {
                "Sheet1": {
                    "cells": {
                        "A1": "=LOG(8,2)"
                    }
                }
            }
        })
        .to_string();

        let mut wb = WasmWorkbook::from_json(&input).unwrap();
        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(3.0)
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn to_json_preserves_sheet_tab_order_roundtrip() {
        // `toJson()` should preserve the engine sheet tab order even though the `sheets` payload is
        // map-based (sorted key order).
        let mut state = WorkbookState::new_empty();
        state.ensure_sheet("B");
        state.ensure_sheet("A");
        state.ensure_sheet("C");

        let wb = WasmWorkbook { inner: state };
        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        assert_eq!(parsed["sheetOrder"], json!(["B", "A", "C"]));

        // Round-trip back through `fromJson()` to ensure `sheetOrder` is respected on hydration.
        let wb2 = WasmWorkbook::from_json(&json_str).unwrap();
        let json_str2 = wb2.to_json().unwrap();
        let parsed2: serde_json::Value = serde_json::from_str(&json_str2).unwrap();
        assert_eq!(parsed2["sheetOrder"], json!(["B", "A", "C"]));
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn to_json_uses_stable_sheet_keys_when_display_names_differ() {
        // Hosts like the desktop DocumentController can assign stable sheet ids (keys) that differ
        // from the user-visible tab name (display name). The JSON workbook schema uses the stable
        // keys as sheet map keys; `sheetOrder` must reference those same keys so tab ordering
        // round-trips correctly through `toJson()`/`fromJson()`.
        let mut state = WorkbookState::new_empty();
        state.ensure_sheet("Sheet1");
        state.ensure_sheet("sheet_2");
        state.engine.set_sheet_display_name("sheet_2", "Budget");

        let wb = WasmWorkbook { inner: state };
        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        assert_eq!(parsed["sheetOrder"], json!(["Sheet1", "sheet_2"]));
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_json_accepts_tab_color_string_and_ignores_unknown_visibility() {
        // Some snapshot producers represent tab colors as an ARGB string. Be tolerant so older
        // payloads continue to hydrate after adding structured `tabColor` metadata.
        let input = json!({
            "sheets": {
                "Sheet1": {
                    "visibility": "not_a_real_visibility",
                    "tabColor": "ffff0000",
                    "cells": {}
                }
            }
        })
        .to_string();
        let wb = WasmWorkbook::from_json(&input).unwrap();
        let json_str = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();
        let sheet = parsed["sheets"]["Sheet1"]
            .as_object()
            .expect("sheet should serialize as an object");

        assert_eq!(parsed["sheetOrder"], json!(["Sheet1"]));
        assert!(
            !sheet.contains_key("visibility"),
            "unknown visibility should be treated as default/omitted"
        );
        assert_eq!(sheet["tabColor"]["rgb"], json!("FFFF0000"));
    }

    #[test]
    fn from_xlsx_bytes_imports_calc_settings_into_engine() {
        let bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsx/tests/fixtures/calc_settings.xlsx"
        ));

        let wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
        let settings = wb.inner.engine.calc_settings();

        assert_eq!(settings.calculation_mode, CalculationMode::Manual);
        assert!(settings.calculate_before_save);
        assert!(settings.iterative.enabled);
        assert_eq!(settings.iterative.max_iterations, 10);
        assert!((settings.iterative.max_change - 0.0001).abs() < 1e-12);
        assert!(settings.full_precision);
        assert!(
            !settings.full_calc_on_load,
            "fixture does not set fullCalcOnLoad, default should be false"
        );
    }

    #[test]
    fn from_encrypted_xlsx_bytes_supports_xlsb_payloads() {
        // `fromEncryptedXlsxBytes` historically only supported decrypted `.xlsx` payloads.
        // Office-encrypted files can also wrap ZIP-based `.xlsb` workbooks, which should now be
        // supported in the WASM bindings.
        let xlsb_bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsb/tests/fixtures/simple.xlsb"
        ));

        let password = "secret";
        let ole_bytes = formula_office_crypto::encrypt_package_to_ole(
            xlsb_bytes,
            password,
            formula_office_crypto::EncryptOptions::default(),
        )
        .expect("encrypt xlsb package to OLE");

        let mut wb =
            WasmWorkbook::from_encrypted_xlsx_bytes(&ole_bytes, password.to_string()).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text("Hello".into())
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Number(42.5)
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Number(85.0)
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_forces_manual_calc_mode_even_when_workbook_is_automatic() {
        use std::io::Cursor;

        let mut workbook = formula_model::Workbook::new();
        workbook.calc_settings.calculation_mode = CalculationMode::Automatic;
        workbook.calc_settings.calculate_before_save = false;
        workbook.calc_settings.iterative.enabled = true;
        workbook.calc_settings.iterative.max_iterations = 7;
        workbook.calc_settings.iterative.max_change = 0.123;
        workbook.calc_settings.full_precision = false;
        workbook.calc_settings.full_calc_on_load = true;

        let sheet_id = workbook.add_sheet("Sheet1").unwrap();
        workbook
            .sheet_mut(sheet_id)
            .unwrap()
            .set_value_a1("A1", CellValue::Number(1.0))
            .unwrap();

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        let settings = wb.inner.engine.calc_settings();

        // The WASM worker protocol expects manual recalc regardless of what the XLSX requested.
        assert_eq!(settings.calculation_mode, CalculationMode::Manual);

        // Other workbook calc settings should still round-trip.
        assert!(!settings.calculate_before_save);
        assert!(settings.iterative.enabled);
        assert_eq!(settings.iterative.max_iterations, 7);
        assert!((settings.iterative.max_change - 0.123).abs() < 1e-12);
        assert!(!settings.full_precision);
        assert!(settings.full_calc_on_load);
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn set_workbook_file_metadata_does_not_trigger_automatic_recalc() {
        // Even if callers configure the engine to use automatic calculation mode, the WASM worker
        // protocol relies on explicit `recalculate()` calls to surface value-change deltas back to
        // JS. Metadata setters must therefore avoid triggering an implicit recalc (which would
        // update cell values without reporting them as deltas).

        let mut wb = WasmWorkbook::new();
        wb.inner.engine.set_calc_settings(CalcSettings {
            calculation_mode: CalculationMode::Automatic,
            ..CalcSettings::default()
        });

        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", json!("=CELL(\"filename\")"))
            .unwrap();
        wb.debug_recalculate();
        assert_eq!(
            wb.debug_get_engine_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text(String::new())
        );

        wb.debug_set_workbook_file_metadata(Some("/tmp/"), Some("book.xlsx"));

        // The setter should not have triggered an automatic recalc.
        assert_eq!(
            wb.inner.engine.calc_settings().calculation_mode,
            CalculationMode::Automatic
        );
        assert_eq!(
            wb.debug_get_engine_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text(String::new())
        );

        let changes = wb.debug_recalculate();
        assert!(
            changes.iter().any(|change| {
                change.sheet == DEFAULT_SHEET
                    && change.address == "A1"
                    && change.value == json!("/tmp/[book.xlsx]Sheet1")
            }),
            "expected a value-change delta for A1 after recalc, got {changes:?}"
        );
        assert_eq!(
            wb.debug_get_engine_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text("/tmp/[book.xlsx]Sheet1".to_string())
        );
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

    fn build_inline_string_phonetic_fixture_xlsx() -> Vec<u8> {
        use std::io::{Cursor, Write};
        use zip::write::FileOptions;
        use zip::{CompressionMethod, ZipWriter};

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Base</t>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="4"><t>PHONETIC</t></rPh>
        </is>
      </c>
      <c r="B1">
        <f>PHONETIC(A1)</f>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    fn build_shared_strings_phonetic_fixture_xlsx() -> Vec<u8> {
        use std::io::{Cursor, Write};
        use zip::write::FileOptions;
        use zip::{CompressionMethod, ZipWriter};

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

        let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1">
        <f>PHONETIC(A1)</f>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Base</t>
    <rPh sb="0" eb="4"><t>PHONETIC</t></rPh>
  </si>
</sst>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    fn build_inline_string_formula_phonetic_fixture_xlsx() -> Vec<u8> {
        use std::io::{Cursor, Write};
        use zip::write::FileOptions;
        use zip::{CompressionMethod, ZipWriter};

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

        // A1 contains a formula with an inline string cached result that includes phonetic guides.
        // Some real-world XLSX producers encode string formula results using `t="inlineStr"` with
        // `<is>` content (instead of shared strings).
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <f>"Base"</f>
        <is>
          <t>Base</t>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="4"><t>PHONETIC</t></rPh>
        </is>
      </c>
      <c r="B1">
        <f>PHONETIC(A1)</f>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn from_xlsx_bytes_imports_cell_phonetic_metadata_for_phonetic_function() {
        let bytes = build_inline_string_phonetic_fixture_xlsx();
        let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(wb.inner.engine.get_cell_phonetic(DEFAULT_SHEET, "A1"), Some("PHONETIC"));
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("PHONETIC".to_string())
        );
    }

    #[test]
    fn from_xlsx_bytes_preserves_phonetic_guides_for_formula_cells() {
        let bytes = build_inline_string_formula_phonetic_fixture_xlsx();
        let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Text("Base".to_string())
        );
        assert_eq!(wb.inner.engine.get_cell_phonetic(DEFAULT_SHEET, "A1"), Some("PHONETIC"));
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("PHONETIC".to_string())
        );
    }

    #[test]
    fn from_xlsx_bytes_imports_shared_string_phonetic_metadata_for_phonetic_function() {
        let bytes = build_shared_strings_phonetic_fixture_xlsx();
        let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"), EngineValue::Text("Base".to_string()));
        assert_eq!(wb.inner.engine.get_cell_phonetic(DEFAULT_SHEET, "A1"), Some("PHONETIC"));
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("PHONETIC".to_string())
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
    fn from_xlsx_bytes_imports_hidden_columns_for_cell_width() {
        let bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/basic/row-col-attrs.xlsx"
        ));

        let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "D1", json!(r#"=CELL("width",C1)"#))
            .unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "D1"),
            EngineValue::Number(0.0)
        );
    }

    #[test]
    fn from_model_json_propagates_codepage_and_cell_phonetic_metadata() {
        let mut model = formula_model::Workbook::new();
        model.codepage = 932;

        let sheet_id = model.add_sheet("Sheet1").unwrap();
        let sheet = model.sheet_mut(sheet_id).unwrap();

        let mut cell = formula_model::Cell::new(formula_model::CellValue::String("漢字".to_string()));
        cell.phonetic = Some("かんじ".to_string());
        sheet.set_cell(formula_model::CellRef::from_a1("A1").unwrap(), cell);

        sheet.set_formula_a1("B1", Some("PHONETIC(A1)".to_string()))
            .unwrap();
        sheet.set_formula_a1("C1", Some("LENB(\"あ\")".to_string()))
            .unwrap();

        let json = serde_json::to_string(&model).unwrap();
        let mut wb = WasmWorkbook::from_model_json(json).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("かんじ".to_string())
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "C1"),
            EngineValue::Number(2.0)
        );
    }

    #[test]
    fn from_model_json_preserves_phonetic_metadata_for_formula_cells() {
        let mut model = formula_model::Workbook::new();
        let sheet_id = model.add_sheet("Sheet1").unwrap();
        let sheet = model.sheet_mut(sheet_id).unwrap();

        // A1 is a formula cell that returns a string; the model also stores cached value and
        // phonetic guide metadata.
        let mut cell = formula_model::Cell::new(formula_model::CellValue::String("漢字".to_string()));
        // `formula-model` stores formulas without a leading '='.
        cell.formula = Some("\"漢字\"".to_string());
        cell.phonetic = Some("かんじ".to_string());
        sheet.set_cell(formula_model::CellRef::from_a1("A1").unwrap(), cell);

        sheet.set_formula_a1("B1", Some("PHONETIC(A1)".to_string()))
            .unwrap();

        let json = serde_json::to_string(&model).unwrap();
        let mut wb = WasmWorkbook::from_model_json(json).unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb.inner.engine.get_cell_phonetic(DEFAULT_SHEET, "A1"),
            Some("かんじ")
        );
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("かんじ".to_string())
        );
    }

    #[test]
    fn to_json_and_from_json_roundtrip_cell_phonetic_metadata() {
        let mut wb = WasmWorkbook::new();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", JsonValue::String("漢字".to_string()))
            .unwrap();
        wb.set_cell_phonetic("A1".to_string(), Some("かんじ".to_string()), None)
            .unwrap();
        wb.inner
            .set_cell_internal(
                DEFAULT_SHEET,
                "B1",
                JsonValue::String("=PHONETIC(A1)".to_string()),
            )
            .unwrap();
        wb.inner.recalculate_internal(None).unwrap();

        let json = wb.to_json().unwrap();
        let mut wb2 = WasmWorkbook::from_json(&json).unwrap();
        wb2.inner.recalculate_internal(None).unwrap();

        assert_eq!(
            wb2.inner.engine.get_cell_phonetic(DEFAULT_SHEET, "A1"),
            Some("かんじ")
        );
        assert_eq!(
            wb2.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("かんじ".to_string())
        );
    }

    #[test]
    fn from_json_imports_text_codepage_and_to_json_roundtrips_it() {
        let json = serde_json::json!({
            "textCodepage": 932,
            "sheets": {
                "Sheet1": {
                    "cells": {
                        "A1": "=LENB(\"あ\")"
                    }
                }
            }
        })
        .to_string();

        let mut wb = WasmWorkbook::from_json(&json).unwrap();
        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(2.0)
        );

        let roundtrip = wb.to_json().unwrap();
        let parsed: serde_json::Value = serde_json::from_str(&roundtrip).unwrap();
        assert_eq!(parsed["textCodepage"], serde_json::json!(932));

        let mut wb2 = WasmWorkbook::from_json(&roundtrip).unwrap();
        wb2.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb2.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(2.0)
        );
    }

    #[test]
    fn set_text_codepage_api_updates_lenb_behavior() {
        let mut wb = WasmWorkbook::new();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", serde_json::json!("=LENB(\"あ\")"))
            .unwrap();

        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(1.0)
        );

        assert_eq!(wb.get_text_codepage(), 1252);
        wb.set_text_codepage(932).unwrap();
        assert_eq!(wb.get_text_codepage(), 932);

        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(2.0)
        );
    }

    #[test]
    fn set_locale_sets_text_codepage_for_dbcs_locales() {
        let mut wb = WasmWorkbook::new();
        assert_eq!(wb.get_text_codepage(), 1252);

        assert!(wb.set_locale("ja-JP".to_string()));
        assert_eq!(wb.get_text_codepage(), 932);

        assert!(wb.set_locale("zh-CN".to_string()));
        assert_eq!(wb.get_text_codepage(), 936);
    }

    #[test]
    fn set_locale_updates_lenb_behavior_for_ja_jp() {
        let mut wb = WasmWorkbook::new();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", serde_json::json!("=LENB(\"あ\")"))
            .unwrap();

        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(1.0)
        );

        assert!(wb.set_locale("ja-JP".to_string()));
        assert_eq!(wb.get_text_codepage(), 932);

        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "A1"),
            EngineValue::Number(2.0)
        );
    }

    #[test]
    fn get_cell_phonetic_api_roundtrips() {
        let mut wb = WasmWorkbook::new();
        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "A1", serde_json::json!("漢字"))
            .unwrap();
        wb.set_cell_phonetic("A1".to_string(), Some("かんじ".to_string()), None)
            .unwrap();

        let phonetic = wb.get_cell_phonetic("A1".to_string(), None).unwrap();
        assert_eq!(phonetic.as_deref(), Some("かんじ"));

        wb.inner
            .set_cell_internal(DEFAULT_SHEET, "B1", serde_json::json!("=PHONETIC(A1)"))
            .unwrap();
        wb.inner.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.inner.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("かんじ".to_string())
        );

        wb.set_cell_phonetic("A1".to_string(), None, None).unwrap();
        let cleared = wb.get_cell_phonetic("A1".to_string(), None).unwrap();
        assert!(cleared.is_none());
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
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_imports_cell_styles_for_pivot_date_inference() {
        use std::io::Cursor;

        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1").unwrap();

        // Add a date-like numeric column + number format applied via the cell style id.
        let date_style_id = workbook.styles.intern(formula_model::Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Default::default()
        });
        {
            let sheet = workbook.sheet_mut(sheet_id).unwrap();
            sheet
                .set_value_a1("A1", CellValue::String("Date".to_string()))
                .unwrap();
            sheet
                .set_value_a1("B1", CellValue::String("Amount".to_string()))
                .unwrap();

            let date_1 = ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;
            let date_2 = ymd_to_serial(ExcelDate::new(2024, 1, 16), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;

            sheet.set_value_a1("A2", CellValue::Number(date_1)).unwrap();
            sheet.set_value_a1("B2", CellValue::Number(10.0)).unwrap();
            sheet.set_value_a1("A3", CellValue::Number(date_2)).unwrap();
            sheet.set_value_a1("B3", CellValue::Number(20.0)).unwrap();

            sheet.set_style_id_a1("A2", date_style_id).unwrap();
            sheet.set_style_id_a1("A3", date_style_id).unwrap();
        }

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        let schema = wb
            .inner
            .get_pivot_schema_internal("Sheet1", "A1:B3", 10)
            .unwrap();

        let date_field = schema
            .fields
            .iter()
            .find(|f| f.name == "Date")
            .expect("expected Date field in schema");
        assert_eq!(date_field.field_type, pivot_engine::PivotFieldType::Date);

        let amount_field = schema
            .fields
            .iter()
            .find(|f| f.name == "Amount")
            .expect("expected Amount field in schema");
        assert_eq!(
            amount_field.field_type,
            pivot_engine::PivotFieldType::Number
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_imports_col_styles_for_pivot_date_inference() {
        use std::io::Cursor;

        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1").unwrap();

        // Apply the date number format via the column default style.
        let date_style_id = workbook.styles.intern(formula_model::Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Default::default()
        });

        {
            let sheet = workbook.sheet_mut(sheet_id).unwrap();
            sheet.set_col_style_id(0, Some(date_style_id));

            sheet
                .set_value_a1("A1", CellValue::String("Date".to_string()))
                .unwrap();
            sheet
                .set_value_a1("B1", CellValue::String("Amount".to_string()))
                .unwrap();

            let date_1 = ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;
            let date_2 = ymd_to_serial(ExcelDate::new(2024, 1, 16), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;

            sheet.set_value_a1("A2", CellValue::Number(date_1)).unwrap();
            sheet.set_value_a1("B2", CellValue::Number(10.0)).unwrap();
            sheet.set_value_a1("A3", CellValue::Number(date_2)).unwrap();
            sheet.set_value_a1("B3", CellValue::Number(20.0)).unwrap();
        }

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        let schema = wb
            .inner
            .get_pivot_schema_internal("Sheet1", "A1:B3", 10)
            .unwrap();

        let date_field = schema
            .fields
            .iter()
            .find(|f| f.name == "Date")
            .expect("expected Date field in schema");
        assert_eq!(date_field.field_type, pivot_engine::PivotFieldType::Date);
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_infers_dates_from_column_styles_when_cells_have_other_styles() {
        use std::io::Cursor;

        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1").unwrap();

        // Column has date number format.
        let date_style_id = workbook.styles.intern(formula_model::Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Default::default()
        });
        // Cells have an additional style layer (bold) that does not specify a number format.
        let bold_style_id = workbook.styles.intern(formula_model::Style {
            font: Some(Font {
                bold: true,
                ..Font::default()
            }),
            ..Default::default()
        });

        {
            let sheet = workbook.sheet_mut(sheet_id).unwrap();
            sheet.set_col_style_id(0, Some(date_style_id));

            sheet
                .set_value_a1("A1", CellValue::String("Date".to_string()))
                .unwrap();
            sheet
                .set_value_a1("B1", CellValue::String("Amount".to_string()))
                .unwrap();

            let date_1 = ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;
            let date_2 = ymd_to_serial(ExcelDate::new(2024, 1, 16), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;

            sheet.set_value_a1("A2", CellValue::Number(date_1)).unwrap();
            sheet.set_value_a1("B2", CellValue::Number(10.0)).unwrap();
            sheet.set_value_a1("A3", CellValue::Number(date_2)).unwrap();
            sheet.set_value_a1("B3", CellValue::Number(20.0)).unwrap();

            // Apply the bold style to the date column cells, without overriding the number format.
            sheet.set_style_id_a1("A2", bold_style_id).unwrap();
            sheet.set_style_id_a1("A3", bold_style_id).unwrap();
        }

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        let schema = wb
            .inner
            .get_pivot_schema_internal("Sheet1", "A1:B3", 10)
            .unwrap();

        let date_field = schema
            .fields
            .iter()
            .find(|f| f.name == "Date")
            .expect("expected Date field in schema");
        assert_eq!(date_field.field_type, pivot_engine::PivotFieldType::Date);
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn from_xlsx_bytes_imports_row_styles_for_pivot_date_inference() {
        use std::io::Cursor;

        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = formula_model::Workbook::new();
        let sheet_id = workbook.add_sheet("Sheet1").unwrap();

        let date_style_id = workbook.styles.intern(formula_model::Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Default::default()
        });

        {
            let sheet = workbook.sheet_mut(sheet_id).unwrap();
            // Apply the date number format via row defaults for the record rows.
            sheet.set_row_style_id(1, Some(date_style_id)); // row 2
            sheet.set_row_style_id(2, Some(date_style_id)); // row 3

            sheet
                .set_value_a1("A1", CellValue::String("Date".to_string()))
                .unwrap();
            sheet
                .set_value_a1("B1", CellValue::String("Amount".to_string()))
                .unwrap();

            let date_1 = ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;
            let date_2 = ymd_to_serial(ExcelDate::new(2024, 1, 16), ExcelDateSystem::EXCEL_1900)
                .unwrap() as f64;

            sheet.set_value_a1("A2", CellValue::Number(date_1)).unwrap();
            sheet.set_value_a1("B2", CellValue::Number(10.0)).unwrap();
            sheet.set_value_a1("A3", CellValue::Number(date_2)).unwrap();
            sheet.set_value_a1("B3", CellValue::Number(20.0)).unwrap();
        }

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
        let bytes = cursor.into_inner();

        let wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
        let schema = wb
            .inner
            .get_pivot_schema_internal("Sheet1", "A1:B3", 10)
            .unwrap();

        let date_field = schema
            .fields
            .iter()
            .find(|f| f.name == "Date")
            .expect("expected Date field in schema");
        assert_eq!(date_field.field_type, pivot_engine::PivotFieldType::Date);
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
    fn apply_operation_insert_rows_preserves_phonetic_metadata_on_formula_cells() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=\"漢字\""))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=PHONETIC(A1)"))
            .unwrap();
        wb.engine
            .set_cell_phonetic(DEFAULT_SHEET, "A1", Some("かんじ".to_string()))
            .unwrap();
        wb.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "B1"),
            EngineValue::Text("かんじ".to_string())
        );

        wb.apply_operation_internal(EditOpDto::InsertRows {
            sheet: DEFAULT_SHEET.to_string(),
            row: 0,
            count: 1,
        })
        .unwrap();

        wb.recalculate_internal(None).unwrap();
        assert_eq!(
            wb.engine.get_cell_phonetic(DEFAULT_SHEET, "A2"),
            Some("かんじ")
        );
        assert_eq!(
            wb.engine.get_cell_value(DEFAULT_SHEET, "B2"),
            EngineValue::Text("かんじ".to_string())
        );
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
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("A"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(10.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!("A"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B3", json!(5.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A4", json!("B"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B4", json!(7.0))
            .unwrap();

        // No formulas, but run a recalc to mirror typical usage where pivots reflect calculated
        // values.
        wb.recalculate_internal(None).unwrap();

        let config = formula_model::pivots::PivotConfig {
            row_fields: vec![formula_model::pivots::PivotField::new("Category")],
            column_fields: vec![],
            value_fields: vec![formula_model::pivots::ValueField {
                source_field: formula_model::pivots::PivotFieldRef::CacheFieldName(
                    "Amount".to_string(),
                ),
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

    #[test]
    fn calculate_pivot_writes_dates_as_serial_numbers_and_includes_date_number_format() {
        // Pivot source dates are represented in the worksheet as Excel serial numbers + a date
        // number format. Ensure `calculatePivot` emits the same underlying values and includes a
        // date number-format hint for the output label cell.
        let date = formula_engine::date::ExcelDate::new(1904, 1, 1);

        for system in [
            formula_engine::date::ExcelDateSystem::EXCEL_1900,
            formula_engine::date::ExcelDateSystem::Excel1904,
        ] {
            let mut wb = WorkbookState::new_with_default_sheet();
            wb.engine.set_date_system(system);

            wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("Date")).unwrap();
            wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("Sales")).unwrap();

            let serial = formula_engine::date::ymd_to_serial(
                date,
                system,
            )
            .unwrap() as f64;
            wb.set_cell_internal(DEFAULT_SHEET, "A2", json!(serial)).unwrap();
            wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(10.0)).unwrap();

            let date_style = wb.engine.intern_style(Style {
                number_format: Some("m/d/yyyy".to_string()),
                ..Style::default()
            });
            wb.engine
                .set_cell_style_id(DEFAULT_SHEET, "A2", date_style)
                .unwrap();

            wb.recalculate_internal(None).unwrap();

            let config = formula_model::pivots::PivotConfig {
                row_fields: vec![formula_model::pivots::PivotField::new("Date")],
                column_fields: vec![],
                value_fields: vec![formula_model::pivots::ValueField {
                    source_field: formula_model::pivots::PivotFieldRef::CacheFieldName("Sales".to_string()),
                    name: "Sum of Sales".to_string(),
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
                    rows: false,
                    columns: false,
                },
            };

            let engine_config = pivot_config_model_to_engine(&config);
            let writes = wb
                .calculate_pivot_writes_internal(DEFAULT_SHEET, "A1:B2", "D1", &engine_config)
                .unwrap();

            let date_write = writes
                .iter()
                .find(|w| w.address == "D2")
                .unwrap_or_else(|| panic!("missing date label write, got {writes:?}"));
            assert_eq!(date_write.value, json!(serial));
            assert_eq!(date_write.number_format.as_deref(), Some("m/d/yyyy"));
        }
    }

    #[test]
    fn calculate_pivot_includes_value_field_number_format_hints() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("Region")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("Sales")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("East")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(100.0)).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!("East")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B3", json!(150.0)).unwrap();

        wb.recalculate_internal(None).unwrap();

        let config = formula_model::pivots::PivotConfig {
            row_fields: vec![formula_model::pivots::PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![formula_model::pivots::ValueField {
                source_field: formula_model::pivots::PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: formula_model::pivots::AggregationType::Sum,
                number_format: Some("$#,##0.00".to_string()),
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
                rows: false,
                columns: false,
            },
        };

        let engine_config = pivot_config_model_to_engine(&config);
        let writes = wb
            .calculate_pivot_writes_internal(DEFAULT_SHEET, "A1:B3", "D1", &engine_config)
            .unwrap();

        let value_write = writes
            .iter()
            .find(|w| w.address == "E2")
            .unwrap_or_else(|| panic!("missing value write, got {writes:?}"));
        assert_eq!(value_write.value, json!(250.0));
        assert_eq!(value_write.number_format.as_deref(), Some("$#,##0.00"));
    }

    #[test]
    fn calculate_pivot_includes_default_percent_format_for_percent_show_as() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("Region")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("Sales")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("East")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(1.0)).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!("West")).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B3", json!(3.0)).unwrap();

        wb.recalculate_internal(None).unwrap();

        let config = formula_model::pivots::PivotConfig {
            row_fields: vec![formula_model::pivots::PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![formula_model::pivots::ValueField {
                source_field: formula_model::pivots::PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: formula_model::pivots::AggregationType::Sum,
                number_format: None,
                show_as: Some(formula_model::pivots::ShowAsType::PercentOfGrandTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: formula_model::pivots::Layout::Tabular,
            subtotals: formula_model::pivots::SubtotalPosition::None,
            grand_totals: formula_model::pivots::GrandTotals {
                rows: false,
                columns: false,
            },
        };

        let engine_config = pivot_config_model_to_engine(&config);
        let writes = wb
            .calculate_pivot_writes_internal(DEFAULT_SHEET, "A1:B3", "D1", &engine_config)
            .unwrap();

        let value_write = writes
            .iter()
            .find(|w| w.address == "E2")
            .unwrap_or_else(|| panic!("missing value write, got {writes:?}"));
        // 1 / (1 + 3) = 0.25
        assert_eq!(value_write.value, json!(0.25));
        assert_eq!(value_write.number_format.as_deref(), Some("0.00%"));
    }

    #[test]
    fn get_pivot_schema_reports_field_types_and_limits_samples() {
        let mut wb = WorkbookState::new_with_default_sheet();

        // Source data (headers + records).
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("Category"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("Amount"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A2", json!("A"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B2", json!(10.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A3", json!("A"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B3", json!(5.0))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "A4", json!("B"))
            .unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B4", json!(7.0))
            .unwrap();

        wb.recalculate_internal(None).unwrap();

        // Only sample the first two records.
        let schema = wb
            .get_pivot_schema_internal(DEFAULT_SHEET, "A1:B4", 2)
            .unwrap();

        assert_eq!(schema.record_count, 3);
        assert_eq!(schema.fields.len(), 2);

        assert_eq!(schema.fields[0].name, "Category");
        assert_eq!(
            schema.fields[0].field_type,
            pivot_engine::PivotFieldType::Text
        );
        assert_eq!(
            schema.fields[0].sample_values,
            vec![
                pivot_engine::PivotValue::Text("A".to_string()),
                pivot_engine::PivotValue::Text("A".to_string()),
            ]
        );

        assert_eq!(schema.fields[1].name, "Amount");
        assert_eq!(
            schema.fields[1].field_type,
            pivot_engine::PivotFieldType::Number
        );
        assert_eq!(
            schema.fields[1].sample_values,
            vec![
                pivot_engine::PivotValue::Number(10.0),
                pivot_engine::PivotValue::Number(5.0),
            ]
        );
    }

    #[test]
    fn goal_seek_converges_and_returns_changes() {
        use formula_engine::what_if::goal_seek::GoalSeekStatus;

        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!(1.0)).unwrap();
        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!("=A1*A1"))
            .unwrap();

        let (result, changes) = wb
            .goal_seek_internal(
                DEFAULT_SHEET,
                "B1",
                9.0,
                "A1",
                GoalSeekTuning::default(),
            )
            .unwrap();

        assert_eq!(result.status, GoalSeekStatus::Converged);
        assert!(
            (result.solution - 3.0).abs() < 1e-3,
            "expected solution near 3, got {result:?}"
        );

        let a1 = changes
            .iter()
            .find(|c| c.sheet == DEFAULT_SHEET && c.address == "A1")
            .expect("expected A1 change");
        let a1_val = a1
            .value
            .as_f64()
            .unwrap_or_else(|| panic!("expected numeric A1 value, got {:?}", a1.value));
        assert!((a1_val - 3.0).abs() < 1e-3);

        let b1 = changes
            .iter()
            .find(|c| c.sheet == DEFAULT_SHEET && c.address == "B1")
            .expect("expected B1 change");
        let b1_val = b1
            .value
            .as_f64()
            .unwrap_or_else(|| panic!("expected numeric B1 value, got {:?}", b1.value));
        assert!((b1_val - 9.0).abs() < 1e-3);
    }

    #[test]
    fn style_json_to_model_style_accepts_ui_camel_case_number_format() {
        let style = style_json_to_model_style(&json!({ "numberFormat": "0.00" }));
        assert_eq!(style.number_format.as_deref(), Some("0.00"));
    }

    #[test]
    fn style_json_to_model_style_treats_null_number_format_as_explicit_general_override() {
        let style = style_json_to_model_style(&json!({ "numberFormat": null }));
        assert_eq!(style.number_format.as_deref(), Some("General"));
    }

    #[test]
    fn style_json_to_model_style_prefers_null_number_format_over_imported_snake_case() {
        let style = style_json_to_model_style(&json!({
            "number_format": "0.00",
            "numberFormat": null,
        }));
        assert_eq!(style.number_format.as_deref(), Some("General"));
    }

    #[test]
    fn style_json_to_model_style_accepts_ui_protection_locked() {
        let style = style_json_to_model_style(&json!({ "protection": { "locked": false } }));
        assert_eq!(style.protection.as_ref().map(|p| p.locked), Some(false));
    }

    #[test]
    fn style_json_to_model_style_treats_null_locked_as_explicit_default_override() {
        // Explicit null clears a lower-precedence `locked=false` back to Excel's default locked=true.
        let style = style_json_to_model_style(&json!({ "locked": null }));
        assert_eq!(style.protection, Some(Protection::default()));
    }

    #[test]
    fn style_json_to_model_style_accepts_top_level_locked() {
        let style = style_json_to_model_style(&json!({ "locked": false }));
        assert_eq!(style.protection.as_ref().map(|p| p.locked), Some(false));
    }

    #[test]
    fn style_json_to_model_style_treats_null_alignment_horizontal_as_general_override() {
        let style = style_json_to_model_style(&json!({ "alignment": { "horizontal": null } }));
        assert_eq!(
            style
                .alignment
                .as_ref()
                .and_then(|a| a.horizontal)
                .unwrap_or(HorizontalAlignment::Left),
            HorizontalAlignment::General
        );
    }

    #[test]
    fn style_json_to_model_style_treats_nested_null_locked_as_explicit_default_override() {
        let style = style_json_to_model_style(&json!({ "protection": { "locked": null } }));
        assert_eq!(style.protection, Some(Protection::default()));
    }

    #[test]
    fn style_json_to_model_style_treats_null_hidden_as_explicit_default_override() {
        let style = style_json_to_model_style(&json!({ "hidden": null }));
        assert_eq!(style.protection, Some(Protection::default()));
    }

    #[test]
    fn style_json_to_model_style_prefers_full_model_style_fields_when_available() {
        // When the caller provides snake_case `Style`, preserve its fields while still honoring
        // UI-friendly overlay keys.
        let style = style_json_to_model_style(&json!({
            "font": { "bold": true },
            "number_format": "0.00",
        }));
        assert_eq!(style.number_format.as_deref(), Some("0.00"));
        assert_eq!(style.font.as_ref().map(|f| f.bold), Some(true));
    }

    #[test]
    fn intern_style_parses_font_color_rgb_object() {
        let style = style_json_to_model_style(&json!({
            "font": { "color": { "rgb": "FF112233" } }
        }));
        let mut engine = Engine::new();
        let style_id = engine.intern_style(style);
        let stored = engine
            .style_table()
            .get(style_id)
            .unwrap_or_else(|| panic!("expected style for id {style_id}"));

        assert_eq!(
            stored.font.as_ref().and_then(|f| f.color),
            Some(Color::new_argb(0xFF112233))
        );
    }

    #[test]
    fn style_json_explicit_null_number_format_clears_lower_precedence_layers() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

        // Sheet default style applies `0.00` (CELL("format") => "F2").
        let sheet_style = engine.intern_style(Style {
            number_format: Some("0.00".to_string()),
            ..Style::default()
        });
        engine.set_sheet_default_style_id("Sheet1", Some(sheet_style));

        // Column style explicitly clears back to General via `null`.
        let clear_style = style_json_to_model_style(&json!({ "numberFormat": null }));
        let clear_style_id = engine.intern_style(clear_style);
        assert_ne!(clear_style_id, 0);
        engine.set_col_style_id("Sheet1", 0, Some(clear_style_id));

        engine
            .set_cell_formula("Sheet1", "B1", "=CELL(\"format\",A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            EngineValue::Text("G".to_string())
        );
    }

    #[test]
    fn intern_style_parses_top_level_font_color_hex() {
        let style = style_json_to_model_style(&json!({
            "fontColor": "#112233"
        }));
        let mut engine = Engine::new();
        let style_id = engine.intern_style(style);
        let stored = engine
            .style_table()
            .get(style_id)
            .unwrap_or_else(|| panic!("expected style for id {style_id}"));

        assert_eq!(
            stored.font.as_ref().and_then(|f| f.color),
            Some(Color::new_argb(0xFF112233))
        );
    }

    #[test]
    fn style_json_explicit_null_alignment_horizontal_clears_lower_precedence_layers() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", "x").unwrap();

        let sheet_style = engine.intern_style(Style {
            alignment: Some(Alignment {
                horizontal: Some(HorizontalAlignment::Left),
                ..Alignment::default()
            }),
            ..Style::default()
        });
        engine.set_sheet_default_style_id("Sheet1", Some(sheet_style));

        let clear_style = style_json_to_model_style(&json!({ "alignment": { "horizontal": null } }));
        let clear_style_id = engine.intern_style(clear_style);
        assert_ne!(clear_style_id, 0);
        engine.set_col_style_id("Sheet1", 0, Some(clear_style_id));

        engine
            .set_cell_formula("Sheet1", "B1", "=CELL(\"prefix\",A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            EngineValue::Text(String::new())
        );
    }

    #[test]
    fn style_json_to_model_style_parses_theme_and_tint_font_color() {
        let style = style_json_to_model_style(&json!({
            "font": { "color": { "theme": 1, "tint": 0.5 } }
        }));
        assert_eq!(
            style.font.as_ref().and_then(|f| f.color),
            Some(Color::Theme {
                theme: 1,
                tint: Some(500),
            })
        );
    }

    #[test]
    fn style_json_explicit_null_protection_locked_clears_lower_precedence_layers() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

        let sheet_style = engine.intern_style(Style {
            protection: Some(Protection {
                locked: false,
                hidden: false,
            }),
            ..Style::default()
        });
        engine.set_sheet_default_style_id("Sheet1", Some(sheet_style));

        // Use the top-level `{ locked: null }` alias.
        let clear_style = style_json_to_model_style(&json!({ "locked": null }));
        let clear_style_id = engine.intern_style(clear_style);
        assert_ne!(clear_style_id, 0);
        engine.set_col_style_id("Sheet1", 0, Some(clear_style_id));

        engine
            .set_cell_formula("Sheet1", "B1", "=CELL(\"protect\",A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            EngineValue::Number(1.0)
        );
    }
}
