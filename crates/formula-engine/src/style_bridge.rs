use formula_model::{Alignment, HorizontalAlignment, Protection, Style};
use serde_json::Value as JsonValue;

/// Best-effort conversion from UI formatting JSON (typically camelCase) into a `formula_model::Style`.
///
/// UI-facing style objects (e.g. document/controller style tables, `apply_sheet_formatting_deltas`
/// payloads) tend to use JS-friendly keys like `numberFormat`. The engine and storage layers use
/// `formula_model::Style`, which is serde-deserialized from a Rust-friendly schema (`number_format`,
/// typed enums, etc).
///
/// This helper is intentionally **best-effort**:
/// - Unknown keys are ignored.
/// - Unexpected shapes/types are ignored.
/// - The function never panics.
///
/// Currently mapped (minimum required for `CELL`/`INFO` style metadata):
/// - Number format: `numberFormat` or `number_format` -> `Style.number_format`.
/// - Protection: `protection.locked` or top-level `locked` -> `Style.protection.locked`.
/// - Alignment: `alignment.horizontal` -> `Style.alignment.horizontal` (general/left/center/right/fill/justify).
pub fn ui_style_to_model_style(value: &JsonValue) -> Style {
    let Some(obj) = value.as_object() else {
        return Style::default();
    };

    let mut out = Style::default();

    // --- number_format ---
    if let Some(fmt) = obj
        .get("numberFormat")
        .or_else(|| obj.get("number_format"))
        .and_then(|v| v.as_str())
        .map(str::trim)
        .filter(|s| !s.is_empty())
    {
        out.number_format = Some(fmt.to_string());
    }

    // --- protection ---
    let (locked, hidden) = {
        let protection = obj.get("protection").and_then(|v| v.as_object());
        let locked = protection
            .and_then(|p| p.get("locked"))
            .or_else(|| obj.get("locked"))
            .and_then(|v| v.as_bool());
        let hidden = protection
            .and_then(|p| p.get("hidden"))
            .and_then(|v| v.as_bool());
        (locked, hidden)
    };
    if locked.is_some() || hidden.is_some() {
        out.protection = Some(Protection {
            // `Protection.locked` defaults to `true`.
            locked: locked.unwrap_or(true),
            hidden: hidden.unwrap_or(false),
        });
    }

    // --- alignment ---
    if let Some(alignment) = obj.get("alignment").and_then(|v| v.as_object()) {
        if let Some(horizontal_value) = alignment.get("horizontal") {
            // The UI model uses `null` to represent an explicit clear (fall back to General),
            // whereas missing keys mean "leave unchanged".
            let horizontal = if horizontal_value.is_null() {
                Some(HorizontalAlignment::General)
            } else {
                horizontal_value
                    .as_str()
                    .and_then(parse_horizontal_alignment)
            };
            if horizontal.is_some() {
                out.alignment = Some(Alignment {
                    horizontal,
                    ..Default::default()
                });
            }
        }
    }

    out
}

fn parse_horizontal_alignment(raw: &str) -> Option<HorizontalAlignment> {
    let raw = raw.trim();
    if raw.eq_ignore_ascii_case("general") {
        return Some(HorizontalAlignment::General);
    }
    if raw.eq_ignore_ascii_case("left") {
        return Some(HorizontalAlignment::Left);
    }
    if raw.eq_ignore_ascii_case("center") || raw.eq_ignore_ascii_case("centre") {
        return Some(HorizontalAlignment::Center);
    }
    if raw.eq_ignore_ascii_case("right") {
        return Some(HorizontalAlignment::Right);
    }
    if raw.eq_ignore_ascii_case("fill") {
        return Some(HorizontalAlignment::Fill);
    }
    if raw.eq_ignore_ascii_case("justify") {
        return Some(HorizontalAlignment::Justify);
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn ui_style_number_format_camel_case_maps_to_model_style() {
        let style = ui_style_to_model_style(&json!({ "numberFormat": "0.00" }));
        assert_eq!(style.number_format.as_deref(), Some("0.00"));
    }

    #[test]
    fn ui_style_number_format_snake_case_maps_to_model_style() {
        let style = ui_style_to_model_style(&json!({ "number_format": "0.00" }));
        assert_eq!(style.number_format.as_deref(), Some("0.00"));
    }

    #[test]
    fn ui_style_protection_locked_maps_to_model_style() {
        let style = ui_style_to_model_style(&json!({ "protection": { "locked": false } }));
        assert_eq!(style.protection.as_ref().map(|p| p.locked), Some(false));
    }

    #[test]
    fn ui_style_top_level_locked_maps_to_model_style() {
        let style = ui_style_to_model_style(&json!({ "locked": false }));
        assert_eq!(style.protection.as_ref().map(|p| p.locked), Some(false));
    }

    #[test]
    fn ui_style_alignment_horizontal_maps_to_model_style() {
        let style = ui_style_to_model_style(&json!({ "alignment": { "horizontal": "left" } }));
        assert_eq!(
            style.alignment.as_ref().and_then(|a| a.horizontal).as_ref(),
            Some(&HorizontalAlignment::Left)
        );
    }

    #[test]
    fn ui_style_alignment_horizontal_null_maps_to_general() {
        let style = ui_style_to_model_style(&json!({ "alignment": { "horizontal": null } }));
        assert_eq!(
            style.alignment.as_ref().and_then(|a| a.horizontal).as_ref(),
            Some(&HorizontalAlignment::General)
        );
    }

    #[test]
    fn ui_style_ignores_unexpected_shapes_and_preserves_mapped_keys() {
        // This is a regression test: naïvely deserializing into `formula_model::Style` would fail
        // on invalid color shapes like "#RRGGBB". We should still extract mapped fields.
        let style = ui_style_to_model_style(&json!({
            "font": { "color": "#FF0000" },
            "numberFormat": "0.00",
        }));
        assert_eq!(style.number_format.as_deref(), Some("0.00"));
        assert!(style.font.is_none(), "font should be ignored");
    }
}
