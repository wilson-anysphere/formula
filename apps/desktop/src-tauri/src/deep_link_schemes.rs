use serde_json::Value;
use std::sync::OnceLock;

#[derive(Debug)]
struct DeepLinkSchemesConfig {
    /// Normalized scheme names (lowercase, no trailing :/).
    schemes: Vec<String>,
    /// Normalized scheme prefixes including the trailing ':' (e.g. "formula:").
    scheme_prefixes: Vec<String>,
}

fn normalize_scheme(raw: &str) -> String {
    raw.trim()
        .to_ascii_lowercase()
        .trim_end_matches(|c| c == ':' || c == '/')
        .to_string()
}

fn is_reserved_scheme(normalized: &str) -> bool {
    // These schemes are not valid deep-link candidates for our desktop app. Treating them as
    // deep-link schemes would break other pipelines:
    // - `http`/`https`: would bypass the OAuth loopback security checks and accept arbitrary URLs.
    // - `file`: would cause file-open URLs to be misclassified as OAuth redirects.
    matches!(normalized, "http" | "https" | "file")
}

fn normalize_and_validate_scheme(raw: &str) -> Option<String> {
    let normalized = normalize_scheme(raw);
    if normalized.is_empty() {
        return None;
    }
    if normalized.contains(':') || normalized.contains('/') {
        return None;
    }
    if is_reserved_scheme(&normalized) {
        return None;
    }
    Some(normalized)
}

fn schemes_from_tauri_config() -> Vec<String> {
    let config_str = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/tauri.conf.json"));
    let config: Value = match serde_json::from_str(config_str) {
        Ok(v) => v,
        Err(_) => return vec!["formula".to_string()],
    };

    let desktop = config
        .get("plugins")
        .and_then(|p| p.get("deep-link"))
        .and_then(|p| p.get("desktop"));

    let mut out = Vec::<String>::new();

    let mut add_protocol = |protocol: &Value| {
        let Some(schemes) = protocol.get("schemes") else {
            return;
        };
        match schemes {
            Value::String(s) => {
                if let Some(normalized) = normalize_and_validate_scheme(s) {
                    out.push(normalized);
                }
            }
            Value::Array(items) => {
                for item in items {
                    let Some(s) = item.as_str() else { continue };
                    if let Some(normalized) = normalize_and_validate_scheme(s) {
                        out.push(normalized);
                    }
                }
            }
            _ => {}
        }
    };

    match desktop {
        Some(Value::Array(protocols)) => {
            for protocol in protocols {
                if protocol.is_object() {
                    add_protocol(protocol);
                }
            }
        }
        Some(protocol) if protocol.is_object() => {
            add_protocol(protocol);
        }
        _ => {}
    }

    out.sort();
    out.dedup();

    // The repo requires the stable `formula://` scheme for deep links / OAuth redirects.
    // Keep it present even if the config is missing or empty so runtime behavior remains stable.
    if !out.iter().any(|s| s == "formula") {
        out.push("formula".to_string());
    }
    out.sort();

    out
}

static CONFIG: OnceLock<DeepLinkSchemesConfig> = OnceLock::new();

fn config() -> &'static DeepLinkSchemesConfig {
    CONFIG.get_or_init(|| {
        let schemes = schemes_from_tauri_config();
        let scheme_prefixes = schemes.iter().map(|s| format!("{s}:")).collect();
        DeepLinkSchemesConfig {
            schemes,
            scheme_prefixes,
        }
    })
}

/// Returns the configured desktop deep-link schemes (normalized).
///
/// Source of truth: `apps/desktop/src-tauri/tauri.conf.json` → `plugins.deep-link.desktop.schemes`.
pub fn configured_schemes() -> &'static [String] {
    config().schemes.as_slice()
}

/// Returns true when `value` appears to be a deep-link URL for one of the configured schemes.
pub fn is_deep_link_url(value: &str) -> bool {
    let trimmed = value.trim().trim_matches('"');
    if trimmed.is_empty() {
        return false;
    }

    for prefix in &config().scheme_prefixes {
        let Some(candidate) = trimmed.get(..prefix.len()) else {
            continue;
        };
        if candidate.eq_ignore_ascii_case(prefix) {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_scheme_strips_trailing_delimiters() {
        assert_eq!(normalize_scheme("formula://"), "formula");
        assert_eq!(normalize_scheme(" formula: "), "formula");
        assert_eq!(normalize_scheme("formula/"), "formula");
        assert_eq!(normalize_scheme("FORMULA://"), "formula");
    }

    #[test]
    fn normalize_and_validate_scheme_filters_invalid_and_reserved_schemes() {
        assert_eq!(normalize_and_validate_scheme("formula://evil"), None);
        assert_eq!(normalize_and_validate_scheme("http"), None);
        assert_eq!(normalize_and_validate_scheme("https:"), None);
        assert_eq!(normalize_and_validate_scheme("file://"), None);
        assert_eq!(
            normalize_and_validate_scheme("Formula-Extra://"),
            Some("formula-extra".to_string())
        );
    }

    #[test]
    fn deep_link_url_detection_matches_normalized_scheme_prefixes() {
        // This repo always configures the `formula` scheme; assert we detect a typical URL.
        assert!(is_deep_link_url("formula://oauth/callback?code=123"));
        assert!(is_deep_link_url("FORMULA://oauth/callback?code=123"));
        assert!(is_deep_link_url("\"formula://oauth/callback?code=123\""));
        assert!(!is_deep_link_url("file:///tmp/book.xlsx"));
        assert!(!is_deep_link_url("not-a-url"));
    }
}
