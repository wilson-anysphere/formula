/// Limits for payloads that cross the Tauri IPC boundary.
///
/// These are defensive: webview input should be treated as untrusted, and very large payloads can
/// lead to excessive memory usage and slow processing (e.g. workbook cloning, script parsing, or
/// spawning subprocesses).
use serde::de;
use serde::Deserialize;
use serde_json::Value as JsonValue;
use std::fmt;

/// Maximum size (in bytes) of a filesystem path string accepted over IPC.
///
/// Rationale: OS limits are typically far lower (e.g. 4KiB on many POSIX systems) and real-world
/// paths are usually a few hundred bytes at most. This cap is intentionally generous (8KiB) while
/// still bounding allocations from untrusted WebView input.
pub const MAX_IPC_PATH_BYTES: usize = 8_192; // 8 KiB

/// Maximum size (in bytes) of a URL string accepted over IPC.
///
/// Rationale: URLs used by the app (marketplace, OAuth, external links) are expected to be small.
/// 8KiB is comfortably above common practical limits while preventing a compromised WebView from
/// forcing the backend to allocate and parse multi-megabyte "URLs".
pub const MAX_IPC_URL_BYTES: usize = 8_192; // 8 KiB

/// Maximum size (in bytes) of the OAuth loopback redirect URI accepted over IPC.
///
/// This is kept separate from `MAX_IPC_URL_BYTES` so the limit can be tightened independently if
/// needed in the future.
pub const MAX_OAUTH_REDIRECT_URI_BYTES: usize = MAX_IPC_URL_BYTES;
// Guardrail: the OAuth redirect URI limit must never exceed the global IPC URL limit.
const _: () = assert!(MAX_OAUTH_REDIRECT_URI_BYTES <= MAX_IPC_URL_BYTES);

/// Maximum size (in bytes) of a system notification title accepted over IPC.
///
/// Rationale: notification titles are short UI strings; allowing unbounded payloads provides an
/// easy memory/CPU DoS vector against the privileged backend.
pub const MAX_NOTIFICATION_TITLE_BYTES: usize = 256;

/// Maximum size (in bytes) of a system notification body accepted over IPC.
///
/// Rationale: notification bodies should still be human-readable snippets. 4KiB is large enough
/// for multi-line error messages while keeping worst-case allocations bounded.
pub const MAX_NOTIFICATION_BODY_BYTES: usize = 4_096; // 4 KiB

/// Maximum size (in bytes) of a script `code` payload accepted over IPC.
pub const MAX_SCRIPT_CODE_BYTES: usize = 1_000_000; // ~1MB

/// Maximum number of sheet IDs accepted by the `reorder_sheets` IPC command.
///
/// The frontend sends the full sheet ID ordering, so this must be high enough for very large
/// workbooks while still preventing unbounded allocations during deserialization.
pub const MAX_REORDER_SHEET_IDS: usize = 10_000;

/// Maximum size (in bytes) of a single sheet ID accepted over IPC.
///
/// Sheet IDs are typically UUID strings (36 bytes), so this is intentionally conservative.
pub const MAX_SHEET_ID_BYTES: usize = 128;

/// Maximum number of print-area ranges accepted by the `set_sheet_print_area` IPC command.
///
/// Print areas are typically a small number of disjoint ranges; this cap prevents allocating large
/// vectors when deserializing untrusted IPC inputs.
pub const MAX_PRINT_AREA_RANGES: usize = 1_000;

/// Maximum size (in bytes) of keys used to index OS-keychain-backed encrypted stores over IPC.
///
/// These keys are used as document IDs inside encrypted JSON blobs (tokens, encryption keys,
/// refresh state, etc). They should be short (typically an identifier or URL-ish string), but a
/// compromised WebView could otherwise send arbitrarily large values and force large allocations or
/// create pathological on-disk store entries.
pub const MAX_IPC_SECURE_STORE_KEY_BYTES: usize = 1_024; // 1 KiB

/// Maximum size (in bytes) of collaboration token strings accepted over IPC.
///
/// Tokens are secrets and should be relatively small (e.g. JWTs). This cap prevents a compromised
/// WebView from persisting multi-megabyte "tokens" into the encrypted store (memory + disk DoS).
pub const MAX_IPC_COLLAB_TOKEN_BYTES: usize = 64 * 1024; // 64 KiB

/// Maximum size (in bytes) of a base64-encoded collaboration cell encryption key accepted over IPC.
///
/// Cell encryption keys are fixed-size (32 bytes) so the expected base64 string is tiny (~44 bytes).
/// This cap is intentionally generous while preventing untrusted IPC inputs from forcing large
/// allocations during base64 decode.
pub const MAX_IPC_COLLAB_ENCRYPTION_KEY_BASE64_BYTES: usize = 256;

/// Maximum size (in bytes) of the tray-status string accepted over IPC.
///
/// `set_tray_status` only supports a few short status tokens; this cap prevents a compromised
/// WebView from sending arbitrarily large strings even though the backend will reject unknown
/// values.
pub const MAX_IPC_TRAY_STATUS_BYTES: usize = 32;

/// Maximum size (in bytes) of system notification titles accepted over IPC.
///
/// Rationale: notification titles are short UI strings; anything larger is likely unintended or
/// malicious.
pub const MAX_IPC_NOTIFICATION_TITLE_BYTES: usize = MAX_NOTIFICATION_TITLE_BYTES;

/// Maximum size (in bytes) of system notification bodies accepted over IPC.
///
/// Rationale: notification bodies can be longer than titles, but should still be bounded to avoid
/// untrusted IPC inputs allocating excessive memory.
pub const MAX_IPC_NOTIFICATION_BODY_BYTES: usize = MAX_NOTIFICATION_BODY_BYTES;

/// A `String` wrapper that enforces a maximum UTF-8 byte length during deserialization.
///
/// This is intended for high-risk IPC command arguments (paths/URLs/notification strings) so
/// oversized payloads fail fast during deserialization, before any further parsing or processing.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct LimitedString<const MAX_BYTES: usize>(String);

impl<const MAX_BYTES: usize> LimitedString<MAX_BYTES> {
    pub fn as_str(&self) -> &str {
        &self.0
    }

    pub fn into_string(self) -> String {
        self.0
    }
}

impl<const MAX_BYTES: usize> std::ops::Deref for LimitedString<MAX_BYTES> {
    type Target = str;

    fn deref(&self) -> &Self::Target {
        self.as_str()
    }
}

impl<const MAX_BYTES: usize> AsRef<str> for LimitedString<MAX_BYTES> {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl<const MAX_BYTES: usize> From<LimitedString<MAX_BYTES>> for String {
    fn from(value: LimitedString<MAX_BYTES>) -> Self {
        value.0
    }
}

impl<'de, const MAX_BYTES: usize> Deserialize<'de> for LimitedString<MAX_BYTES> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct LimitedStringVisitor<const MAX_BYTES: usize>;

        impl<'de, const MAX_BYTES: usize> de::Visitor<'de> for LimitedStringVisitor<MAX_BYTES> {
            type Value = LimitedString<MAX_BYTES>;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("a string")
            }

            fn visit_str<E>(self, value: &str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                if value.len() > MAX_BYTES {
                    return Err(E::custom(format!(
                        "String is too large (max {MAX_BYTES} bytes)"
                    )));
                }
                Ok(LimitedString(value.to_owned()))
            }

            fn visit_borrowed_str<E>(self, value: &'de str) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                self.visit_str(value)
            }

            fn visit_string<E>(self, value: String) -> Result<Self::Value, E>
            where
                E: de::Error,
            {
                if value.len() > MAX_BYTES {
                    return Err(E::custom(format!(
                        "String is too large (max {MAX_BYTES} bytes)"
                    )));
                }
                Ok(LimitedString(value))
            }
        }

        deserializer.deserialize_str(LimitedStringVisitor::<MAX_BYTES>)
    }
}

pub fn enforce_script_code_size(code: &str) -> Result<(), String> {
    if code.len() > MAX_SCRIPT_CODE_BYTES {
        return Err(format!(
            "Script is too large (max {MAX_SCRIPT_CODE_BYTES} bytes)"
        ));
    }
    Ok(())
}

/// Maximum size (in bytes) of the SQL statement text accepted over IPC.
///
/// Rationale: SQL text is typically small (often < 10KiB). Allowing unbounded statement strings
/// enables a compromised webview to force large allocations and slow parsing even before any
/// database call is made. 1 MiB is intentionally generous for complex queries while still keeping
/// worst-case input bounded.
pub const MAX_SQL_QUERY_TEXT_BYTES: usize = 1024 * 1024; // 1 MiB

/// Maximum number of bound parameters accepted over IPC for a single SQL query.
///
/// Rationale: parameter arrays are an easy way to create large inputs (allocation + per-param
/// processing). This is conservative while still allowing common uses like `IN (...)` queries.
pub const MAX_SQL_QUERY_PARAMS: usize = 1_000;

/// Maximum size (in bytes) of any single SQL parameter when serialized to JSON.
///
/// Rationale: large JSON values are expensive to parse/serialize and may be converted to strings
/// for binding; this bounds per-parameter allocations.
pub const MAX_SQL_QUERY_PARAM_BYTES: usize = 64 * 1024; // 64 KiB

/// Maximum size (in bytes) of the SQL `credentials` payload when serialized to JSON.
///
/// Rationale: credentials should be small (user/password, token, etc). Bounding prevents a
/// compromised webview from sending huge objects.
pub const MAX_SQL_QUERY_CREDENTIALS_BYTES: usize = 64 * 1024; // 64 KiB

/// Maximum size (in bytes) of the SQL `connection` descriptor when serialized to JSON.
///
/// Rationale: connection descriptors are small (kind/host/path/etc). Bounding prevents oversized
/// connection objects from consuming memory and slowing down backend processing.
pub const MAX_SQL_QUERY_CONNECTION_BYTES: usize = 64 * 1024; // 64 KiB

/// Enforce an upper bound on the size of a JSON value when serialized as UTF-8 bytes.
///
/// This performs a bounded streaming serialization (via `serde_json::to_writer`) to avoid
/// allocating a `Vec<u8>` for oversized/untrusted payloads.
///
/// Errors are deterministic and mention which limit was exceeded so they can be surfaced directly
/// over IPC.
pub fn enforce_json_byte_size(
    value: &JsonValue,
    max_bytes: usize,
    value_name: &str,
    limit_name: &'static str,
) -> Result<(), String> {
    struct ByteLimitWriter {
        remaining: usize,
    }

    impl std::io::Write for ByteLimitWriter {
        fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
            if buf.len() > self.remaining {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Other,
                    "payload too large",
                ));
            }
            self.remaining -= buf.len();
            Ok(buf.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    let mut writer = ByteLimitWriter {
        remaining: max_bytes,
    };
    match serde_json::to_writer(&mut writer, value) {
        Ok(()) => Ok(()),
        Err(err) if err.is_io() => Err(format!(
            "{value_name} exceeds {limit_name} ({max_bytes} bytes)"
        )),
        Err(_) => Err(format!("Failed to serialize {value_name} as JSON")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn enforce_script_code_size_rejects_oversized_payloads() {
        let oversized = "x".repeat(MAX_SCRIPT_CODE_BYTES + 1);
        let err = enforce_script_code_size(&oversized).expect_err("expected size check to fail");
        assert!(
            err.contains("Script is too large"),
            "unexpected error message: {err}"
        );
        assert!(
            err.contains(&MAX_SCRIPT_CODE_BYTES.to_string()),
            "expected error message to mention limit: {err}"
        );
    }

    #[test]
    fn enforce_json_byte_size_rejects_oversized_payloads() {
        // JSON string serialization adds 2 bytes for the surrounding quotes.
        let ok = JsonValue::String("a".repeat(8));
        assert!(enforce_json_byte_size(&ok, 10, "value", "MAX").is_ok());

        let too_big = JsonValue::String("a".repeat(9));
        let err = enforce_json_byte_size(&too_big, 10, "value", "MAX").unwrap_err();
        assert!(
            err.contains("exceeds MAX"),
            "expected error to mention the limit name, got: {err}"
        );
    }

    #[test]
    fn limited_string_rejects_oversized_payloads_during_deserialization() {
        let oversized = "x".repeat(MAX_IPC_PATH_BYTES + 1);
        let json = format!("\"{oversized}\"");
        let err = serde_json::from_str::<LimitedString<MAX_IPC_PATH_BYTES>>(&json)
            .expect_err("expected LimitedString deserialization to fail");
        assert!(
            err.to_string().contains(&MAX_IPC_PATH_BYTES.to_string()),
            "expected error to mention limit, got: {err}"
        );
    }

    #[test]
    fn limited_string_allows_payloads_at_or_under_the_limit() {
        let ok = "x".repeat(MAX_IPC_PATH_BYTES);
        let json = format!("\"{ok}\"");
        let parsed = serde_json::from_str::<LimitedString<MAX_IPC_PATH_BYTES>>(&json)
            .expect("expected LimitedString to deserialize");
        assert_eq!(parsed.as_str(), ok);
    }

    #[test]
    fn source_guardrail_commands_use_limited_string_for_paths() {
        let src = include_str!("commands.rs");
        let start = src
            .find("pub async fn open_workbook")
            .expect("expected commands.rs to define open_workbook");
        let end = (start + 300).min(src.len());
        let snippet = &src[start..end];
        assert!(
            snippet.contains("path: LimitedString<MAX_IPC_PATH_BYTES>"),
            "expected `open_workbook` to use `LimitedString<MAX_IPC_PATH_BYTES>` for `path`"
        );
    }

    #[test]
    fn source_guardrail_commands_use_limited_string_for_network_fetch_url() {
        let src = include_str!("commands.rs");
        let start = src
            .find("pub async fn network_fetch")
            .expect("expected commands.rs to define network_fetch");
        let end = (start + 200).min(src.len());
        let snippet = &src[start..end];
        assert!(
            snippet.contains("url: LimitedString<MAX_IPC_URL_BYTES>"),
            "expected `network_fetch` to use `LimitedString<MAX_IPC_URL_BYTES>` for `url`"
        );
    }

    #[test]
    fn source_guardrail_main_use_limited_string_for_oauth_redirect_uri() {
        let src = include_str!("main.rs");
        assert!(
            src.contains("redirect_uri: LimitedString<MAX_OAUTH_REDIRECT_URI_BYTES>"),
            "expected `oauth_loopback_listen` to use `LimitedString<MAX_OAUTH_REDIRECT_URI_BYTES>` for `redirect_uri`"
        );
    }
}
