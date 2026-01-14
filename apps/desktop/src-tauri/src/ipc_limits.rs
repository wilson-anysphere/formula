/// Limits for payloads that cross the Tauri IPC boundary.
///
/// These are defensive: webview input should be treated as untrusted, and very large payloads can
/// lead to excessive memory usage and slow processing (e.g. workbook cloning, script parsing, or
/// spawning subprocesses).

use serde_json::Value as JsonValue;

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
}
