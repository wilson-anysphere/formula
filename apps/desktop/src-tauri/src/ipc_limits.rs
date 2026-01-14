/// Limits for payloads that cross the Tauri IPC boundary.
///
/// These are defensive: webview input should be treated as untrusted, and very large payloads can
/// lead to excessive memory usage and slow processing (e.g. workbook cloning, script parsing, or
/// spawning subprocesses).

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
}
