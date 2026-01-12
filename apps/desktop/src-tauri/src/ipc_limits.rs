/// Limits for payloads that cross the Tauri IPC boundary.
///
/// These are defensive: webview input should be treated as untrusted, and very large payloads can
/// lead to excessive memory usage and slow processing (e.g. workbook cloning, script parsing, or
/// spawning subprocesses).

/// Maximum size (in bytes) of a script `code` payload accepted over IPC.
pub const MAX_SCRIPT_CODE_BYTES: usize = 1_000_000; // ~1MB

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

