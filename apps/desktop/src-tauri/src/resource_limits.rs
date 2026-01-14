//! Backend-enforced resource limits for APIs that accept untrusted inputs.
//!
//! The desktop backend accepts requests from a WebView. If that WebView is compromised it can
//! send arbitrarily large ranges/exports, or access very large directories/files. These
//! conservative limits ensure the backend fails fast and deterministically instead of exhausting
//! memory/CPU (OOM/freeze).

/// Maximum number of filesystem entries processed by the `list_dir` command.
///
/// The limit is enforced even when `recursive=false`, since a single directory can contain an
/// unbounded number of files.
///
/// Note: `list_dir` returns file entries only, but this limit counts all directory entries
/// encountered during traversal (files + directories). This ensures CPU usage is bounded even for
/// directory trees that contain very few files.
pub const MAX_LIST_DIR_ENTRIES: usize = 50_000;

/// Maximum recursion depth for the `list_dir` command when `recursive=true`.
///
/// Depth starts at `0` for the root directory passed to `list_dir`.
pub const MAX_LIST_DIR_DEPTH: usize = 20;

/// Maximum number of cells that `get_range`/`set_range` will process in a single call.
pub const MAX_RANGE_CELLS_PER_CALL: usize = 1_000_000;

/// Maximum number of rows or columns allowed in a single `get_range`/`set_range` call.
///
/// This prevents pathological "skinny" ranges (e.g. 1×1_000_000) which can still be expensive even
/// when the total cell count is capped.
pub const MAX_RANGE_DIM: usize = 10_000;

/// Maximum number of cells that a single PDF export is allowed to render.
///
/// PDF generation is substantially more expensive than a plain range read, so this limit is
/// intentionally more conservative than `MAX_RANGE_CELLS_PER_CALL`.
pub const MAX_PDF_CELLS_PER_CALL: usize = 250_000;

/// Maximum size (in bytes) of a generated PDF before it is rejected.
///
/// Note: this is checked after generation, so it primarily protects against extremely large
/// base64 responses rather than CPU time.
pub const MAX_PDF_BYTES: usize = 20 * 1024 * 1024; // 20 MiB

/// Maximum size (in bytes) of a marketplace extension package download.
///
/// The desktop backend base64-encodes the downloaded package for IPC, which expands the payload by
/// ~33%. This limit is applied to the *raw* bytes to keep memory usage bounded even if the
/// marketplace server (or a network attacker) attempts to return an arbitrarily large response.
pub const MAX_MARKETPLACE_PACKAGE_BYTES: usize = 20 * 1024 * 1024; // 20 MiB

/// Maximum size (in bytes) of individual marketplace metadata headers (e.g. signatures/hashes).
///
/// Headers are expected to be small (an ed25519 signature is 64 bytes, SHA-256 is 32 bytes), but
/// the backend still treats them as untrusted input and enforces a cap to avoid large string
/// allocations.
pub const MAX_MARKETPLACE_HEADER_BYTES: usize = 4 * 1024; // 4 KiB

/// Maximum size (in bytes) of the XLSX/XLSM package we will retain in
/// [`crate::file_io::Workbook::origin_xlsx_bytes`].
///
/// Retaining the original package enables patch-based saves (we can patch edited worksheet XML
/// onto the original ZIP). However, keeping arbitrarily large packages in memory is a footgun:
/// a compromised webview can OOM the Rust backend by opening a huge workbook.
///
/// When a workbook exceeds this limit we still open it, but we *do not* keep the raw bytes.
/// Subsequent saves fall back to the regeneration-based code path.
pub const MAX_ORIGIN_XLSX_BYTES: usize = 64 * 1024 * 1024; // 64 MiB

/// Maximum number of bytes allowed for `Workbook::origin_xlsx_bytes`.
///
/// This reads the `FORMULA_MAX_ORIGIN_XLSX_BYTES` environment variable.
/// - In debug builds we allow any value (so developers can test both branches).
/// - In release builds we only honor stricter values (tightening); relaxing the limit is ignored.
pub fn max_origin_xlsx_bytes() -> usize {
    let default = MAX_ORIGIN_XLSX_BYTES;
    let Some(value) = std::env::var("FORMULA_MAX_ORIGIN_XLSX_BYTES")
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
    else {
        return default;
    };
    if cfg!(debug_assertions) {
        value
    } else {
        value.min(default)
    }
}

/// Maximum number of log lines captured from a single VBA macro execution.
///
/// VBA code can emit host output via `Debug.Print`/`MsgBox`. The desktop backend forwards this
/// output to the WebView over IPC. This limit ensures a compromised workbook or WebView cannot
/// force the backend to accumulate an unbounded `Vec<String>`.
pub const MAX_MACRO_OUTPUT_LINES: usize = 1_000;

/// Maximum total UTF-8 byte size of log output captured from a single VBA macro execution.
///
/// This is a coarse limit over the sum of bytes of all captured log lines (before IPC framing).
/// Together with `MAX_MACRO_OUTPUT_LINES` it bounds both memory usage and the size of the IPC
/// response payload.
pub const MAX_MACRO_OUTPUT_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

/// Maximum number of cell updates a single VBA macro execution is allowed to generate.
///
/// Macros can write to many cells and trigger large recalculation fanout (dependent formulas),
/// producing large `Vec<CellUpdateData>` payloads. This limit bounds memory usage and prevents
/// sending extremely large IPC responses. When exceeded, macro execution is aborted with an
/// explicit runtime error (no truncation) so callers can handle the failure deterministically.
pub const MAX_MACRO_UPDATES: usize = 10_000;

// -----------------------------------------------------------------------------
// Native Python runner limits
// -----------------------------------------------------------------------------

/// Maximum number of bytes captured from the native Python runner's stderr stream.
///
/// The WebView can execute arbitrary code in the Python subprocess. A compromised/buggy script can
/// spam stderr indefinitely; capturing stderr without a cap can OOM the Rust host.
///
/// 1 MiB keeps error messages and tracebacks useful while bounding host memory.
pub const MAX_PYTHON_STDERR_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

/// Maximum size (in bytes) of a single JSON protocol line emitted by the native Python runner.
///
/// The host reads newline-delimited JSON messages from the runner. Each line is bounded to prevent
/// pathological allocations (e.g. a single multi-gigabyte "line") in the Rust process.
pub const MAX_PYTHON_PROTOCOL_LINE_BYTES: usize = 1 * 1024 * 1024; // 1 MiB

/// Maximum number of distinct cell updates accumulated during a single Python execution.
///
/// Updates are deduped by `(sheet_id, row, col)` before being returned to the frontend, but without
/// an overall cap a script can still touch an unbounded number of cells and force the backend to
/// retain (and eventually serialize) a massive update payload.
///
/// This is aligned with `MAX_RANGE_CELLS_PER_CALL` so a single "large range" operation doesn't
/// immediately fail, but scripts that fan out to multiple large operations will be rejected.
pub const MAX_PYTHON_UPDATES: usize = MAX_RANGE_CELLS_PER_CALL;
