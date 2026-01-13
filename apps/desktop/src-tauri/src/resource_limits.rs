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
