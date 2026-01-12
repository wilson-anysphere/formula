//! Backend-enforced resource limits for APIs that touch local resources.
//!
//! These guards exist to prevent a compromised webview (or accidental usage on very large
//! directories/files) from exhausting memory/CPU in the desktop process.

/// Maximum number of entries returned by the `list_dir` command.
///
/// The limit is enforced even when `recursive=false`, since a single directory can contain
/// an unbounded number of files.
pub const MAX_LIST_DIR_ENTRIES: usize = 50_000;

/// Maximum recursion depth for the `list_dir` command when `recursive=true`.
///
/// Depth starts at `0` for the root directory passed to `list_dir`.
pub const MAX_LIST_DIR_DEPTH: usize = 20;

