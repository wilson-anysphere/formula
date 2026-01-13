//! Guardrails for desktop filesystem scope enforcement in IPC commands.
//!
//! These are **source-level** tests: we scan `src/commands.rs` as text instead of compiling the
//! full Tauri shell (which would require the `desktop` feature and the system WebView toolchain on
//! Linux).
//!
//! Security goal:
//! - All IPC commands that accept a filesystem path must *canonicalize* the path (resolving `..`
//!   and symlinks) and enforce that the resolved path stays within the desktop filesystem scope:
//!   **HOME / DOCUMENTS / DOWNLOADS**.
//!
//! The canonicalization + scope policy lives in `src/fs_scope.rs`.

/// IPC commands that accept a filesystem path and must enforce the desktop filesystem scope.
const PATH_TAKING_COMMANDS: &[&str] = &[
    "open_workbook",
    "save_workbook",
    "read_text_file",
    "read_binary_file",
    "read_binary_file_range",
    "stat_file",
    "list_dir",
];

/// The helper that defines the desktop filesystem scope roots.
const DESKTOP_ALLOWED_ROOTS_CALL: &str = "desktop_allowed_roots(";

/// Helper calls that *both* canonicalize and enforce allowed-roots scoping.
///
/// Notes:
/// - We intentionally match on the function name (not the full module path) so refactors like
///   `use crate::fs_scope::canonicalize_in_allowed_roots;` don't break the test.
/// - We include the trailing `(` to reduce false positives from mentions in comments.
const SCOPE_ENFORCING_HELPER_CALLS: &[&str] = &[
    "canonicalize_in_allowed_roots(",
    "canonicalize_in_allowed_roots_with_error(",
    "resolve_save_path_in_allowed_roots(",
];

#[test]
fn path_taking_ipc_commands_must_enforce_fs_scope_canonicalization() {
    let commands_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/commands.rs"));

    let mut failures = Vec::new();
    for command in PATH_TAKING_COMMANDS {
        if let Err(msg) = check_command(commands_rs, command) {
            failures.push(msg);
        }
    }

    if !failures.is_empty() {
        panic!("{}", failures.join("\n\n"));
    }
}

fn check_command(commands_rs: &str, command: &str) -> Result<(), String> {
    let cmd_body = find_pub_async_fn_body(commands_rs, command)?;
    if body_enforces_desktop_fs_scope(cmd_body) {
        return Ok(());
    }

    // `list_dir` intentionally delegates its implementation to `list_dir_blocking`. Since these
    // tests are source-level, we follow that delegation to keep the guardrail focused on the
    // actual filesystem-touching implementation.
    if command == "list_dir" {
        let stripped = strip_comments_and_strings(cmd_body);
        if stripped.contains("list_dir_blocking(") {
            let impl_body = find_fn_body(commands_rs, "list_dir_blocking")?;
            if body_enforces_desktop_fs_scope(impl_body) {
                return Ok(());
            }
            return Err(missing_scope_enforcement_message(
                command,
                "src/commands.rs::list_dir_blocking",
                impl_body,
            ));
        }
    }

    Err(missing_scope_enforcement_message(
        command,
        "src/commands.rs",
        cmd_body,
    ))
}

fn body_enforces_desktop_fs_scope(body: &str) -> bool {
    let stripped = strip_comments_and_strings(body);

    // `desktop_allowed_roots()` defines the HOME/DOCUMENTS/DOWNLOADS policy. By requiring this in
    // addition to canonicalization helpers, we fail conservatively when the implementation is
    // refactored in a way we can't confidently validate (forcing a human review).
    let has_roots = stripped.contains(DESKTOP_ALLOWED_ROOTS_CALL);
    let has_scope_helper = SCOPE_ENFORCING_HELPER_CALLS
        .iter()
        .any(|needle| stripped.contains(needle));

    has_roots && has_scope_helper
}

fn missing_scope_enforcement_message(command: &str, source_hint: &str, body: &str) -> String {
    let stripped = strip_comments_and_strings(body);

    let has_roots = stripped.contains(DESKTOP_ALLOWED_ROOTS_CALL);
    let found_helpers: Vec<&'static str> = SCOPE_ENFORCING_HELPER_CALLS
        .iter()
        .copied()
        .filter(|needle| stripped.contains(needle))
        .collect();

    let mut missing = Vec::new();
    if !has_roots {
        missing.push(format!("missing `{DESKTOP_ALLOWED_ROOTS_CALL}`"));
    }
    if found_helpers.is_empty() {
        missing.push(format!(
            "missing one of: {}",
            SCOPE_ENFORCING_HELPER_CALLS
                .iter()
                .map(|s| format!("`{s}`"))
                .collect::<Vec<_>>()
                .join(", ")
        ));
    }

    format!(
        "fs_scope guardrail failed for IPC command `{command}` ({source_hint}).\n\
         {}\n\
         \n\
         This command must canonicalize paths (resolving `..` + symlinks) and enforce the desktop\n\
         filesystem scope policy (HOME / DOCUMENTS / DOWNLOADS).\n\
         \n\
         See: apps/desktop/src-tauri/src/fs_scope.rs",
        if missing.is_empty() {
            // This should be rare (e.g. if the function moved and our extraction got confused).
            "Unable to detect fs_scope enforcement via source scanning.".to_string()
        } else {
            format!("Problem: {}", missing.join("; "))
        }
    )
}

fn find_pub_async_fn_body<'a>(src: &'a str, fn_name: &str) -> Result<&'a str, String> {
    let needle = format!("pub async fn {fn_name}");
    find_fn_body_by_needle(src, &needle).map_err(|err| {
        format!(
            "{err}\n\
             \n\
             Guardrail expected to find `{needle}` in src/commands.rs so it can ensure filesystem\n\
             scope enforcement isn't accidentally removed."
        )
    })
}

fn find_fn_body<'a>(src: &'a str, fn_name: &str) -> Result<&'a str, String> {
    let needle = format!("fn {fn_name}");
    find_fn_body_by_needle(src, &needle)
}

fn find_fn_body_by_needle<'a>(src: &'a str, needle: &str) -> Result<&'a str, String> {
    let start = src
        .find(needle)
        .ok_or_else(|| format!("failed to find `{needle}`"))?;
    let (open_brace, end) = find_brace_block(src, start)?;
    Ok(&src[open_brace..end])
}

fn find_brace_block(src: &str, start: usize) -> Result<(usize, usize), String> {
    let bytes = src.as_bytes();
    let mut i = start;
    let mut mode = LexMode::Normal;

    // 1) Find the `{` that begins the function body.
    let open_brace = loop {
        if i >= bytes.len() {
            return Err("failed to find function body `{`".to_string());
        }
        match &mut mode {
            LexMode::Normal => {
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'/') {
                    mode = LexMode::LineComment;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    mode = LexMode::BlockComment { depth: 1 };
                    i += 2;
                    continue;
                }
                if let Some((hashes, after_quote)) = parse_raw_string_start(bytes, i) {
                    mode = LexMode::RawString { hashes };
                    i = after_quote;
                    continue;
                }
                if bytes[i] == b'b' && bytes.get(i + 1) == Some(&b'"') {
                    mode = LexMode::String { escaped: false };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'"' {
                    mode = LexMode::String { escaped: false };
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\'' && is_char_literal_start(bytes, i) {
                    mode = LexMode::Char { escaped: false };
                    i += 1;
                    continue;
                }

                if bytes[i] == b'{' {
                    break i;
                }
                i += 1;
            }
            LexMode::LineComment => {
                if bytes[i] == b'\n' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::BlockComment { depth } => {
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    *depth += 1;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && bytes.get(i + 1) == Some(&b'/') {
                    *depth = depth.saturating_sub(1);
                    i += 2;
                    if *depth == 0 {
                        mode = LexMode::Normal;
                    }
                    continue;
                }
                i += 1;
            }
            LexMode::String { escaped } => {
                if *escaped {
                    *escaped = false;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\\' {
                    *escaped = true;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'"' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::Char { escaped } => {
                if *escaped {
                    *escaped = false;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\\' {
                    *escaped = true;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\'' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::RawString { hashes } => {
                if bytes[i] == b'"' && raw_string_ends_here(bytes, i, *hashes) {
                    let h = *hashes;
                    i += 1 + h;
                    mode = LexMode::Normal;
                    continue;
                }
                i += 1;
            }
        }
    };

    // 2) Find the matching `}` by brace counting.
    let mut depth = 0usize;
    i = open_brace;
    mode = LexMode::Normal;
    while i < bytes.len() {
        match &mut mode {
            LexMode::Normal => {
                if bytes[i] == b'{' {
                    depth += 1;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'}' {
                    depth = depth.saturating_sub(1);
                    i += 1;
                    if depth == 0 {
                        return Ok((open_brace, i));
                    }
                    continue;
                }

                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'/') {
                    mode = LexMode::LineComment;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    mode = LexMode::BlockComment { depth: 1 };
                    i += 2;
                    continue;
                }
                if let Some((hashes, after_quote)) = parse_raw_string_start(bytes, i) {
                    mode = LexMode::RawString { hashes };
                    i = after_quote;
                    continue;
                }
                if bytes[i] == b'b' && bytes.get(i + 1) == Some(&b'"') {
                    mode = LexMode::String { escaped: false };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'"' {
                    mode = LexMode::String { escaped: false };
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\'' && is_char_literal_start(bytes, i) {
                    mode = LexMode::Char { escaped: false };
                    i += 1;
                    continue;
                }

                i += 1;
            }
            LexMode::LineComment => {
                if bytes[i] == b'\n' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::BlockComment { depth } => {
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    *depth += 1;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && bytes.get(i + 1) == Some(&b'/') {
                    *depth = depth.saturating_sub(1);
                    i += 2;
                    if *depth == 0 {
                        mode = LexMode::Normal;
                    }
                    continue;
                }
                i += 1;
            }
            LexMode::String { escaped } => {
                if *escaped {
                    *escaped = false;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\\' {
                    *escaped = true;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'"' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::Char { escaped } => {
                if *escaped {
                    *escaped = false;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\\' {
                    *escaped = true;
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\'' {
                    mode = LexMode::Normal;
                }
                i += 1;
            }
            LexMode::RawString { hashes } => {
                if bytes[i] == b'"' && raw_string_ends_here(bytes, i, *hashes) {
                    let h = *hashes;
                    i += 1 + h;
                    mode = LexMode::Normal;
                    continue;
                }
                i += 1;
            }
        }
    }

    Err("failed to find matching `}` for function body".to_string())
}

/// Remove comments and string/char literals from `src` (replacing their contents with spaces),
/// so substring checks don't accidentally match e.g. comments mentioning helper names.
fn strip_comments_and_strings(src: &str) -> String {
    let bytes = src.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0usize;
    let mut mode = LexMode::Normal;

    while i < bytes.len() {
        match &mut mode {
            LexMode::Normal => {
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'/') {
                    out.extend_from_slice(b"  ");
                    mode = LexMode::LineComment;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    out.extend_from_slice(b"  ");
                    mode = LexMode::BlockComment { depth: 1 };
                    i += 2;
                    continue;
                }
                if let Some((hashes, after_quote)) = parse_raw_string_start(bytes, i) {
                    // Blank out the raw-string prefix + opening quote.
                    for _ in i..after_quote {
                        out.push(b' ');
                    }
                    mode = LexMode::RawString { hashes };
                    i = after_quote;
                    continue;
                }
                if bytes[i] == b'b' && bytes.get(i + 1) == Some(&b'"') {
                    out.extend_from_slice(b"  ");
                    mode = LexMode::String { escaped: false };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'"' {
                    out.push(b' ');
                    mode = LexMode::String { escaped: false };
                    i += 1;
                    continue;
                }
                if bytes[i] == b'\'' && is_char_literal_start(bytes, i) {
                    out.push(b' ');
                    mode = LexMode::Char { escaped: false };
                    i += 1;
                    continue;
                }

                out.push(bytes[i]);
                i += 1;
            }
            LexMode::LineComment => {
                if bytes[i] == b'\n' {
                    out.push(b'\n');
                    mode = LexMode::Normal;
                } else {
                    out.push(b' ');
                }
                i += 1;
            }
            LexMode::BlockComment { depth } => {
                if bytes[i] == b'/' && bytes.get(i + 1) == Some(&b'*') {
                    out.extend_from_slice(b"  ");
                    *depth += 1;
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && bytes.get(i + 1) == Some(&b'/') {
                    out.extend_from_slice(b"  ");
                    *depth = depth.saturating_sub(1);
                    i += 2;
                    if *depth == 0 {
                        mode = LexMode::Normal;
                    }
                    continue;
                }
                if bytes[i] == b'\n' {
                    out.push(b'\n');
                } else {
                    out.push(b' ');
                }
                i += 1;
            }
            LexMode::String { escaped } => {
                let b = bytes[i];
                if b == b'\n' {
                    out.push(b'\n');
                } else {
                    out.push(b' ');
                }
                i += 1;

                if *escaped {
                    *escaped = false;
                    continue;
                }
                if b == b'\\' {
                    *escaped = true;
                    continue;
                }
                if b == b'"' {
                    mode = LexMode::Normal;
                }
            }
            LexMode::Char { escaped } => {
                let b = bytes[i];
                if b == b'\n' {
                    out.push(b'\n');
                } else {
                    out.push(b' ');
                }
                i += 1;

                if *escaped {
                    *escaped = false;
                    continue;
                }
                if b == b'\\' {
                    *escaped = true;
                    continue;
                }
                if b == b'\'' {
                    mode = LexMode::Normal;
                }
            }
            LexMode::RawString { hashes } => {
                if bytes[i] == b'"' && raw_string_ends_here(bytes, i, *hashes) {
                    // Blank out closing quote + hashes.
                    out.push(b' ');
                    for _ in 0..*hashes {
                        out.push(b' ');
                    }
                    i += 1 + *hashes;
                    mode = LexMode::Normal;
                    continue;
                }
                if bytes[i] == b'\n' {
                    out.push(b'\n');
                } else {
                    out.push(b' ');
                }
                i += 1;
            }
        }
    }

    String::from_utf8(out).expect("strip_comments_and_strings output must be valid UTF-8")
}

#[derive(Clone, Debug)]
enum LexMode {
    Normal,
    LineComment,
    BlockComment { depth: usize },
    String { escaped: bool },
    Char { escaped: bool },
    RawString { hashes: usize },
}

fn parse_raw_string_start(bytes: &[u8], idx: usize) -> Option<(usize, usize)> {
    // `r"..."`, `r#"..."#`, ... (or `br"..."`, `br#"..."#`, ...)
    if bytes.get(idx) == Some(&b'r') {
        return parse_raw_string_start_after_r(bytes, idx + 1);
    }
    if bytes.get(idx) == Some(&b'b') && bytes.get(idx + 1) == Some(&b'r') {
        return parse_raw_string_start_after_r(bytes, idx + 2);
    }
    None
}

fn parse_raw_string_start_after_r(bytes: &[u8], mut idx: usize) -> Option<(usize, usize)> {
    let mut hashes = 0usize;
    while bytes.get(idx) == Some(&b'#') {
        hashes += 1;
        idx += 1;
    }
    if bytes.get(idx) != Some(&b'"') {
        return None;
    }
    Some((hashes, idx + 1))
}

fn raw_string_ends_here(bytes: &[u8], quote_idx: usize, hashes: usize) -> bool {
    for j in 0..hashes {
        if bytes.get(quote_idx + 1 + j) != Some(&b'#') {
            return false;
        }
    }
    true
}

fn is_char_literal_start(bytes: &[u8], quote_idx: usize) -> bool {
    // Char literal starts with `'` and must have a closing `'`.
    // We only need a best-effort heuristic to avoid treating lifetimes (`'a`) as char literals.
    let next = bytes.get(quote_idx + 1).copied();
    match next {
        Some(b'\\') => true, // escape sequence: '\n', '\xNN', '\u{...}', etc.
        Some(_) => bytes.get(quote_idx + 2) == Some(&b'\''), // simple one-byte char: 'a', '{', etc.
        None => false,
    }
}
