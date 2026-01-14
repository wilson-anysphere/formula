use std::fs;
use std::path::PathBuf;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mode {
    Code,
    LineComment,
    BlockComment { depth: usize },
    NormalString { escape: bool },
    RawString { hashes: usize },
}

fn is_ascii_hexdigit(b: u8) -> bool {
    matches!(b, b'0'..=b'9' | b'a'..=b'f' | b'A'..=b'F')
}

fn consume_char_literal(src: &str, bytes: &[u8], start: usize) -> Option<usize> {
    if bytes.get(start) != Some(&b'\'') {
        return None;
    }
    let mut j = start + 1;
    if j >= bytes.len() {
        return None;
    }
    // Char literals can't span lines; treat lifetimes (no closing `'`) as normal code.
    if matches!(bytes[j], b'\n' | b'\r') {
        return None;
    }

    let j_end = if bytes[j] == b'\\' {
        let esc = *bytes.get(j + 1)?;
        match esc {
            b'\\' | b'\'' | b'"' | b'n' | b'r' | b't' | b'0' => j + 2,
            b'x' => {
                let h1 = *bytes.get(j + 2)?;
                let h2 = *bytes.get(j + 3)?;
                if is_ascii_hexdigit(h1) && is_ascii_hexdigit(h2) {
                    j + 4
                } else {
                    return None;
                }
            }
            b'u' => {
                if bytes.get(j + 2) != Some(&b'{') {
                    return None;
                }
                let mut k = j + 3;
                let mut saw_digit = false;
                while k < bytes.len() {
                    let b = bytes[k];
                    if b == b'}' {
                        break;
                    }
                    if !is_ascii_hexdigit(b) {
                        return None;
                    }
                    saw_digit = true;
                    k += 1;
                }
                if k >= bytes.len() || bytes[k] != b'}' || !saw_digit {
                    return None;
                }
                k + 1
            }
            _ => return None,
        }
    } else {
        // Reject empty literals.
        if bytes[j] == b'\'' {
            return None;
        }
        let ch = src[j..].chars().next()?;
        j += ch.len_utf8();
        j
    };

    if bytes.get(j_end) == Some(&b'\'') {
        Some(j_end + 1)
    } else {
        None
    }
}

fn find_next_code_byte(src: &str, start: usize, needle: u8) -> Option<usize> {
    let bytes = src.as_bytes();
    let mut mode = Mode::Code;
    let mut i = start;

    while i < bytes.len() {
        match mode {
            Mode::Code => {
                if bytes[i] == b'/' && i + 1 < bytes.len() {
                    match bytes[i + 1] {
                        b'/' => {
                            mode = Mode::LineComment;
                            i += 2;
                            continue;
                        }
                        b'*' => {
                            mode = Mode::BlockComment { depth: 1 };
                            i += 2;
                            continue;
                        }
                        _ => {}
                    }
                }

                if bytes[i] == b'\'' {
                    if let Some(next) = consume_char_literal(src, bytes, i) {
                        i = next;
                        continue;
                    }
                }

                // Raw string literals: r"..." / r#"..."# / br"..." / br#"..."#
                if bytes[i] == b'r' {
                    let mut j = i + 1;
                    while j < bytes.len() && bytes[j] == b'#' {
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString {
                            hashes: j - (i + 1),
                        };
                        i = j + 1;
                        continue;
                    }
                } else if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'r' {
                    let mut j = i + 2;
                    while j < bytes.len() && bytes[j] == b'#' {
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString {
                            hashes: j - (i + 2),
                        };
                        i = j + 1;
                        continue;
                    }
                }

                // Normal string literals: "..." and b"..."
                if bytes[i] == b'"' {
                    mode = Mode::NormalString { escape: false };
                    i += 1;
                    continue;
                }
                if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                    mode = Mode::NormalString { escape: false };
                    i += 2; // consume b"
                    continue;
                }

                if bytes[i] == needle {
                    return Some(i);
                }
                i += 1;
            }
            Mode::LineComment => {
                if bytes[i] == b'\n' {
                    mode = Mode::Code;
                }
                i += 1;
            }
            Mode::BlockComment {
                depth: mut comment_depth,
            } => {
                if bytes[i] == b'/' && i + 1 < bytes.len() && bytes[i + 1] == b'*' {
                    comment_depth += 1;
                    mode = Mode::BlockComment {
                        depth: comment_depth,
                    };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && i + 1 < bytes.len() && bytes[i + 1] == b'/' {
                    comment_depth = comment_depth.saturating_sub(1);
                    if comment_depth == 0 {
                        mode = Mode::Code;
                    } else {
                        mode = Mode::BlockComment {
                            depth: comment_depth,
                        };
                    }
                    i += 2;
                    continue;
                }
                mode = Mode::BlockComment {
                    depth: comment_depth,
                };
                i += 1;
            }
            Mode::NormalString { mut escape } => {
                if escape {
                    escape = false;
                    mode = Mode::NormalString { escape };
                    i += 1;
                    continue;
                }
                match bytes[i] {
                    b'\\' => {
                        escape = true;
                        mode = Mode::NormalString { escape };
                        i += 1;
                    }
                    b'"' => {
                        mode = Mode::Code;
                        i += 1;
                    }
                    _ => {
                        mode = Mode::NormalString { escape };
                        i += 1;
                    }
                }
            }
            Mode::RawString { hashes } => {
                if bytes[i] == b'"' {
                    let mut ok = true;
                    for k in 0..hashes {
                        if i + 1 + k >= bytes.len() || bytes[i + 1 + k] != b'#' {
                            ok = false;
                            break;
                        }
                    }
                    if ok {
                        i += 1 + hashes;
                        mode = Mode::Code;
                        continue;
                    }
                }
                i += 1;
            }
        }
    }

    None
}

fn find_next_code_substring(src: &str, start: usize, needle: &str) -> Option<usize> {
    let needle_bytes = needle.as_bytes();
    if needle_bytes.is_empty() {
        return Some(start);
    }
    let mut i = start;
    while let Some(pos) = find_next_code_byte(src, i, needle_bytes[0]) {
        if src.as_bytes().get(pos..pos + needle_bytes.len()) == Some(needle_bytes) {
            return Some(pos);
        }
        i = pos + 1;
    }
    None
}

fn skip_ws_and_comments(src: &str, start: usize) -> usize {
    let bytes = src.as_bytes();
    let mut i = start;
    loop {
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }

        if i + 1 >= bytes.len() {
            break;
        }

        // Line comment
        if bytes[i] == b'/' && bytes[i + 1] == b'/' {
            i += 2;
            while i < bytes.len() && bytes[i] != b'\n' {
                i += 1;
            }
            continue;
        }

        // Block comment (nested)
        if bytes[i] == b'/' && bytes[i + 1] == b'*' {
            i += 2;
            let mut depth = 1usize;
            while i < bytes.len() {
                if i + 1 < bytes.len() && bytes[i] == b'/' && bytes[i + 1] == b'*' {
                    depth += 1;
                    i += 2;
                    continue;
                }
                if i + 1 < bytes.len() && bytes[i] == b'*' && bytes[i + 1] == b'/' {
                    depth = depth.saturating_sub(1);
                    i += 2;
                    if depth == 0 {
                        break;
                    }
                    continue;
                }
                i += 1;
            }
            continue;
        }

        break;
    }
    i
}

fn parse_string_literal(src: &str, start: usize) -> Option<(String, usize)> {
    let bytes = src.as_bytes();
    let b0 = *bytes.get(start)?;

    // Normal string: "..."
    if b0 == b'"' {
        let mut out = String::new();
        let mut i = start + 1;
        let mut escape = false;
        while i < bytes.len() {
            let b = bytes[i];
            if escape {
                match b {
                    b'n' => out.push('\n'),
                    b'r' => out.push('\r'),
                    b't' => out.push('\t'),
                    b'\\' => out.push('\\'),
                    b'"' => out.push('"'),
                    // Best-effort: keep unknown escapes as-is.
                    _ if b.is_ascii() => out.push(b as char),
                    _ => {
                        let ch = src[i..].chars().next()?;
                        out.push(ch);
                    }
                }
                escape = false;
                i += 1;
                continue;
            }

            match b {
                b'\\' => {
                    escape = true;
                    i += 1;
                }
                b'"' => return Some((out, i + 1)),
                _ if b.is_ascii() => {
                    out.push(b as char);
                    i += 1;
                }
                _ => {
                    let ch = src[i..].chars().next()?;
                    out.push(ch);
                    i += ch.len_utf8();
                }
            }
        }
        return None;
    }

    // Byte string: b"..."
    if b0 == b'b' && bytes.get(start + 1) == Some(&b'"') {
        let (s, end) = parse_string_literal(src, start + 1)?;
        return Some((s, end));
    }

    // Raw string: r#"..."#
    if b0 == b'r' || (b0 == b'b' && bytes.get(start + 1) == Some(&b'r')) {
        let mut i = start;
        if b0 == b'b' {
            i += 1; // consume b
        }
        if bytes.get(i) != Some(&b'r') {
            return None;
        }
        i += 1;
        let mut hashes = 0usize;
        while bytes.get(i) == Some(&b'#') {
            hashes += 1;
            i += 1;
        }
        if bytes.get(i) != Some(&b'"') {
            return None;
        }
        i += 1;
        let content_start = i;
        while i < bytes.len() {
            if bytes[i] == b'"' {
                let mut ok = true;
                for k in 0..hashes {
                    if i + 1 + k >= bytes.len() || bytes[i + 1 + k] != b'#' {
                        ok = false;
                        break;
                    }
                }
                if ok {
                    let content = &src[content_start..i];
                    return Some((content.to_string(), i + 1 + hashes));
                }
            }
            i += 1;
        }
        return None;
    }

    None
}

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn load_main_rs_source() -> String {
    let main_rs_path = repo_path("src/main.rs");
    fs::read_to_string(&main_rs_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", main_rs_path.display()))
}

fn extract_tauri_scheme_protocol_handler_block<'a>(main_rs_src: &'a str) -> &'a str {
    // The production desktop binary overrides Tauri's internal asset protocol handler for the
    // `tauri://` scheme so we can inject COOP/COEP headers (required for SharedArrayBuffer /
    // `globalThis.crossOriginIsolated`) while still preserving the CSP computed by Tauri's
    // AssetResolver.
    //
    // This file (`src/main.rs`) only compiles with `--features desktop` because it depends on the
    // system WebView toolchain, so we use a source-level scan to guard against accidental header
    // regressions in headless CI.
    let register_call_start = {
        let mut search = 0usize;
        let needle = ".register_uri_scheme_protocol";
        loop {
            let Some(start) = find_next_code_substring(main_rs_src, search, needle) else {
                panic!(
                    "failed to find `.register_uri_scheme_protocol(\"tauri\", ...)` in src/main.rs"
                );
            };
            let after = start + needle.len();
            let open_paren = find_next_code_byte(main_rs_src, after, b'(').unwrap_or_else(|| {
                panic!(
                    "failed to find `(` after `{needle}` while scanning for the `tauri://` handler"
                )
            });
            let arg_start = skip_ws_and_comments(main_rs_src, open_paren + 1);
            if let Some((scheme, _end)) = parse_string_literal(main_rs_src, arg_start) {
                if scheme == "tauri" {
                    break start;
                }
            }
            search = after;
        }
    };

    let rest = &main_rs_src[register_call_start..];

    // Extract the closure body for the `tauri://` handler.
    //
    // Historically this guard test sliced the builder chain at the next `.plugin(...)`, but that
    // is brittle if `.plugin(` ever appears in a string literal inside the handler.
    let first_pipe = find_next_code_byte(rest, 0, b'|')
        .unwrap_or_else(|| panic!("failed to find closure `|` in the `tauri://` handler block"));
    let second_pipe = find_next_code_byte(rest, first_pipe + 1, b'|').unwrap_or_else(|| {
        panic!("failed to find closing closure `|` in the `tauri://` handler block")
    });
    let open_brace = find_next_code_byte(rest, second_pipe + 1, b'{').unwrap_or_else(|| {
        panic!("failed to find `{{` after closure args in the `tauri://` handler block")
    });

    extract_brace_block(rest, open_brace)
}

fn extract_brace_block<'a>(src: &'a str, open_brace: usize) -> &'a str {
    // Extract a `{ ... }` block while ignoring braces inside strings and comments.
    //
    // We can't rely on a full Rust parser in this guardrail test, but a naive brace counter is
    // brittle because Rust format strings (`"{foo}"`) contain braces frequently.
    let bytes = src.as_bytes();
    assert!(
        bytes.get(open_brace) == Some(&b'{'),
        "expected `open_brace` to point at `{{`, got {:?} at byte offset {open_brace}",
        bytes.get(open_brace).copied()
    );

    let mut mode = Mode::Code;
    let mut depth: i32 = 1;
    let mut i = open_brace + 1;

    while i < bytes.len() {
        match mode {
            Mode::Code => {
                if bytes[i] == b'/' && i + 1 < bytes.len() {
                    match bytes[i + 1] {
                        b'/' => {
                            mode = Mode::LineComment;
                            i += 2;
                            continue;
                        }
                        b'*' => {
                            mode = Mode::BlockComment { depth: 1 };
                            i += 2;
                            continue;
                        }
                        _ => {}
                    }
                }

                // Char literals can contain braces via unicode escapes (`'\u{...}'`). Ignore braces
                // inside those literals.
                if bytes[i] == b'\'' {
                    if let Some(next) = consume_char_literal(src, bytes, i) {
                        i = next;
                        continue;
                    }
                }

                // Raw string literals: r"..." / r#"..."# / br"..." / br#"..."#
                if bytes[i] == b'r' {
                    let mut j = i + 1;
                    while j < bytes.len() && bytes[j] == b'#' {
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString {
                            hashes: j - (i + 1),
                        };
                        i = j + 1;
                        continue;
                    }
                } else if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'r' {
                    let mut j = i + 2;
                    while j < bytes.len() && bytes[j] == b'#' {
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString {
                            hashes: j - (i + 2),
                        };
                        i = j + 1;
                        continue;
                    }
                }

                // Normal string literals: "..." and b"..."
                if bytes[i] == b'"' {
                    mode = Mode::NormalString { escape: false };
                    i += 1;
                    continue;
                }
                if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                    mode = Mode::NormalString { escape: false };
                    i += 2; // consume b"
                    continue;
                }

                match bytes[i] {
                    b'{' => depth += 1,
                    b'}' => {
                        depth -= 1;
                        if depth == 0 {
                            return &src[open_brace..=i];
                        }
                    }
                    _ => {}
                }
                i += 1;
            }
            Mode::LineComment => {
                if bytes[i] == b'\n' {
                    mode = Mode::Code;
                }
                i += 1;
            }
            Mode::BlockComment {
                depth: mut comment_depth,
            } => {
                if bytes[i] == b'/' && i + 1 < bytes.len() && bytes[i + 1] == b'*' {
                    comment_depth += 1;
                    mode = Mode::BlockComment {
                        depth: comment_depth,
                    };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && i + 1 < bytes.len() && bytes[i + 1] == b'/' {
                    comment_depth = comment_depth.saturating_sub(1);
                    if comment_depth == 0 {
                        mode = Mode::Code;
                    } else {
                        mode = Mode::BlockComment {
                            depth: comment_depth,
                        };
                    }
                    i += 2;
                    continue;
                }
                // Keep scanning inside block comment.
                mode = Mode::BlockComment {
                    depth: comment_depth,
                };
                i += 1;
            }
            Mode::NormalString { mut escape } => {
                if escape {
                    escape = false;
                    mode = Mode::NormalString { escape };
                    i += 1;
                    continue;
                }
                match bytes[i] {
                    b'\\' => {
                        escape = true;
                        mode = Mode::NormalString { escape };
                        i += 1;
                    }
                    b'"' => {
                        mode = Mode::Code;
                        i += 1;
                    }
                    _ => {
                        mode = Mode::NormalString { escape };
                        i += 1;
                    }
                }
            }
            Mode::RawString { hashes } => {
                if bytes[i] == b'"' {
                    let mut ok = true;
                    for k in 0..hashes {
                        if i + 1 + k >= bytes.len() || bytes[i + 1 + k] != b'#' {
                            ok = false;
                            break;
                        }
                    }
                    if ok {
                        // Consume the closing delimiter (`"` plus `hashes` `#` characters).
                        i += 1 + hashes;
                        mode = Mode::Code;
                        continue;
                    }
                }
                i += 1;
            }
        }
    }

    panic!("unterminated brace block starting at byte offset {open_brace}");
}

fn extract_tauri_scheme_protocol_asset_success_block<'a>(handler_block: &'a str) -> &'a str {
    // The `tauri://` handler should apply COOP/COEP headers on *successful* asset responses.
    //
    // Note: production requests hit the `AssetResolver` path, not the `--startup-bench` early return.
    // We therefore extract the `Some(asset)` / `if let Some(asset)` block specifically.
    let if_let_marker = "if let Some(asset)";
    if let Some(start) = handler_block.find(if_let_marker) {
        let rest = &handler_block[start..];
        let open = rest.find('{').unwrap_or_else(|| {
            panic!("failed to find `{{` after `{if_let_marker}` in the `tauri://` handler block")
        });
        let block = extract_brace_block(rest, open);
        return &rest[..open + block.len()];
    }

    // Backwards-compat: older implementations used a `match` arm (`Some(asset) => { ... }`).
    let match_arm_marker = "Some(asset)";
    if let Some(start) = handler_block.find(match_arm_marker) {
        let rest = &handler_block[start..];
        let open = rest.find('{').unwrap_or_else(|| {
            panic!("failed to find `{{` after `{match_arm_marker}` in the `tauri://` handler block")
        });
        let block = extract_brace_block(rest, open);
        return &rest[..open + block.len()];
    }

    panic!(
        "failed to find the `AssetResolver` success path (`if let Some(asset)` or `Some(asset) =>`) in the `tauri://` handler block.\n\
         Searched within:\n\
         {handler_block}"
    );
}

#[test]
fn extract_brace_block_ignores_braces_in_strings_and_comments() {
    // Keep the outer raw string delimiter longer than the inner raw string so the inner `"#`
    // doesn't prematurely terminate the test source string.
    let src = r##"
{
  // { braces in comment }
  let _a = "{ braces in string }";
  let _b = r#"raw { braces }"#;
  let _c = br#"raw bytes { braces }"#;
  let _d = '\u{7B}'; // { braces in char escape }
  /* nested { block { comment } } */
  if true { let _e = 1; }
}
"##;

    let open = src.find('{').expect("test src should contain `{`");
    let block = extract_brace_block(src, open);

    assert!(block.contains(r#"let _a = "{ braces in string }";"#));
    assert!(block.contains("let _b = r#\"raw { braces }\"#;"));
    assert!(block.contains("let _c = br#\"raw bytes { braces }\"#;"));
    assert!(block.contains(r#"let _d = '\u{7B}';"#));
    assert!(block.contains(r#"/* nested { block { comment } } */"#));
    assert!(
        block.contains("if true { let _e = 1; }"),
        "expected nested code braces to still be counted"
    );
    assert!(
        block.trim_end().ends_with('}'),
        "expected brace matcher to include the final outer `}}`, got:\n{block}"
    );
}

#[test]
fn extract_tauri_handler_extractor_ignores_tokens_in_comments_and_strings() {
    let src = r#"
// This comment intentionally includes a tauri protocol registration-like snippet.
// If the extractor naively searches by substring without tracking comment state, it can start in
// the middle of this comment and mis-parse the handler.
// .register_uri_scheme_protocol("tauri", move |_ctx, _request| { /* not real */ })

tauri::Builder::default()
    .register_uri_scheme_protocol(
        "asset",
        asset_protocol::handler,
    )
    .register_uri_scheme_protocol("tauri", /* comment with | and { } */ move |_ctx, request|
        /* comment with { braces } before the body */ {
        let _s = ".plugin(";
        if let Some(asset) = _ctx.app_handle().asset_resolver().get("index.html".to_string()) {
            let mut response = Response::builder()
                .status(StatusCode::OK)
                .body(Vec::new())
                .unwrap();
            apply_cross_origin_isolation_headers(&mut response);
            return response;
        }
        Response::builder()
            .status(StatusCode::NOT_FOUND)
            .body(Vec::new())
            .unwrap()
    })
    .plugin(tauri_plugin_dialog::init());
"#;

    let handler_block = extract_tauri_scheme_protocol_handler_block(src);
    assert!(
        handler_block.contains("apply_cross_origin_isolation_headers("),
        "expected extracted handler block to contain the COOP/COEP helper call.\nExtracted:\n{handler_block}"
    );
    assert!(
        handler_block.contains("if let Some(asset)"),
        "expected extracted handler block to include the AssetResolver success path.\nExtracted:\n{handler_block}"
    );
    assert!(
        !handler_block.contains(".plugin(tauri_plugin_dialog"),
        "expected extracted handler block to stop at the end of the closure body, not include the surrounding builder chain.\nExtracted:\n{handler_block}"
    );
}

#[test]
fn tauri_scheme_protocol_handler_injects_cross_origin_isolation_headers_and_preserves_csp() {
    let main_rs_src = load_main_rs_source();

    // 1) Ensure `src/main.rs` registers a custom `tauri://` protocol handler.
    //
    // `extract_tauri_scheme_protocol_handler_block` will panic with a targeted error message if
    // the call is missing.
    let handler_block = extract_tauri_scheme_protocol_handler_block(&main_rs_src);
    let asset_success_block = extract_tauri_scheme_protocol_asset_success_block(handler_block);

    // 2) Ensure the handler applies COOP/COEP headers (required for cross-origin isolation).
    //
    // Prefer checking for the helper call (current implementation), but allow direct insertion of
    // either the header names or the associated constants in case the handler is refactored.
    //
    // Important: scan the *successful asset response path* (`Some(asset)` / `if let Some(asset)`),
    // not just the overall handler block, so:
    // - a `--startup-bench`-only header injection doesn't mask a production regression
    // - a "headers only on 404 responses" bug doesn't pass the guardrail.
    let has_apply_call = asset_success_block.contains("apply_cross_origin_isolation_headers(");
    let has_header_strings = asset_success_block.contains("cross-origin-opener-policy")
        && asset_success_block.contains("cross-origin-embedder-policy");
    let has_header_constants = asset_success_block.contains("CROSS_ORIGIN_OPENER_POLICY")
        && asset_success_block.contains("CROSS_ORIGIN_EMBEDDER_POLICY");

    assert!(
        has_apply_call || has_header_strings || has_header_constants,
        "expected the `tauri://` protocol handler to inject COOP/COEP headers.\n\
         Missing `apply_cross_origin_isolation_headers(...)` (or direct insertion of the COOP/COEP header names).\n\
         This is required for `globalThis.crossOriginIsolated` and SharedArrayBuffer in packaged desktop builds.\n\
         Searched within the AssetResolver success path of the `.register_uri_scheme_protocol(\"tauri\", ...)` handler:\n\
         {asset_success_block}"
    );

    // 3) Ensure we preserve Tauri's configured CSP header when available.
    //
    // Dropping the CSP in the custom handler would reduce production parity with the stock Tauri
    // asset protocol and can break worker/module loading.
    assert!(
        asset_success_block.contains("asset.csp_header"),
        "expected the `tauri://` protocol handler to reference `asset.csp_header` (to preserve the CSP computed by Tauri's AssetResolver)"
    );
    assert!(
        asset_success_block.contains("Content-Security-Policy"),
        "expected the `tauri://` protocol handler to set the `Content-Security-Policy` header when `asset.csp_header` is present"
    );

    let csp_is_conditional = (asset_success_block.contains("if let Some")
        && asset_success_block.contains("asset.csp_header"))
        || asset_success_block.contains("match asset.csp_header")
        || asset_success_block.contains("asset.csp_header.map")
        || asset_success_block.contains("asset.csp_header.is_some")
        || asset_success_block.contains("asset.csp_header.is_some_and");

    assert!(
        csp_is_conditional,
        "expected the `Content-Security-Policy` header to be set conditionally based on `asset.csp_header` (so we don't emit an empty/malformed CSP when absent).\n\
         Searched within the AssetResolver success path of the `.register_uri_scheme_protocol(\"tauri\", ...)` handler:\n\
         {asset_success_block}"
    );
}
