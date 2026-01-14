use std::fs;
use std::path::PathBuf;

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
    let start_marker = ".register_uri_scheme_protocol(\"tauri\"";
    let start = main_rs_src
        .find(start_marker)
        .unwrap_or_else(|| panic!("failed to find `{start_marker}` in src/main.rs"));

    let rest = &main_rs_src[start..];

    // Extract the closure body for the `tauri://` handler.
    //
    // Historically this guard test sliced the builder chain at the next `.plugin(...)`, but that
    // is brittle if `.plugin(` ever appears in a string literal inside the handler.
    let first_pipe = rest
        .find('|')
        .unwrap_or_else(|| panic!("failed to find closure `|` in the `tauri://` handler block"));
    let second_pipe = rest[first_pipe + 1..]
        .find('|')
        .map(|idx| first_pipe + 1 + idx)
        .unwrap_or_else(|| {
            panic!("failed to find closing closure `|` in the `tauri://` handler block")
        });
    let open_brace = rest[second_pipe + 1..]
        .find('{')
        .map(|idx| second_pipe + 1 + idx)
        .unwrap_or_else(|| {
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
fn tauri_scheme_protocol_handler_injects_cross_origin_isolation_headers_and_preserves_csp() {
    let main_rs_src = load_main_rs_source();

    // 1) Ensure `src/main.rs` registers a custom `tauri://` protocol handler.
    assert!(
        main_rs_src.contains(".register_uri_scheme_protocol(\"tauri\""),
        "expected src/main.rs to register a `tauri://` URI scheme handler via `.register_uri_scheme_protocol(\"tauri\", ...)`"
    );

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
