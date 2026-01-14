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

    // The `tauri://` protocol handler is registered in the `tauri::Builder` chain and is
    // immediately followed by a `.plugin(...)` call in current production builds.
    //
    // We intentionally avoid brittle brace/paren matching and instead slice the builder chain at
    // the next method call boundary.
    let end = rest
        .find(".plugin(")
        .or_else(|| rest.find(".build("))
        .unwrap_or(rest.len());

    &rest[..end]
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
