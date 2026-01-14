//! Guardrails for marketplace IPC size caps.
//!
//! These are source-level tests so they run in headless CI without the `desktop` feature
//! (which would pull in the system WebView toolchain on Linux).

fn is_likely_char_literal(bytes: &[u8], start: usize) -> bool {
    // Heuristic: treat `'` as a char literal only if we can find a closing `'` within a small
    // window. This avoids misclassifying Rust lifetimes like `'static` as char literals.
    let mut idx = start + 1;
    let mut escape = false;
    while idx < bytes.len() && idx - start <= 16 {
        let b = bytes[idx];
        if b == b'\n' {
            break;
        }
        if escape {
            escape = false;
            idx += 1;
            continue;
        }
        if b == b'\\' {
            escape = true;
            idx += 1;
            continue;
        }
        if b == b'\'' {
            return true;
        }
        idx += 1;
    }
    false
}

fn find_matching_brace(source: &str, open_brace: usize) -> Option<usize> {
    #[derive(Clone, Copy, Debug)]
    enum Mode {
        Normal,
        LineComment,
        BlockComment { depth: usize },
        String { escape: bool },
        Char { escape: bool },
        RawString { hashes: usize },
    }

    let bytes = source.as_bytes();
    let mut mode = Mode::Normal;
    let mut depth: i32 = 0;
    let mut i = open_brace;

    while i < bytes.len() {
        match mode {
            Mode::Normal => {
                // Raw byte string: br###"..."###
                if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'r' {
                    let mut j = i + 2;
                    let mut hashes = 0usize;
                    while j < bytes.len() && bytes[j] == b'#' {
                        hashes += 1;
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString { hashes };
                        i = j + 1;
                        continue;
                    }
                }

                // Raw string: r###"..."###
                if bytes[i] == b'r' {
                    let mut j = i + 1;
                    let mut hashes = 0usize;
                    while j < bytes.len() && bytes[j] == b'#' {
                        hashes += 1;
                        j += 1;
                    }
                    if j < bytes.len() && bytes[j] == b'"' {
                        mode = Mode::RawString { hashes };
                        i = j + 1;
                        continue;
                    }
                }

                // Byte string: b"..."
                if bytes[i] == b'b' && i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                    mode = Mode::String { escape: false };
                    i += 2;
                    continue;
                }

                if bytes[i] == b'"' {
                    mode = Mode::String { escape: false };
                    i += 1;
                    continue;
                }

                if bytes[i] == b'\'' && is_likely_char_literal(bytes, i) {
                    mode = Mode::Char { escape: false };
                    i += 1;
                    continue;
                }

                if bytes[i] == b'/' && i + 1 < bytes.len() {
                    if bytes[i + 1] == b'/' {
                        mode = Mode::LineComment;
                        i += 2;
                        continue;
                    }
                    if bytes[i + 1] == b'*' {
                        mode = Mode::BlockComment { depth: 1 };
                        i += 2;
                        continue;
                    }
                }

                match bytes[i] {
                    b'{' => depth += 1,
                    b'}' => {
                        depth -= 1;
                        if depth == 0 {
                            return Some(i);
                        }
                    }
                    _ => {}
                }

                i += 1;
            }
            Mode::LineComment => {
                if bytes[i] == b'\n' {
                    mode = Mode::Normal;
                }
                i += 1;
            }
            Mode::BlockComment { mut depth } => {
                if bytes[i] == b'/' && i + 1 < bytes.len() && bytes[i + 1] == b'*' {
                    depth += 1;
                    mode = Mode::BlockComment { depth };
                    i += 2;
                    continue;
                }
                if bytes[i] == b'*' && i + 1 < bytes.len() && bytes[i + 1] == b'/' {
                    depth = depth.saturating_sub(1);
                    i += 2;
                    if depth == 0 {
                        mode = Mode::Normal;
                    } else {
                        mode = Mode::BlockComment { depth };
                    }
                    continue;
                }
                i += 1;
            }
            Mode::String { mut escape } => {
                let b = bytes[i];
                if escape {
                    escape = false;
                    mode = Mode::String { escape };
                    i += 1;
                    continue;
                }
                if b == b'\\' {
                    escape = true;
                    mode = Mode::String { escape };
                    i += 1;
                    continue;
                }
                if b == b'"' {
                    mode = Mode::Normal;
                    i += 1;
                    continue;
                }
                i += 1;
            }
            Mode::Char { mut escape } => {
                let b = bytes[i];
                if escape {
                    escape = false;
                    mode = Mode::Char { escape };
                    i += 1;
                    continue;
                }
                if b == b'\\' {
                    escape = true;
                    mode = Mode::Char { escape };
                    i += 1;
                    continue;
                }
                if b == b'\'' {
                    mode = Mode::Normal;
                    i += 1;
                    continue;
                }
                i += 1;
            }
            Mode::RawString { hashes } => {
                if bytes[i] == b'"' {
                    let mut ok = true;
                    for h in 0..hashes {
                        if i + 1 + h >= bytes.len() || bytes[i + 1 + h] != b'#' {
                            ok = false;
                            break;
                        }
                    }
                    if ok {
                        mode = Mode::Normal;
                        i += 1 + hashes;
                        continue;
                    }
                }
                i += 1;
            }
        }
    }

    None
}

fn find_fn_body_open_brace(source: &str, start: usize) -> Option<usize> {
    // Some signatures contain `{}` braces inside const generics (`Foo<{ MAX }>`). A naive
    // `find('{')` would stop at the wrong brace. Best-effort: scan for the first `{` that is not
    // inside `()`, `[]`, or `<>`.
    let bytes = source.as_bytes();
    let mut paren_depth = 0usize;
    let mut bracket_depth = 0usize;
    let mut angle_depth = 0usize;

    let mut i = start;
    while i < bytes.len() {
        match bytes[i] {
            b'(' => paren_depth += 1,
            b')' => paren_depth = paren_depth.saturating_sub(1),
            b'[' => bracket_depth += 1,
            b']' => bracket_depth = bracket_depth.saturating_sub(1),
            b'<' => angle_depth += 1,
            b'>' => angle_depth = angle_depth.saturating_sub(1),
            b'{' if paren_depth == 0 && bracket_depth == 0 && angle_depth == 0 => return Some(i),
            _ => {}
        }
        i += 1;
    }

    None
}

fn function_body<'a>(source: &'a str, fn_name: &str) -> &'a str {
    let async_pat = format!("pub async fn {fn_name}");
    let sync_pat = format!("pub fn {fn_name}");
    let start = source
        .find(&async_pat)
        .or_else(|| source.find(&sync_pat))
        .unwrap_or_else(|| panic!("commands.rs missing {fn_name}()"));

    let open_brace = find_fn_body_open_brace(source, start)
        .unwrap_or_else(|| panic!("commands.rs missing opening brace for {fn_name}()"));
    let close_brace = find_matching_brace(source, open_brace)
        .unwrap_or_else(|| panic!("commands.rs missing closing brace for {fn_name}()"));

    &source[start..=close_brace]
}

#[test]
fn marketplace_commands_use_bounded_body_reads() {
    let commands_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/commands.rs"));

    let search = function_body(commands_rs, "marketplace_search");
    assert!(
        search.contains("marketplace_fetch_json_with_limit")
            || search.contains("read_response_body_with_limit"),
        "marketplace_search must fetch JSON via the bounded body reader"
    );
    assert!(
        search.contains("MARKETPLACE_JSON_MAX_BODY_BYTES"),
        "marketplace_search must apply the marketplace JSON body cap constant"
    );
    assert!(
        !search.contains(".json().await") && !search.contains(".bytes().await"),
        "marketplace_search must not call reqwest Response::json/bytes directly"
    );

    let get_ext = function_body(commands_rs, "marketplace_get_extension");
    assert!(
        get_ext.contains("marketplace_fetch_optional_json_with_limit")
            || get_ext.contains("read_response_body_with_limit"),
        "marketplace_get_extension must fetch JSON via the bounded body reader"
    );
    assert!(
        get_ext.contains("MARKETPLACE_JSON_MAX_BODY_BYTES"),
        "marketplace_get_extension must apply the marketplace JSON body cap constant"
    );
    assert!(
        !get_ext.contains(".json().await") && !get_ext.contains(".bytes().await"),
        "marketplace_get_extension must not call reqwest Response::json/bytes directly"
    );

    let download = function_body(commands_rs, "marketplace_download_package");
    assert!(
        download.contains("marketplace_fetch_optional_download_payload")
            || download.contains("read_response_body_with_limit"),
        "marketplace_download_package must fetch bytes via the bounded body reader"
    );
    assert!(
        !download.contains(".json().await") && !download.contains(".bytes().await"),
        "marketplace_download_package must not call reqwest Response::json/bytes directly"
    );
}

