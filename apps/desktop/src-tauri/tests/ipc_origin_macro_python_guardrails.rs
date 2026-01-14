//! Guardrails for macro + python IPC origin enforcement.
//!
//! These are source-level tests so they run in headless CI without the `desktop` feature
//! (which would pull in the system WebView toolchain on Linux).

fn is_likely_char_literal(bytes: &[u8], start: usize) -> bool {
    // Heuristic: treat `'` as a char literal only if we can find a closing `'` within a small
    // window. This avoids misclassifying Rust lifetimes like `'static` as char literals.
    //
    // We only need enough accuracy to reliably match braces for these guardrail tests.
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

fn function_body<'a>(source: &'a str, fn_name: &str) -> &'a str {
    let async_pat = format!("pub async fn {fn_name}");
    let sync_pat = format!("pub fn {fn_name}");
    let start = source
        .find(&async_pat)
        .or_else(|| source.find(&sync_pat))
        .unwrap_or_else(|| panic!("commands.rs missing {fn_name}()"));

    let open_brace = source[start..]
        .find('{')
        .map(|idx| start + idx)
        .unwrap_or_else(|| panic!("commands.rs missing opening brace for {fn_name}()"));

    let close_brace = find_matching_brace(source, open_brace)
        .unwrap_or_else(|| panic!("commands.rs missing closing brace for {fn_name}()"));

    &source[start..=close_brace]
}

#[test]
fn macro_and_python_commands_enforce_ipc_origin() {
    let commands_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/commands.rs"));

    for command in [
        "get_macro_security_status",
        "set_macro_trust",
        "get_vba_project",
        "list_macros",
        "set_macro_ui_context",
        "run_macro",
        "validate_vba_migration",
        "fire_workbook_open",
        "fire_workbook_before_close",
        "fire_worksheet_change",
        "fire_selection_change",
        "run_python_script",
    ] {
        let body = function_body(commands_rs, command);
        assert!(
            body.contains("window: tauri::WebviewWindow"),
            "{command} must accept window: tauri::WebviewWindow so Tauri can inject the caller window for origin enforcement"
        );
        let has_main = body.contains("ensure_main_window(")
            || body.contains("ensure_main_window_and_stable_origin(")
            || body.contains("ensure_main_window_and_trusted_origin(");
        let has_origin = body.contains("ensure_stable_origin(")
            || body.contains("ensure_trusted_origin(")
            || body.contains("ensure_main_window_and_stable_origin(")
            || body.contains("ensure_main_window_and_trusted_origin(");
        assert!(
            has_main,
            "{command} must enforce ipc_origin main-window checks"
        );
        assert!(
            has_origin,
            "{command} must enforce ipc_origin origin checks"
        );
    }
}
