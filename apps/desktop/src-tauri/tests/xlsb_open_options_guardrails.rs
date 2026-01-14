//! Guardrails to ensure XLSB imports use minimal preservation options.
//!
//! This is a source-level test so it runs in headless CI without the `desktop` feature.
//! Opening `.xlsb` with `formula-xlsb` defaults to preserving raw package parts for future
//! round-tripping, but that can retain very large buffers (notably shared strings) and inflate
//! open-time memory usage.

fn extract_braced_block(source: &str, open_brace_index: usize) -> &str {
    assert_eq!(
        source.as_bytes().get(open_brace_index).copied(),
        Some(b'{'),
        "expected '{{' at open_brace_index"
    );

    let bytes = source.as_bytes();
    let mut depth: i32 = 0;
    let mut in_string: Option<u8> = None;
    let mut escape = false;

    let mut start: Option<usize> = None;
    for (rel_idx, &b) in bytes[open_brace_index..].iter().enumerate() {
        if let Some(quote) = in_string {
            if escape {
                escape = false;
                continue;
            }
            if b == b'\\' {
                escape = true;
                continue;
            }
            if b == quote {
                in_string = None;
            }
            continue;
        }

        match b {
            b'"' | b'\'' => in_string = Some(b),
            b'{' => {
                depth += 1;
                if depth == 1 {
                    start = Some(open_brace_index + rel_idx + 1);
                }
            }
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    let start = start.expect("missing block start index");
                    let end = open_brace_index + rel_idx;
                    return &source[start..end];
                }
                if depth < 0 {
                    panic!("unbalanced braces while scanning source");
                }
            }
            _ => {}
        }
    }

    panic!("unclosed brace block while scanning source");
}

fn strip_whitespace(source: &str) -> String {
    source.chars().filter(|ch| !ch.is_whitespace()).collect()
}

#[test]
fn read_xlsb_blocking_opens_with_minimal_preservation() {
    let source = include_str!("../src/file_io.rs");
    let fn_start = source
        .find("fn read_xlsb_blocking")
        .expect("expected file_io.rs to define read_xlsb_blocking");
    let open_brace = source[fn_start..]
        .find('{')
        .map(|idx| fn_start + idx)
        .expect("expected read_xlsb_blocking to have an opening brace");
    let body = extract_braced_block(source, open_brace);

    let body = strip_whitespace(body);

    assert!(
        body.contains("XlsbWorkbook::open_with_options("),
        "expected read_xlsb_blocking to call XlsbWorkbook::open_with_options(...)"
    );
    assert!(
        body.contains("preserve_unknown_parts:false"),
        "expected read_xlsb_blocking to set preserve_unknown_parts: false"
    );
    assert!(
        body.contains("preserve_parsed_parts:false"),
        "expected read_xlsb_blocking to set preserve_parsed_parts: false"
    );

    // Guard against future changes that may unintentionally regress behavior or memory usage.
    assert!(
        body.contains("preserve_worksheets:false"),
        "expected read_xlsb_blocking to keep preserve_worksheets: false"
    );
    assert!(
        body.contains("decode_formulas:true"),
        "expected read_xlsb_blocking to keep decode_formulas: true"
    );
}

