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

#[test]
fn save_workbook_xlsb_branch_uses_no_bytes_writer() {
    let source = include_str!("../src/commands.rs");

    let save_fn_start = source
        .find("pub async fn save_workbook")
        .expect("expected commands.rs to define save_workbook");
    let save_fn_open_brace = source[save_fn_start..]
        .find('{')
        .map(|idx| save_fn_start + idx)
        .expect("expected save_workbook to have an opening brace");
    let save_fn_body = extract_braced_block(source, save_fn_open_brace);

    let xlsb_if_start = save_fn_body
        .find("if ext.eq_ignore_ascii_case(\"xlsb\")")
        .expect("expected save_workbook to contain an .xlsb branch");
    let xlsb_if_open_brace = save_fn_body[xlsb_if_start..]
        .find('{')
        .map(|idx| xlsb_if_start + idx)
        .expect("expected .xlsb branch to have an opening brace");
    let xlsb_if_body = extract_braced_block(save_fn_body, xlsb_if_open_brace);

    assert!(
        xlsb_if_body.contains("write_xlsb_to_disk_blocking"),
        "expected .xlsb save path to call the no-bytes XLSB writer"
    );
    assert!(
        !xlsb_if_body.contains("write_xlsx_blocking"),
        "expected .xlsb save path to avoid write_xlsx_blocking (which reads the saved file back into memory)"
    );
}

