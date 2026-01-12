use pretty_assertions::assert_eq;

fn extract_const_str_list(src: &str, const_name: &str) -> Vec<String> {
    let marker = format!("const {const_name}: &[&str] = &[");
    let start = src
        .find(&marker)
        .unwrap_or_else(|| panic!("could not find start of `{const_name}` list"));
    let rest = &src[start + marker.len()..];
    let end = rest
        .find("];")
        .unwrap_or_else(|| panic!("could not find end of `{const_name}` list"));
    let body = &rest[..end];

    body.lines()
        .filter_map(|line| {
            let line = line.trim();
            let line = line.strip_prefix('"')?;
            let end = line.find('"')?;
            Some(line[..end].to_string())
        })
        .collect()
}

#[test]
fn xlfn_required_function_lists_are_in_sync() {
    let xlsx_src = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/formula_text.rs"));
    let biff_src = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-biff/src/ftab.rs"
    ));

    let xlsx = extract_const_str_list(xlsx_src, "XL_FN_REQUIRED_FUNCTIONS");
    let biff = extract_const_str_list(biff_src, "FUTURE_UDF_FUNCTIONS");

    assert_eq!(
        xlsx, biff,
        "The OOXML `_xlfn.` required-function list (formula-xlsx) must match the BIFF future/UDF \
         list (formula-biff) so roundtrips and minimal builds preserve Excel's file format."
    );
}

