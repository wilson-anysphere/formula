use formula_xlsb::{OpenOptions, XlsbWorkbook};

#[test]
fn reads_defined_names_from_workbook_bin() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures_metadata/defined-names.xlsb"
    );
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    let names = wb.defined_names();
    assert_eq!(names.len(), 3, "expected 3 defined names, got: {names:?}");

    let zed = names.iter().find(|n| n.name == "ZedName").expect("ZedName");
    assert_eq!(zed.scope_sheet, None);
    assert!(!zed.hidden);
    assert_eq!(
        zed.formula
            .as_ref()
            .and_then(|f| f.text.as_deref())
            .expect("ZedName formula text"),
        "Sheet1!$B$1"
    );

    let local = names
        .iter()
        .find(|n| n.name == "LocalName")
        .expect("LocalName");
    assert_eq!(local.scope_sheet, Some(0));
    assert!(!local.hidden);
    assert_eq!(
        local
            .formula
            .as_ref()
            .and_then(|f| f.text.as_deref())
            .expect("LocalName formula text"),
        "Sheet1!$A$1"
    );

    let hidden = names
        .iter()
        .find(|n| n.name == "HiddenName")
        .expect("HiddenName");
    assert_eq!(hidden.scope_sheet, None);
    assert!(hidden.hidden);
    assert_eq!(
        hidden
            .formula
            .as_ref()
            .and_then(|f| f.text.as_deref())
            .expect("HiddenName formula text"),
        "Sheet1!$A$1:$B$2"
    );
}

#[test]
fn defined_names_do_not_require_preserve_parsed_parts() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures_metadata/defined-names.xlsb"
    );
    let opts = OpenOptions {
        preserve_unknown_parts: false,
        preserve_parsed_parts: false,
        preserve_worksheets: false,
    };
    let wb = XlsbWorkbook::open_with_options(path, opts).expect("open xlsb");
    assert!(
        wb.defined_names().iter().any(|n| n.name == "ZedName"),
        "expected defined names to be parsed even when preserve_parsed_parts=false"
    );
}
