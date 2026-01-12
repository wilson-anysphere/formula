use formula_xlsb::{OpenOptions, XlsbWorkbook};
use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord};

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
        decode_formulas: true,
    };
    let wb = XlsbWorkbook::open_with_options(path, opts).expect("open xlsb");
    assert!(
        wb.defined_names().iter().any(|n| n.name == "ZedName"),
        "expected defined names to be parsed even when preserve_parsed_parts=false"
    );
}

#[test]
fn rgce_codec_resolves_defined_name_tokens_via_workbook_context() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures_metadata/defined-names.xlsb"
    );
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    // Workbook-scoped name -> PtgName with index 1.
    let encoded = encode_rgce_with_context("=ZedName", ctx, CellCoord::new(0, 0))
        .expect("encode formula with defined name");
    assert_eq!(
        encoded.rgce,
        vec![
            0x23, // PtgName (ref class)
            0x01, 0x00, 0x00, 0x00, // nameId
            0x00, 0x00, // reserved
        ]
    );
    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode PtgName");
    assert_eq!(decoded, "ZedName");

    // Sheet-scoped name -> PtgName with index 2.
    let encoded = encode_rgce_with_context("=Sheet1!LocalName", ctx, CellCoord::new(0, 0))
        .expect("encode formula with sheet-scoped defined name");
    assert_eq!(
        encoded.rgce,
        vec![
            0x23, // PtgName (ref class)
            0x02, 0x00, 0x00, 0x00, // nameId
            0x00, 0x00, // reserved
        ]
    );
    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode PtgName");
    assert_eq!(decoded, "Sheet1!LocalName");
}
