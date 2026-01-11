use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord, EncodeError};
use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

#[test]
fn encodes_and_decodes_sheet_qualified_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context("=Sheet2!A1+1", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!A1+1");
    assert!(encoded.rgcb.is_empty());
}

#[test]
fn encodes_and_decodes_sheet_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded = encode_rgce_with_context("=SUM(Sheet1:Sheet3!A1)", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn encodes_and_decodes_workbook_name() {
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyNamedRange", 1);

    let encoded =
        encode_rgce_with_context("=MyNamedRange", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "MyNamedRange");
}

#[test]
fn unknown_name_is_a_structured_error() {
    let ctx = WorkbookContext::default();
    let err = encode_rgce_with_context("=NoSuchName", &ctx, CellCoord::new(0, 0))
        .expect_err("should fail");
    assert_eq!(
        err,
        EncodeError::UnknownName {
            name: "NoSuchName".to_string()
        }
    );
}

#[test]
fn encodes_and_decodes_builtin_function_via_ftab() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=ABS(-1)", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "ABS(-1)");
}

#[test]
fn encodes_addin_udf_calls_via_namex() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded = encode_rgce_with_context("=MyAddinFunc(1,2)", ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x39, 0x00, 0x00, 0x01, 0x00, // PtgNameX(ixti=0, nameIndex=1)
            0x22, 0x03, 0xFF, 0x00, // PtgFuncVar(argc=3, iftab=0x00FF)
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "MyAddinFunc(1,2)");
}
