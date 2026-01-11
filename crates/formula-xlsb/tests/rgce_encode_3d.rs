use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord, EncodeError};
use formula_xlsb::workbook_context::WorkbookContext;
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
