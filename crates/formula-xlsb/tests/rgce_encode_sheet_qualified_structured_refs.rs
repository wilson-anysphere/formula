#![cfg(feature = "write")]

use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context_ast, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn ctx_table1_on_sheet1() -> WorkbookContext {
    let mut ctx = WorkbookContext::default();
    ctx.add_table(1, "Table1");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    // Table1 range: A1:B3 on Sheet1.
    ctx.add_table_range(1, "Sheet1".to_string(), 0, 0, 2, 1);
    ctx
}

#[test]
fn ast_encoder_supports_sheet_qualified_structured_ref_with_explicit_table_name() {
    let ctx = ctx_table1_on_sheet1();

    let base = CellCoord::new(0, 0);
    let qualified =
        encode_rgce_with_context_ast("=Sheet1!Table1[Qty]", &ctx, base).expect("encode qualified");
    let unqualified =
        encode_rgce_with_context_ast("=Table1[Qty]", &ctx, base).expect("encode unqualified");

    // BIFF12 `PtgList` does not encode the sheet qualifier; the token stream should match the
    // unqualified form.
    assert_eq!(qualified.rgce, unqualified.rgce);

    let decoded = decode_rgce_with_context(&qualified.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Table1[Qty]");
}

#[test]
fn ast_encoder_supports_sheet_qualified_tableless_structured_ref_by_inferring_table_id() {
    let ctx = ctx_table1_on_sheet1();

    // Base cell inside the table range (A2).
    let base = CellCoord::new(1, 0);
    let encoded =
        encode_rgce_with_context_ast("=Sheet1![@Qty]", &ctx, base).expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[@Qty]");
}

#[test]
fn ast_encoder_rejects_sheet_qualified_structured_ref_when_sheet_does_not_match_table() {
    let ctx = ctx_table1_on_sheet1();

    let err = encode_rgce_with_context_ast("=Sheet2!Table1[Qty]", &ctx, CellCoord::new(0, 0))
        .expect_err("expected sheet mismatch error");

    assert!(
        err.to_string().to_ascii_lowercase().contains("not on sheet"),
        "expected mismatch error, got: {err}"
    );
}

