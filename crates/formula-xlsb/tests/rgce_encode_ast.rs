#![cfg(feature = "write")]

use formula_engine::parse_formula;
use formula_xlsb::rgce::{
    decode_rgce, decode_rgce_with_context, decode_rgce_with_rgcb, encode_rgce_with_context_ast,
    CellCoord,
};
use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

fn ctx_table1() -> WorkbookContext {
    let mut ctx = WorkbookContext::default();
    ctx.add_table(1, "Table1");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    ctx.add_table_column(1, 3, "Price");
    ctx.add_table_column(1, 4, "Total");
    ctx
}

fn normalize_formula(src: &str) -> String {
    let ast = parse_formula(src, Default::default()).expect("parse");
    ast.to_string(Default::default()).expect("serialize")
}

#[test]
fn ast_encoder_roundtrips_3d_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context_ast("=Sheet2!A1+1", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!A1+1");
}

#[test]
fn ast_encoder_roundtrips_area_ref() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=A1:B2", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "A1:B2");
}

#[test]
fn ast_encoder_roundtrips_intersection_operator() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=A1 B1", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "A1 B1");
}

#[test]
fn ast_encoder_roundtrips_union_operator_inside_function_arg() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=IF(1,(A1,B1))", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "IF(1,(A1,B1))");
}

#[test]
fn ast_encoder_roundtrips_percent_operator() {
    let ctx = WorkbookContext::default();

    let encoded = encode_rgce_with_context_ast("=10%", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "10%");
}

#[test]
fn ast_encoder_roundtrips_spill_operator() {
    let ctx = WorkbookContext::default();

    let encoded = encode_rgce_with_context_ast("=A1#", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "A1#");
}

#[test]
fn ast_encoder_roundtrips_implicit_intersection_on_area() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=@A1:A10", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "@A1:A10");
}

#[test]
fn ast_encoder_roundtrips_3d_area_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context_ast("=Sheet2!A1:B2", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!A1:B2");
}

#[test]
fn ast_encoder_roundtrips_sheet_range_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted =
        encode_rgce_with_context_ast("=SUM(Sheet1:Sheet3!A1)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert!(encoded_unquoted.rgcb.is_empty());

    let encoded_quoted =
        encode_rgce_with_context_ast("=SUM('Sheet1:Sheet3'!A1)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert!(encoded_quoted.rgcb.is_empty());
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('Sheet1:Sheet3'!A1)");
}

#[test]
fn ast_encoder_roundtrips_external_workbook_sheet_range_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsx", "SheetA", "SheetB", 0);

    let encoded_unquoted = encode_rgce_with_context_ast(
        "=SUM([Book2.xlsx]SheetA:SheetB!A1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert!(encoded_unquoted.rgcb.is_empty());

    // Round-trip through the parseable single-token prefix emitted by the rgce decoder.
    let encoded_quoted = encode_rgce_with_context_ast(
        "=SUM('[Book2.xlsx]SheetA:SheetB'!A1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert!(encoded_quoted.rgcb.is_empty());
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('[Book2.xlsx]SheetA:SheetB'!A1)");
}

#[test]
fn ast_encoder_roundtrips_implicit_intersection_on_name() {
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyNamedRange", 1);

    let encoded =
        encode_rgce_with_context_ast("=@MyNamedRange", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@MyNamedRange");
}

#[test]
fn ast_encoder_roundtrips_defined_name() {
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyNamedRange", 1);

    let encoded =
        encode_rgce_with_context_ast("=MyNamedRange", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "MyNamedRange");
}

#[test]
fn ast_encoder_roundtrips_structured_ref_table_column() {
    let ctx = ctx_table1();
    let encoded =
        encode_rgce_with_context_ast("=Table1[Qty]", &ctx, CellCoord::new(0, 0)).expect("encode");

    assert_eq!(
        encoded.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            1, 0, 0, 0, // table id
            0, 0, // flags
            2, 0, // col_first
            2, 0, // col_last
            0, 0, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Table1[Qty]");
}

#[test]
fn ast_encoder_roundtrips_structured_ref_this_row_without_table_name() {
    let ctx = ctx_table1();

    let encoded =
        encode_rgce_with_context_ast("=[@Qty]", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            1, 0, 0, 0, // table id (inferred)
            0x10, 0x00, // flags (#This Row)
            2, 0, // col_first
            2, 0, // col_last
            0, 0, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[@Qty]");
    assert_eq!(normalize_formula(&decoded), normalize_formula("[@Qty]"));
}

#[test]
fn ast_encoder_roundtrips_structured_ref_this_row_all_columns_without_table_name() {
    let ctx = ctx_table1();

    let encoded =
        encode_rgce_with_context_ast("=[@]", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            1, 0, 0, 0, // table id (inferred)
            0x10, 0x00, // flags (#This Row)
            0, 0, // col_first (all columns)
            0, 0, // col_last (all columns)
            0, 0, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[@]");
    assert_eq!(normalize_formula(&decoded), normalize_formula("[@]"));
}

#[test]
fn ast_encoder_roundtrips_structured_ref_this_row_column_range_without_table_name() {
    let ctx = ctx_table1();

    let encoded =
        encode_rgce_with_context_ast("=[@[Qty]:[Total]]", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            1, 0, 0, 0, // table id (inferred)
            0x10, 0x00, // flags (#This Row)
            2, 0, // col_first (Qty)
            4, 0, // col_last (Total)
            0, 0, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[@[Qty]:[Total]]");
    assert_eq!(
        normalize_formula(&decoded),
        normalize_formula("[@[Qty]:[Total]]")
    );
}

#[test]
fn ast_encoder_rejects_ambiguous_tableless_structured_ref() {
    let mut ctx = ctx_table1();
    ctx.add_table(2, "Table2");
    ctx.add_table_column(2, 1, "Item");
    ctx.add_table_column(2, 2, "Qty");

    let err = encode_rgce_with_context_ast("=[@Qty]", &ctx, CellCoord::new(0, 0))
        .expect_err("expected ambiguity error");
    assert!(
        err.to_string().to_ascii_lowercase().contains("ambiguous"),
        "expected error to mention ambiguity, got: {err}"
    );
}

#[test]
fn ast_encoder_roundtrips_structured_ref_headers_column() {
    let ctx = ctx_table1();
    let encoded = encode_rgce_with_context_ast("=Table1[[#Headers],[Qty]]", &ctx, CellCoord::new(0, 0))
        .expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Table1[[#Headers],[Qty]]");
}

#[test]
fn ast_encoder_roundtrips_structured_ref_column_range() {
    let ctx = ctx_table1();
    let encoded =
        encode_rgce_with_context_ast("=Table1[[Qty]:[Total]]", &ctx, CellCoord::new(0, 0))
            .expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Table1[[Qty]:[Total]]");
}

#[test]
fn ast_encoder_roundtrips_implicit_intersection_on_structured_ref() {
    let ctx = ctx_table1();
    let encoded =
        encode_rgce_with_context_ast("=@Table1[Qty]", &ctx, CellCoord::new(0, 0)).expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@Table1[Qty]");
}

#[test]
fn ast_encoder_roundtrips_array_literal_inside_function() {
    let ctx = WorkbookContext::default();

    let encoded = encode_rgce_with_context_ast("=SUM({1,2;3,4})", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert!(!encoded.rgcb.is_empty(), "array literals must emit rgcb");

    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(decoded, "SUM({1,2;3,4})");
}

#[test]
fn ast_encoder_encodes_udf_call_via_namex() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded = encode_rgce_with_context_ast("=MyAddinFunc(1,2)", ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert!(encoded.rgcb.is_empty());

    // args..., PtgNameX(func), PtgFuncVar(argc+1, 0x00FF)
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

#[test]
fn ast_encoder_encodes_namex_reference() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded =
        encode_rgce_with_context_ast("=MyAddinFunc", ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // PtgNameX(ixti=0, nameIndex=1)
    assert_eq!(encoded.rgce, vec![0x39, 0x00, 0x00, 0x01, 0x00]);

    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "MyAddinFunc");
}

#[test]
fn ast_encoder_encodes_namex_reference_for_addin_const() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded = encode_rgce_with_context_ast(
        "='[AddIn]MyAddinConst'",
        ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert!(encoded.rgcb.is_empty());

    // PtgNameX(ixti=0, nameIndex=2)
    assert_eq!(encoded.rgce, vec![0x39, 0x00, 0x00, 0x02, 0x00]);

    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "'[AddIn]MyAddinConst'");
}

#[test]
fn ast_encoder_encodes_workbook_defined_name_function_call_via_ptgname() {
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyLambda", 1);

    let encoded =
        encode_rgce_with_context_ast("=MyLambda(1,2)", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // args..., PtgName(func), PtgFuncVar(argc+1, 0x00FF)
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x23, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, // PtgName(nameId=1)
            0x22, 0x03, 0xFF, 0x00, // PtgFuncVar(argc=3, iftab=0x00FF)
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "MyLambda(1,2)");
}

#[test]
fn ast_encoder_encodes_sheet_scoped_defined_name_function_call_via_call_expr() {
    let mut ctx = WorkbookContext::default();
    ctx.add_sheet_name("Sheet2", "MyLocalLambda", 2);

    let encoded =
        encode_rgce_with_context_ast("=Sheet2!MyLocalLambda(1,2)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert!(encoded.rgcb.is_empty());

    // args..., PtgName(func), PtgFuncVar(argc+1, 0x00FF)
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x1E, 0x02, 0x00, // 2
            0x23, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, // PtgName(nameId=2)
            0x22, 0x03, 0xFF, 0x00, // PtgFuncVar(argc=3, iftab=0x00FF)
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!MyLocalLambda(1,2)");
}

#[test]
fn ast_encoder_encodes_cell_lambda_call() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=A1(1)", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // arg..., PtgRef(callee), PtgFuncVar(argc+1, 0x00FF)
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x24, // PtgRef
            0x00, 0x00, 0x00, 0x00, // row=0
            0x00, 0xC0, // col=A with row/col relative flags
            0x22, 0x02, 0xFF, 0x00, // PtgFuncVar(argc=2, iftab=0x00FF)
        ]
    );

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "A1(1)");
}

#[test]
fn ast_encoder_encodes_absolute_cell_lambda_call() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=$A$1(1)", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // arg..., PtgRef(callee), PtgFuncVar(argc+1, 0x00FF)
    assert_eq!(
        encoded.rgce,
        vec![
            0x1E, 0x01, 0x00, // 1
            0x24, // PtgRef
            0x00, 0x00, 0x00, 0x00, // row=0
            0x00, 0x00, // col=A absolute
            0x22, 0x02, 0xFF, 0x00, // PtgFuncVar(argc=2, iftab=0x00FF)
        ]
    );

    let decoded = decode_rgce(&encoded.rgce).expect("decode");
    assert_eq!(decoded, "$A$1(1)");
}

#[test]
fn ast_encoder_encodes_column_range_as_area() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=A:C", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // `A:C` encodes as a single `PtgArea` spanning all rows.
    assert_eq!(
        encoded.rgce,
        vec![
            0x25, // PtgArea
            0x00, 0x00, 0x00, 0x00, // rowFirst=0
            0xFF, 0xFF, 0x0F, 0x00, // rowLast=1048575
            0x00, 0x80, // colFirst=A (relative column, absolute row)
            0x02, 0x80, // colLast=C (relative column, absolute row)
        ]
    );
}

#[test]
fn ast_encoder_encodes_absolute_column_range_as_area() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=$A:$C", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x25, // PtgArea
            0x00, 0x00, 0x00, 0x00, // rowFirst=0
            0xFF, 0xFF, 0x0F, 0x00, // rowLast=1048575
            0x00, 0x00, // colFirst=$A (absolute column, absolute row)
            0x02, 0x00, // colLast=$C (absolute column, absolute row)
        ]
    );
}

#[test]
fn ast_encoder_encodes_row_range_as_area() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=1:3", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    // `1:3` encodes as a single `PtgArea` spanning all columns.
    assert_eq!(
        encoded.rgce,
        vec![
            0x25, // PtgArea
            0x00, 0x00, 0x00, 0x00, // rowFirst=0 (row 1)
            0x02, 0x00, 0x00, 0x00, // rowLast=2 (row 3)
            0x00, 0x40, // colFirst=A (absolute column, relative row)
            0xFF, 0x7F, // colLast=XFD (absolute column, relative row)
        ]
    );
}

#[test]
fn ast_encoder_encodes_absolute_row_range_as_area() {
    let ctx = WorkbookContext::default();

    let encoded =
        encode_rgce_with_context_ast("=$1:$3", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x25, // PtgArea
            0x00, 0x00, 0x00, 0x00, // rowFirst=0 (row 1)
            0x02, 0x00, 0x00, 0x00, // rowLast=2 (row 3)
            0x00, 0x00, // colFirst=$A (absolute column, absolute row)
            0xFF, 0x3F, // colLast=$XFD (absolute column, absolute row)
        ]
    );
}

#[test]
fn ast_encoder_encodes_sheet_range_column_ref_as_area3d() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded =
        encode_rgce_with_context_ast("=Sheet1:Sheet3!A:A", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x3B, // PtgArea3d
            0x01, 0x00, // ixti=1
            0x00, 0x00, 0x00, 0x00, // rowFirst=0
            0xFF, 0xFF, 0x0F, 0x00, // rowLast=1048575
            0x00, 0x80, // colFirst=A (relative column, absolute row)
            0x00, 0x80, // colLast=A (relative column, absolute row)
        ]
    );
}

#[test]
fn ast_encoder_roundtrips_external_workbook_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsx", "Sheet1", "Sheet1", 0);

    let encoded = encode_rgce_with_context_ast("=[Book2.xlsx]Sheet1!A1+1", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert!(encoded.rgcb.is_empty());

    assert_eq!(
        encoded.rgce,
        vec![
            0x3A, 0x00, 0x00, // PtgRef3d(ixti=0)
            0x00, 0x00, 0x00, 0x00, // row=0
            0x00, 0xC0, // col=A with row/col relative flags
            0x1E, 0x01, 0x00, // 1
            0x03, // +
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[Book2.xlsx]Sheet1!A1+1");
}
