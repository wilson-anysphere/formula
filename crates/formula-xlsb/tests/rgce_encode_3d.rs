use formula_xlsb::rgce::{
    decode_rgce_with_context, encode_rgce_with_context, CellCoord, EncodeError,
};
use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

fn normalize(formula: &str) -> String {
    let ast = formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula");
    ast.to_string(formula_engine::SerializeOptions {
        omit_equals: true,
        ..Default::default()
    })
    .expect("serialize formula")
}

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
fn encodes_and_decodes_sheet_qualified_reordered_area_and_preserves_absolute_flags() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context("=Sheet2!B$1:$A2", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!$A$1:B2");
}

#[test]
fn encodes_and_decodes_sheet_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted = encode_rgce_with_context("=SUM(Sheet1:Sheet3!A1)", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let encoded_quoted = encode_rgce_with_context("=SUM('Sheet1:Sheet3'!A1)", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('Sheet1:Sheet3'!A1)");
}

#[test]
fn encodes_and_decodes_sheet_qualified_column_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded = encode_rgce_with_context("=SUM(Sheet2!A:A)", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM(Sheet2!A:A)");
}

#[test]
fn encodes_and_decodes_sheet_qualified_row_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded = encode_rgce_with_context("=SUM(Sheet2!1:1)", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM(Sheet2!1:1)");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_sheet_qualified_column_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context("=@Sheet2!A:A", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@Sheet2!A:A");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_sheet_qualified_row_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 0);

    let encoded =
        encode_rgce_with_context("=@Sheet2!1:1", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@Sheet2!1:1");
}

#[test]
fn encodes_and_decodes_sheet_range_column_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted =
        encode_rgce_with_context("=SUM(Sheet1:Sheet3!A:A)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let encoded_quoted =
        encode_rgce_with_context("=SUM('Sheet1:Sheet3'!A:A)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('Sheet1:Sheet3'!A:A)");
}

#[test]
fn encodes_and_decodes_sheet_range_row_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted =
        encode_rgce_with_context("=SUM(Sheet1:Sheet3!1:1)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let encoded_quoted =
        encode_rgce_with_context("=SUM('Sheet1:Sheet3'!1:1)", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('Sheet1:Sheet3'!1:1)");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_sheet_range_column_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted =
        encode_rgce_with_context("=@Sheet1:Sheet3!A:A", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let encoded_quoted =
        encode_rgce_with_context("=@'Sheet1:Sheet3'!A:A", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@'Sheet1:Sheet3'!A:A");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_sheet_range_row_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 1);

    let encoded_unquoted =
        encode_rgce_with_context("=@Sheet1:Sheet3!1:1", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let encoded_quoted =
        encode_rgce_with_context("=@'Sheet1:Sheet3'!1:1", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@'Sheet1:Sheet3'!1:1");
}

#[test]
fn encodes_and_decodes_external_workbook_sheet_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    // Excel writes external workbook 3D spans as `[Book]SheetA:SheetB!A1`, but the rgce decoder
    // emits a single quoted identifier (`'[Book]SheetA:SheetB'!A1`) so the prefix is a single
    // token for formula-engine.
    let encoded_unquoted = encode_rgce_with_context(
        "=SUM([Book2.xlsb]SheetA:SheetB!A1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    let encoded_quoted = encode_rgce_with_context(
        "=SUM('[Book2.xlsb]SheetA:SheetB'!A1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('[Book2.xlsb]SheetA:SheetB'!A1)");
}

#[test]
fn encodes_and_decodes_external_workbook_sheet_range_column_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    let encoded_unquoted = encode_rgce_with_context(
        "=SUM([Book2.xlsb]SheetA:SheetB!A:A)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    let encoded_quoted = encode_rgce_with_context(
        "=SUM('[Book2.xlsb]SheetA:SheetB'!A:A)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('[Book2.xlsb]SheetA:SheetB'!A:A)");
}

#[test]
fn encodes_and_decodes_external_workbook_sheet_range_row_range_ref_in_function() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    let encoded_unquoted = encode_rgce_with_context(
        "=SUM([Book2.xlsb]SheetA:SheetB!1:1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    let encoded_quoted = encode_rgce_with_context(
        "=SUM('[Book2.xlsb]SheetA:SheetB'!1:1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('[Book2.xlsb]SheetA:SheetB'!1:1)");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_external_workbook_sheet_range_column_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    let encoded_unquoted = encode_rgce_with_context(
        "=@[Book2.xlsb]SheetA:SheetB!A:A",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    let encoded_quoted = encode_rgce_with_context(
        "=@'[Book2.xlsb]SheetA:SheetB'!A:A",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@'[Book2.xlsb]SheetA:SheetB'!A:A");
}

#[test]
fn encodes_and_decodes_implicit_intersection_on_external_workbook_sheet_range_row_range() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 0);

    let encoded_unquoted = encode_rgce_with_context(
        "=@[Book2.xlsb]SheetA:SheetB!1:1",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    let encoded_quoted = encode_rgce_with_context(
        "=@'[Book2.xlsb]SheetA:SheetB'!1:1",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");
    assert_eq!(encoded_unquoted.rgce, encoded_quoted.rgce);

    let decoded = decode_rgce_with_context(&encoded_unquoted.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "@'[Book2.xlsb]SheetA:SheetB'!1:1");
}

#[test]
fn encodes_and_decodes_external_workbook_ref() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "Sheet1", "Sheet1", 0);

    let encoded = encode_rgce_with_context("=[Book2.xlsb]Sheet1!A1+1", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[Book2.xlsb]Sheet1!A1+1");
}

#[test]
fn encodes_and_decodes_external_workbook_ref_with_workbook_name_containing_lbracket() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("A1[Name.xlsb", "Sheet1", "Sheet1", 0);

    let encoded =
        encode_rgce_with_context("=[A1[Name.xlsb]Sheet1!A1+1", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[A1[Name.xlsb]Sheet1!A1+1");
    formula_engine::parse_formula(&format!("={decoded}"), Default::default())
        .expect("should parse");
}

#[test]
fn encodes_and_decodes_external_workbook_ref_with_workbook_name_containing_lbracket_and_escaped_rbracket(
) {
    let mut ctx = WorkbookContext::default();
    // Workbook name contains a literal `]`. Excel escapes this as `]]` inside the `[...]` prefix.
    ctx.add_extern_sheet_external_workbook("Book[Name].xlsb", "Sheet1", "Sheet1", 0);

    let encoded =
        encode_rgce_with_context("=[Book[Name]].xlsb]Sheet1!A1+1", &ctx, CellCoord::new(0, 0))
            .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "[Book[Name]].xlsb]Sheet1!A1+1");
    formula_engine::parse_formula(&format!("={decoded}"), Default::default())
        .expect("should parse");
}

#[test]
fn encodes_and_decodes_external_workbook_sheet_range_ref_with_escaped_workbook_rbracket() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet_external_workbook("Book[Name].xlsb", "SheetA", "SheetB", 0);

    // Excel writes external workbook 3D spans as `[Book]SheetA:SheetB!A1`, but the rgce decoder
    // emits a single quoted identifier (`'[Book]SheetA:SheetB'!A1`) so the prefix is a single
    // token for formula-engine.
    let encoded = encode_rgce_with_context(
        "=SUM([Book[Name]].xlsb]SheetA:SheetB!A1)",
        &ctx,
        CellCoord::new(0, 0),
    )
    .expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "SUM('[Book[Name]].xlsb]SheetA:SheetB'!A1)");
    formula_engine::parse_formula(&format!("={decoded}"), Default::default())
        .expect("should parse");
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
fn encodes_and_decodes_sheet_scoped_name() {
    let mut ctx = WorkbookContext::default();
    ctx.add_sheet_name("Sheet2", "MyLocalName", 2);

    let encoded = encode_rgce_with_context("=Sheet2!MyLocalName", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert_eq!(
        encoded.rgce,
        vec![
            0x23, // PtgName
            0x02, 0x00, 0x00, 0x00, // nameId=2
            0x00, 0x00, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(decoded, "Sheet2!MyLocalName");
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
fn encodes_and_decodes_if_with_missing_arg() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=IF(,1,0)", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("=IF(,1,0)"), normalize(&decoded));
}

#[test]
fn encodes_and_decodes_discount_securities_and_tbill_functions_via_ftab() {
    let ctx = WorkbookContext::default();
    for formula in [
        "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,1)",
        "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100)",
        "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,2)",
        "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,)",
        "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,3)",
        "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,0)",
        "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05)",
        "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,0)",
        "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,)",
        "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
        "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,0)",
        "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,)",
        "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077)",
        "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,0)",
        "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,)",
        "=TBILLPRICE(DATE(2020,1,1),DATE(2020,7,1),0.05)",
        "=TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),97.47222222222223)",
        "=TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),0.05)",
    ] {
        let encoded =
            encode_rgce_with_context(formula, &ctx, CellCoord::new(0, 0)).expect("encode");
        let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
        assert_eq!(normalize(formula), normalize(&decoded));
    }
}

#[test]
fn encodes_and_decodes_modern_error_literals() {
    let ctx = WorkbookContext::default();

    for (code, lit) in [
        (0x2C, "#SPILL!"),
        (0x2D, "#CALC!"),
        (0x2E, "#FIELD!"),
        (0x2F, "#CONNECT!"),
        (0x30, "#BLOCKED!"),
        (0x31, "#UNKNOWN!"),
    ] {
        let formula = format!("={lit}");
        let encoded =
            encode_rgce_with_context(&formula, &ctx, CellCoord::new(0, 0)).expect("encode");
        assert_eq!(encoded.rgce, vec![0x1C, code], "encode {lit}");
        assert!(encoded.rgcb.is_empty(), "encode {lit} should not emit rgcb");

        let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
        assert_eq!(decoded, lit, "decode code={code:#04x}");
    }
}

#[test]
fn encodes_addin_udf_calls_via_namex() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");
    let ctx = wb.workbook_context();

    let encoded =
        encode_rgce_with_context("=MyAddinFunc(1,2)", ctx, CellCoord::new(0, 0)).expect("encode");
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
