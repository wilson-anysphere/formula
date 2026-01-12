use formula_engine::{parse_formula, ParseOptions};
#[cfg(not(feature = "write"))]
use formula_xlsb::rgce::encode_rgce_with_context;
#[cfg(feature = "write")]
use formula_xlsb::rgce::encode_rgce_with_context_ast;
use formula_xlsb::rgce::{decode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

#[test]
fn sheet_range_3d_ref_decodes_as_quoted_prefix_and_reencodes_with_same_ixti() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 7);

    // PtgRef3d: [ptg][ixti: u16][row: u32][col: u16]
    let mut rgce = vec![0x3A];
    rgce.extend_from_slice(&7u16.to_le_bytes());
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col

    let decoded = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(decoded, "'Sheet1:Sheet3'!A1");

    // Ensure the decoded form is parseable and can be re-encoded back into a 3D ref.
    parse_formula(&format!("={decoded}"), ParseOptions::default()).expect("parse formula");

    let encoded = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };
    assert_eq!(encoded.rgce[0], 0x3A);
    assert_eq!(u16::from_le_bytes([encoded.rgce[1], encoded.rgce[2]]), 7);
}

#[test]
fn sheet_range_3d_area_decodes_as_quoted_prefix_and_reencodes_with_same_ixti() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 7);

    // PtgArea3d: [ptg][ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
    let mut rgce = vec![0x3B];
    rgce.extend_from_slice(&7u16.to_le_bytes());
    rgce.extend_from_slice(&0u32.to_le_bytes()); // rowFirst = 0 (A1)
    rgce.extend_from_slice(&1u32.to_le_bytes()); // rowLast = 1 (A2)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst = A, relative row/col
    rgce.extend_from_slice(&0xC001u16.to_le_bytes()); // colLast = B, relative row/col

    let decoded = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(decoded, "'Sheet1:Sheet3'!A1:B2");

    parse_formula(&format!("={decoded}"), ParseOptions::default()).expect("parse formula");

    let encoded = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };

    assert_eq!(encoded.rgce[0], 0x3B);
    assert_eq!(u16::from_le_bytes([encoded.rgce[1], encoded.rgce[2]]), 7);
}

#[test]
fn external_workbook_sheet_range_3d_ref_decodes_as_quoted_prefix_and_reencodes_with_same_ixti() {
    let mut ctx = WorkbookContext::default();
    // Encoding uses the `[Book]Sheet` form (AST encoder expands external workbook refs this way).
    ctx.add_extern_sheet("[Book2.xlsb]SheetA", "[Book2.xlsb]SheetB", 9);
    // Decoding uses the workbook + sheet span split from the ExternSheet/SupBook tables.
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 9);

    // PtgRef3d: [ptg][ixti: u16][row: u32][col: u16]
    let mut rgce = vec![0x3A];
    rgce.extend_from_slice(&9u16.to_le_bytes());
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col

    let decoded = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(decoded, "'[Book2.xlsb]SheetA:SheetB'!A1");

    parse_formula(&format!("={decoded}"), ParseOptions::default()).expect("parse formula");

    let encoded = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };
    assert_eq!(encoded.rgce[0], 0x3A);
    assert_eq!(u16::from_le_bytes([encoded.rgce[1], encoded.rgce[2]]), 9);
}

#[test]
fn external_workbook_sheet_range_3d_area_decodes_as_quoted_prefix_and_reencodes_with_same_ixti() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("[Book2.xlsb]SheetA", "[Book2.xlsb]SheetB", 9);
    ctx.add_extern_sheet_external_workbook("Book2.xlsb", "SheetA", "SheetB", 9);

    // PtgArea3d: [ptg][ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
    let mut rgce = vec![0x3B];
    rgce.extend_from_slice(&9u16.to_le_bytes());
    rgce.extend_from_slice(&0u32.to_le_bytes()); // rowFirst = 0 (A1)
    rgce.extend_from_slice(&1u32.to_le_bytes()); // rowLast = 1 (A2)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst = A, relative row/col
    rgce.extend_from_slice(&0xC001u16.to_le_bytes()); // colLast = B, relative row/col

    let decoded = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(decoded, "'[Book2.xlsb]SheetA:SheetB'!A1:B2");

    parse_formula(&format!("={decoded}"), ParseOptions::default()).expect("parse formula");

    let encoded = {
        #[cfg(feature = "write")]
        {
            encode_rgce_with_context_ast(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
        #[cfg(not(feature = "write"))]
        {
            encode_rgce_with_context(&format!("={decoded}"), &ctx, CellCoord::new(0, 0))
                .expect("encode")
        }
    };

    assert_eq!(encoded.rgce[0], 0x3B);
    assert_eq!(u16::from_le_bytes([encoded.rgce[1], encoded.rgce[2]]), 9);
}
