#![cfg(feature = "write")]

use std::io::Write;

use formula_xlsb::rgce::{
    decode_rgce_with_context, encode_rgce_with_context_ast_in_sheet, CellCoord,
};
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn table_id_from_ptg_list_rgce(rgce: &[u8]) -> u32 {
    assert!(
        rgce.len() >= 6,
        "expected PtgList rgce with at least 6 bytes, got {}",
        rgce.len()
    );
    assert_eq!(rgce[0], 0x18, "expected PtgExtend (0x18)");
    assert_eq!(rgce[1], 0x19, "expected etpg=PtgList (0x19)");
    u32::from_le_bytes(rgce[2..6].try_into().expect("table id bytes"))
}

#[test]
fn infers_table_for_tableless_structured_ref_by_sheet_and_base_cell() {
    let mut builder = XlsbFixtureBuilder::new();

    let table1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1"
       name="Table1"
       displayName="Table1"
       ref="A1:B3">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;
    let table2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2"
       name="Table2"
       displayName="Table2"
       ref="D5:E7">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;

    builder.add_extra_zip_part("xl/tables/table1.xml", table1_xml.as_bytes().to_vec());
    builder.add_extra_zip_part("xl/tables/table2.xml", table2_xml.as_bytes().to_vec());

    // `sheet1.bin.rels` links the table parts to the owning worksheet.
    let sheet1_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
</Relationships>
"#;
    builder.add_extra_zip_part(
        "xl/worksheets/_rels/sheet1.bin.rels",
        sheet1_rels.as_bytes().to_vec(),
    );

    let bytes = builder.build_bytes();
    let mut tmp = tempfile::Builder::new()
        .prefix("formula_xlsb_table_infer_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    tmp.write_all(&bytes).expect("write temp xlsb");

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let ctx = wb.workbook_context();

    // WorkbookContext sheet-aware lookup.
    assert_eq!(ctx.table_id_for_cell("Sheet1", 0, 0), Some(1)); // A1
    assert_eq!(ctx.table_id_for_cell("Sheet1", 1, 0), Some(1)); // A2
    assert_eq!(ctx.table_id_for_cell("Sheet1", 2, 1), Some(1)); // B3
    assert_eq!(ctx.table_id_for_cell("Sheet1", 0, 2), None); // C1 (outside Table1)

    assert_eq!(ctx.table_id_for_cell("Sheet1", 4, 3), Some(2)); // D5
    assert_eq!(ctx.table_id_for_cell("Sheet1", 5, 3), Some(2)); // D6
    assert_eq!(ctx.table_id_for_cell("Sheet1", 6, 4), Some(2)); // E7
    assert_eq!(ctx.table_id_for_cell("Sheet1", 4, 2), None); // C5 (outside Table2)

    // Encode `[@Qty]` from a base cell within each table range. The table id should be inferred
    // from sheet + base cell position.
    let base_table1 = CellCoord::new(1, 0); // A2
    let encoded1 =
        encode_rgce_with_context_ast_in_sheet("=[@Qty]", ctx, "Sheet1", base_table1).expect("encode");
    assert!(encoded1.rgcb.is_empty());
    assert_eq!(table_id_from_ptg_list_rgce(&encoded1.rgce), 1);
    assert_eq!(
        decode_rgce_with_context(&encoded1.rgce, ctx).expect("decode"),
        "[@Qty]"
    );

    let base_table2 = CellCoord::new(5, 3); // D6
    let encoded2 =
        encode_rgce_with_context_ast_in_sheet("=[@Qty]", ctx, "Sheet1", base_table2).expect("encode");
    assert!(encoded2.rgcb.is_empty());
    assert_eq!(table_id_from_ptg_list_rgce(&encoded2.rgce), 2);
    assert_eq!(
        decode_rgce_with_context(&encoded2.rgce, ctx).expect("decode"),
        "[@Qty]"
    );
}
