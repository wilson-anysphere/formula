#![cfg(feature = "write")]

use std::io::Write;

use formula_xlsb::rgce::{
    decode_rgce_with_context, encode_rgce_with_context_ast_in_sheet, CellCoord,
};
use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

#[test]
fn encodes_tableless_this_row_structured_ref_using_sheet_context() {
    let mut builder = XlsbFixtureBuilder::new();

    // Provide an Office-style table definition XML part.
    let table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1"
       name="Table1"
       displayName="Table1"
       ref="A1:B10">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;
    builder.add_extra_zip_part("xl/tables/table1.xml", table_xml.as_bytes().to_vec());

    // Associate the table with Sheet1 via the worksheet relationships part.
    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>
"#;
    builder.add_extra_zip_part(
        "xl/worksheets/_rels/sheet1.bin.rels",
        sheet_rels.as_bytes().to_vec(),
    );

    let bytes = builder.build_bytes();
    let mut tmp = tempfile::Builder::new()
        .prefix("formula_xlsb_tableless_structured_ref_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    tmp.write_all(&bytes).expect("write temp xlsb");

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let ctx = wb.workbook_context();

    // Table range should be registered for inference.
    assert_eq!(ctx.table_id_for_cell("Sheet1", 1, 0), Some(1));

    let base = CellCoord::new(1, 0); // A2 (inside the table range A1:B10)
    let encoded =
        encode_rgce_with_context_ast_in_sheet("=[@Qty]", ctx, "Sheet1", base).expect("encode");

    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "[@Qty]");
}

