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

#[test]
fn encodes_tableless_this_row_structured_ref_using_sheet_context_in_multi_table_workbook() {
    let mut builder = XlsbFixtureBuilder::new();

    // Provide two Office-style table definition XML parts.
    let table1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    builder.add_extra_zip_part("xl/tables/table1.xml", table1_xml.as_bytes().to_vec());

    let table2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2"
       name="Table2"
       displayName="Table2"
       ref="D1:E10">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;
    builder.add_extra_zip_part("xl/tables/table2.xml", table2_xml.as_bytes().to_vec());

    // Associate both tables with Sheet1 via the worksheet relationships part.
    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
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

    // Pick a base cell inside the Table2 range `D1:E10`.
    let base = CellCoord::new(1, 3); // D2 (inside the table range D1:E10)
    assert_eq!(ctx.table_id_for_cell("Sheet1", base.row, base.col), Some(2));

    let encoded =
        encode_rgce_with_context_ast_in_sheet("=[@Qty]", ctx, "Sheet1", base).expect("encode");
    assert_eq!(
        encoded.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            2, 0, 0, 0, // table id (inferred by sheet+cell)
            0x10, 0x00, // flags (#This Row)
            2, 0, // col_first (Qty)
            2, 0, // col_last (Qty)
            0, 0, // reserved
        ]
    );

    let decoded = decode_rgce_with_context(&encoded.rgce, ctx).expect("decode");
    assert_eq!(decoded, "[@Qty]");
}

#[test]
fn rejects_tableless_this_row_structured_ref_outside_tables_using_sheet_context_in_multi_table_workbook(
) {
    let mut builder = XlsbFixtureBuilder::new();

    // Same two tables as the test above.
    let table1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    builder.add_extra_zip_part("xl/tables/table1.xml", table1_xml.as_bytes().to_vec());

    let table2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="2"
       name="Table2"
       displayName="Table2"
       ref="D1:E10">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;
    builder.add_extra_zip_part("xl/tables/table2.xml", table2_xml.as_bytes().to_vec());

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
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

    // Pick a base cell outside both table ranges.
    let base = CellCoord::new(20, 0); // A21
    assert_eq!(ctx.single_table_id(), None, "workbook should have >1 table");
    assert_eq!(ctx.table_id_for_cell("Sheet1", base.row, base.col), None);

    let err = encode_rgce_with_context_ast_in_sheet("=[@Qty]", ctx, "Sheet1", base)
        .expect_err("expected structured-ref inference error");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("cannot infer table") || msg.contains("inside exactly one table"),
        "unexpected error: {err}"
    );
}
