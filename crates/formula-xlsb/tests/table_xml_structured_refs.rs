use std::io::Write;

use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn ptg_list(table_id: u32, flags: u16, col_first: u16, col_last: u16, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.push(0x19); // etpg=0x19 (PtgList / structured ref)
    out.extend_from_slice(&table_id.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out
}

#[test]
fn loads_table_xml_and_decodes_structured_refs_with_real_names() {
    let mut builder = XlsbFixtureBuilder::new();

    // Cell A1: formula token stream containing a single structured reference.
    builder.set_cell_formula_num(0, 0, 0.0, ptg_list(1, 0x0000, 2, 2, 0x18), Vec::new());

    // Provide an Office-style table definition XML part. The XLSB reader should opportunistically
    // parse these and register (table id -> name, column id -> name) mappings for structured refs.
    let table_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1"
       name="Table1"
       displayName="Table1"
       ref="A1:B2">
  <tableColumns count="2">
    <tableColumn id="1" name="Item"/>
    <tableColumn id="2" name="Qty"/>
  </tableColumns>
</table>
"#;
    builder.add_extra_zip_part("xl/tables/table1.xml", table_xml.as_bytes().to_vec());

    let bytes = builder.build_bytes();
    let mut tmp = tempfile::Builder::new()
        .prefix("formula_xlsb_table_xml_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    tmp.write_all(&bytes).expect("write temp xlsb");

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = sheet.cells.iter().find(|c| c.row == 0 && c.col == 0).expect("cell");
    let formula = cell.formula.as_ref().expect("formula");

    assert_eq!(formula.text.as_deref(), Some("Table1[Qty]"));
}

