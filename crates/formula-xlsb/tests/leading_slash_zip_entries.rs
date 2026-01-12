use std::fs::File;
use std::io::{Cursor, Read, Write};

use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

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

fn rewrite_zip_with_leading_slash_entry_names(bytes: &[u8]) -> Vec<u8> {
    let mut input = ZipArchive::new(Cursor::new(bytes)).expect("read input zip");

    let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
    let base_options = FileOptions::<()>::default();

    for i in 0..input.len() {
        let mut entry = input.by_index(i).expect("open zip entry");
        let name = entry.name().to_string();
        let new_name = if name.starts_with('/') {
            name
        } else {
            format!("/{name}")
        };

        let mut contents = Vec::with_capacity(entry.size() as usize);
        entry.read_to_end(&mut contents).expect("read entry bytes");

        let options = base_options
            .clone()
            .compression_method(entry.compression());

        if entry.is_dir() {
            output
                .add_directory(new_name, options)
                .expect("add directory");
        } else {
            output.start_file(new_name, options).expect("start file");
            output.write_all(&contents).expect("write file");
        }
    }

    output.finish().expect("finish zip").into_inner()
}

fn rewrite_zip_with_leading_slash_entry_names_and_calc_chain(bytes: &[u8]) -> Vec<u8> {
    let mut input = ZipArchive::new(Cursor::new(bytes)).expect("read input zip");

    let mut output = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
    let base_options = FileOptions::<()>::default();

    for i in 0..input.len() {
        let mut entry = input.by_index(i).expect("open zip entry");
        let name = entry.name().to_string();
        let normalized = name.trim_start_matches('/').replace('\\', "/");

        let new_name = if name.starts_with('/') {
            name
        } else {
            format!("/{name}")
        };

        let mut contents = Vec::with_capacity(entry.size() as usize);
        entry.read_to_end(&mut contents).expect("read entry bytes");

        if normalized.eq_ignore_ascii_case("[Content_Types].xml") {
            let xml = String::from_utf8(contents).expect("[Content_Types].xml is utf8");
            let patched = xml.replace(
                "</Types>",
                r#"  <Override PartName="/xl/calcChain.bin" ContentType="application/vnd.ms-excel.calcChain"/>
</Types>"#,
            );
            contents = patched.into_bytes();
        } else if normalized.eq_ignore_ascii_case("xl/_rels/workbook.bin.rels") {
            let xml = String::from_utf8(contents).expect("workbook rels is utf8");
            let patched = xml.replace(
                "</Relationships>",
                r#"  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.bin"/>
</Relationships>"#,
            );
            contents = patched.into_bytes();
        }

        let options = base_options
            .clone()
            .compression_method(entry.compression());

        if entry.is_dir() {
            output
                .add_directory(new_name, options)
                .expect("add directory");
        } else {
            output.start_file(new_name, options).expect("start file");
            output.write_all(&contents).expect("write file");
        }
    }

    output.finish().expect("finish zip").into_inner()
}

#[test]
fn opens_xlsb_with_leading_slash_zip_entry_names_and_loads_tables() {
    let mut builder = XlsbFixtureBuilder::new();

    // Cell A1: formula token stream containing a single structured reference.
    builder.set_cell_formula_num(0, 0, 0.0, ptg_list(1, 0x0000, 2, 2, 0x18), Vec::new());

    // Provide a table definition XML part. The XLSB reader should opportunistically parse these
    // and register (table id -> name, column id -> name) mappings for structured refs, even when
    // ZIP entry names are malformed with a leading `/`.
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
    let bytes = rewrite_zip_with_leading_slash_entry_names(&bytes);

    let mut tmp = tempfile::Builder::new()
        .prefix("formula_xlsb_leading_slash_zip_entries_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    tmp.write_all(&bytes).expect("write temp xlsb");
    tmp.flush().expect("flush temp xlsb");

    let wb = XlsbWorkbook::open(tmp.path()).expect("open xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("cell");
    let formula = cell.formula.as_ref().expect("formula");

    assert_eq!(formula.text.as_deref(), Some("Table1[Qty]"));
}

#[test]
fn edited_save_rewrites_content_types_even_when_zip_entry_names_have_leading_slash() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    // Include a calcChain part and references so the edited-save path must rewrite plumbing.
    builder.add_extra_zip_part("xl/calcChain.bin", vec![0xCA, 0x1C, 0x00, 0x00]);

    let bytes = builder.build_bytes();
    let bytes = rewrite_zip_with_leading_slash_entry_names_and_calc_chain(&bytes);

    let mut src = tempfile::Builder::new()
        .prefix("formula_xlsb_leading_slash_src_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    src.write_all(&bytes).expect("write temp xlsb");
    src.flush().expect("flush temp xlsb");

    let wb = XlsbWorkbook::open(src.path()).expect("open xlsb");

    let dest = tempfile::Builder::new()
        .prefix("formula_xlsb_leading_slash_dest_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create dest xlsb");

    wb.save_with_edits(dest.path(), 0, 0, 0, 2.0)
        .expect("save with edits");

    let mut zip = ZipArchive::new(File::open(dest.path()).expect("open dest"))
        .expect("read dest zip");
    let names: Vec<String> = zip.file_names().map(str::to_string).collect();

    assert!(
        names
            .iter()
            .all(|name| name.trim_start_matches('/').to_ascii_lowercase() != "xl/calcchain.bin"),
        "expected edited save to remove calcChain entry"
    );

    let content_types_entry = names
        .iter()
        .find(|name| {
            name.trim_start_matches('/')
                .eq_ignore_ascii_case("[Content_Types].xml")
        })
        .expect("[Content_Types].xml entry exists");
    assert!(
        content_types_entry.starts_with('/'),
        "test should cover malformed ZIP entry names (expected /[Content_Types].xml)"
    );

    let mut content_types = String::new();
    zip.by_name(content_types_entry)
        .expect("open [Content_Types].xml")
        .read_to_string(&mut content_types)
        .expect("read [Content_Types].xml");
    assert!(
        !content_types.to_ascii_lowercase().contains("calcchain.bin"),
        "expected edited save to remove calcChain content type override"
    );

    let workbook_rels_entry = names
        .iter()
        .find(|name| {
            name.trim_start_matches('/')
                .eq_ignore_ascii_case("xl/_rels/workbook.bin.rels")
        })
        .expect("workbook rels entry exists");
    let mut workbook_rels = String::new();
    zip.by_name(workbook_rels_entry)
        .expect("open workbook rels")
        .read_to_string(&mut workbook_rels)
        .expect("read workbook rels");
    assert!(
        !workbook_rels.to_ascii_lowercase().contains("calcchain"),
        "expected edited save to remove calcChain relationship"
    );
}
