use std::io::{Cursor, Write};

use formula_model::rich_text::{RichText, RichTextRunStyle};
use formula_model::{CellRef, CellValue};

use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_shared_strings(sheet_xml: &str, shared_strings_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(shared_strings_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn apply_cell_patches_preserves_cell_attrs_and_extlst_when_updating_value() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" cm="7" customAttr="x"><v>1</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"cm="7""#),
        "expected cm attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(r#"customAttr="x""#),
        "expected customAttr attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<v>2</v>"),
        "expected cached value to update, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_preserves_vm_attr_but_preserves_extlst_when_updating_value() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" cm="7" vm="7" customAttr="x"><v>1</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"vm="7""#),
        "expected vm attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(r#"customAttr="x""#),
        "expected customAttr attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<v>2</v>"),
        "expected cached value to update, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_preserves_vm_attr_but_preserves_extlst_when_updating_formula_and_value() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" vm="9" customAttr="x"><v>1</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(4.0), "=2+2"),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"vm="9""#),
        "expected vm attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(r#"customAttr="x""#),
        "expected customAttr attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );

    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(cell.attribute("vm"), Some("9"));
    assert_eq!(cell.attribute("customAttr"), Some("x"));

    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .expect("expected <f>");
    assert_eq!(f.text().unwrap_or_default(), "2+2");

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v>");
    assert_eq!(v.text().unwrap_or_default(), "4");
}

#[test]
fn apply_cell_patches_preserves_non_formula_children_when_updating_formula() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" cm="7" customAttr="x"><f aca="1">1+1</f><v>2</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(4.0), "=2+2"),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );

    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .expect("expected <f>");
    assert_eq!(f.attribute("aca"), Some("1"));
    assert_eq!(f.text().unwrap_or_default(), "2+2");

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v>");
    assert_eq!(v.text().unwrap_or_default(), "4");
}

#[test]
fn apply_cell_patches_preserves_vm_attr_when_updating_formula_and_value() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" cm="7" vm="7"><f aca="1">1+1</f><v>2</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(4.0), "=2+2"),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"vm="7""#),
        "expected vm attribute to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );

    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(cell.attribute("vm"), Some("7"));

    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .expect("expected <f>");
    assert_eq!(f.attribute("aca"), Some("1"));
    assert_eq!(f.text().unwrap_or_default(), "2+2");

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v>");
    assert_eq!(v.text().unwrap_or_default(), "4");
}

#[test]
fn apply_cell_patches_preserves_formula_attrs_when_formula_follows_value() {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    // Some generators emit `<v>` before `<f>`. Ensure we still preserve `<f>` attributes while
    // inserting the patched formula before the value.
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" s="5" cm="7" customAttr="x"><v>2</v><f aca="1">1+1</f>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(4.0), "=2+2"),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(extlst),
        "expected extLst subtree to be preserved, got: {out_xml}"
    );

    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    let f = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "f")
        .expect("expected <f>");
    assert_eq!(f.attribute("aca"), Some("1"));
    assert_eq!(f.text().unwrap_or_default(), "2+2");

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v>");
    assert_eq!(v.text().unwrap_or_default(), "4");
}

#[test]
fn apply_cell_patches_does_not_confuse_namespaced_s_attr_with_cell_style() {
    // The patcher should only treat the unprefixed `s` attribute as the cell style index.
    // Namespaced attributes like `x:s="..."` must be preserved but not interpreted as style.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:x="http://example.com">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" x:s="7"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_style(CellValue::Number(1.0), 7),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"x:s="7""#),
        "expected namespaced x:s attr to be preserved, got: {out_xml}"
    );
    assert!(
        out_xml.contains(r#" s="7""#) || out_xml.contains(r#"s="7""#),
        "expected unprefixed s attr to be written by style patch, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_inserts_new_cells_before_row_extlst() {
    let row_extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">row</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c>{row_extlst}</row>
  </sheetData>
</worksheet>"#
    );

    let bytes = build_minimal_xlsx(&worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(row_extlst),
        "expected row extLst subtree to be preserved, got: {out_xml}"
    );

    let b1_pos = out_xml
        .find(r#"r="B1""#)
        .expect("expected patched B1 cell");
    let ext_pos = out_xml.find("<extLst").expect("expected row extLst");
    assert!(
        b1_pos < ext_pos,
        "expected inserted cells to appear before row extLst, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_uses_str_type_for_formula_string_results_without_touching_shared_strings() {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"/>
  </sheetData>
</worksheet>"#;

    let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>existing</t></si>
</sst>"#;

    let bytes = build_minimal_xlsx_with_shared_strings(worksheet_xml, shared_strings_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let original_ss = pkg.part("xl/sharedStrings.xml").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::String("hi".to_string()), r#"="hi""#),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    assert_eq!(
        pkg.part("xl/sharedStrings.xml").unwrap(),
        original_ss.as_slice(),
        "expected sharedStrings.xml to remain unchanged for formula string results"
    );

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(cell.attribute("t"), Some("str"));
    assert!(
        cell.children()
            .any(|n| n.is_element() && n.tag_name().name() == "v" && n.text() == Some("hi")),
        "expected cached string result in <v>, got: {out_xml}"
    );
    assert!(
        !cell
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "is"),
        "expected no inline string <is> when using t=\"str\", got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_writes_inline_rich_text_when_shared_strings_absent() {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"/>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let rich = RichText::from_segments(vec![
        ("Hi ".to_string(), RichTextRunStyle::default()),
        (
            "World".to_string(),
            RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ]);

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::RichText(rich)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(out_xml).expect("parse worksheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(cell.attribute("t"), Some("inlineStr"));

    let is = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "is")
        .expect("expected inline string <is>");

    let runs: Vec<_> = is
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
        .collect();
    assert_eq!(
        runs.len(),
        2,
        "expected two rich text runs inside <is>, got: {out_xml}"
    );

    assert!(
        is.descendants()
            .any(|n| n.is_element() && n.tag_name().name() == "b"),
        "expected <b> element for bold rich text run, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_preserves_prefixes_when_expanding_empty_sheetdata() {
    // Some producers namespace SpreadsheetML using explicit prefixes instead of a default namespace.
    // Ensure we preserve the prefix when expanding `<x:sheetData/>` into a non-empty sheetData and
    // when inserting new rows/cells.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:dimension ref="A1"/>
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    roxmltree::Document::parse(out_xml).expect("output xml should be well-formed");
    assert!(
        out_xml.contains("<x:sheetData>"),
        "expected prefixed sheetData start tag, got: {out_xml}"
    );
    assert!(
        out_xml.contains("</x:sheetData>"),
        "expected prefixed sheetData end tag, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<x:row"),
        "expected prefixed row insertion, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<x:c r=\"A1\""),
        "expected prefixed cell insertion, got: {out_xml}"
    );
}

#[test]
fn apply_cell_patches_preserves_prefixes_when_inserting_missing_sheetdata() {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:dimension ref="A1"/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    roxmltree::Document::parse(out_xml).expect("output xml should be well-formed");
    assert!(
        out_xml.contains("<x:sheetData>"),
        "expected prefixed sheetData insertion, got: {out_xml}"
    );
    assert!(
        out_xml.contains("</x:sheetData>"),
        "expected prefixed sheetData end tag, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<x:row"),
        "expected prefixed row insertion, got: {out_xml}"
    );
    assert!(
        out_xml.contains("<x:c r=\"A1\""),
        "expected prefixed cell insertion, got: {out_xml}"
    );
}
