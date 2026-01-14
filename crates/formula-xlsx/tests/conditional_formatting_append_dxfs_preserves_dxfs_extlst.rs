use std::io::{Cursor, Write};

use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color};
use formula_xlsx::{load_from_bytes, XlsxPackage};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn build_minimal_xlsx_with_styles_and_sheet(sheet1_xml: &str, styles_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/styles.xml", styles_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/worksheets/sheet1.xml",
        sheet1_xml.as_bytes(),
    );

    zip.finish().unwrap().into_inner()
}

fn extract_cf_rule_dxf_ids(sheet_xml: &str) -> Vec<Option<u32>> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    doc.descendants()
        .filter(|n| {
            n.is_element()
                && n.tag_name().name() == "cfRule"
                && n.tag_name().namespace() == Some(MAIN_NS)
        })
        .map(|n| n.attribute("dxfId").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

#[test]
fn xlsxdocument_appends_dxfs_before_extlst_in_styles_dxfs() {
    // styles.xml contains a `<dxfs>` table with an `<extLst>` child. Per the OOXML schema, extLst
    // must come *after* all `<dxf>` children. When we append new dxfs during a round-trip edit, we
    // should insert them before `<extLst>` (not after).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="1">
    <dxf>
      <fill>
        <patternFill patternType="solid">
          <fgColor rgb="FFFF0000"/>
          <bgColor indexed="64"/>
        </patternFill>
      </fill>
    </dxf>
    <extLst>
      <ext uri="{DEADBEEF-DEAD-BEEF-DEAD-BEEFDEADBEEF}">
        <dummy xmlns="urn:dummy" val="1"/>
      </ext>
    </extLst>
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"#;

    // Ensure the base workbook loads conditional formatting so Sheet1’s local dxfs vector aligns
    // with the global table.
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <conditionalFormatting sqref="A1">
    <cfRule type="expression" priority="1" dxfId="0">
      <formula>A1&gt;0</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let bytes = build_minimal_xlsx_with_styles_and_sheet(sheet1_xml, styles_xml);
    let mut doc = load_from_bytes(&bytes).expect("load minimal workbook via XlsxDocument");

    // Add Sheet2 with a new CF Dxf (blue fill).
    let sheet2_id = doc.workbook.add_sheet("Sheet2").expect("add sheet2");
    let dxf_blue = CfStyleOverride {
        fill: Some(Color::new_argb(0xFF0000FF)),
        ..Default::default()
    };
    let rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![parse_range_a1("A1").unwrap()],
        dxf_id: Some(0), // worksheet-local; should be remapped to global index 1.
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "A1>0".to_string(),
        },
        dependencies: vec![],
    }];
    doc.workbook
        .sheet_mut(sheet2_id)
        .unwrap()
        .set_conditional_formatting(rules, vec![dxf_blue]);

    let saved = doc.save_to_vec().expect("save modified workbook");
    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved package");

    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let doc = Document::parse(styles_xml).expect("parse styles.xml");

    let dxfs = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dxfs" && n.tag_name().namespace() == Some(MAIN_NS))
        .expect("dxfs element");

    let children: Vec<&str> = dxfs
        .children()
        .filter(|n| n.is_element() && n.tag_name().namespace() == Some(MAIN_NS))
        .map(|n| n.tag_name().name())
        .collect();
    assert_eq!(
        children,
        vec!["dxf", "dxf", "extLst"],
        "expected new dxf to be inserted before extLst"
    );

    let ext = dxfs
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "ext" && n.tag_name().namespace() == Some(MAIN_NS))
        .expect("ext element should be preserved");
    assert_eq!(
        ext.attribute("uri"),
        Some("{DEADBEEF-DEAD-BEEF-DEAD-BEEFDEADBEEF}"),
        "expected extLst contents to be preserved"
    );

    let dummy = dxfs
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dummy" && n.tag_name().namespace() == Some("urn:dummy"))
        .expect("dummy extension node should be preserved");
    assert_eq!(dummy.attribute("val"), Some("1"));

    let sheet2_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet2.xml").unwrap()).unwrap();
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1)]);
}
