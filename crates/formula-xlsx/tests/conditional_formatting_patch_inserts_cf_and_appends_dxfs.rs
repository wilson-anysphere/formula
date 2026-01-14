use std::io::{Cursor, Read, Write};

use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color};
use formula_xlsx::{load_from_bytes, XlsxPackage};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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

fn zip_part(zip_bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = String::new();
    file.read_to_string(&mut buf).expect("read part");
    buf
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
fn xlsxdocument_inserts_conditional_formatting_into_existing_sheet_and_appends_dxfs() {
    // Base styles.xml contains one `<dxf>` with an unmodeled element (`<alignment>`). This test
    // ensures that:
    // - Adding conditional formatting to an *existing* sheet (i.e. patcher path) inserts
    //   `<conditionalFormatting>` blocks.
    // - Newly introduced DXFs are appended to the existing styles `<dxfs>` table, without dropping
    //   the unknown `<alignment>` child.
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
      <alignment horizontal="center"/>
    </dxf>
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"#;

    // IMPORTANT: no `<conditionalFormatting>` in the original sheet XML so the write path must
    // inject it when we add rules.
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx_with_styles_and_sheet(sheet1_xml, styles_xml);
    let mut doc = load_from_bytes(&bytes).expect("load minimal workbook via XlsxDocument");

    let sheet1_id = doc.workbook.sheets[0].id;

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
        .sheet_mut(sheet1_id)
        .unwrap()
        .set_conditional_formatting(rules, vec![dxf_blue]);

    let saved = doc.save_to_vec().expect("save modified workbook");

    // Verify sheet XML contains injected conditional formatting and dxfId is mapped to the
    // appended global entry (base has 1 dxf, new is index 1).
    let saved_sheet1_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    assert!(
        saved_sheet1_xml.contains("conditionalFormatting"),
        "expected conditional formatting to be injected into existing worksheet XML"
    );
    assert_eq!(extract_cf_rule_dxf_ids(&saved_sheet1_xml), vec![Some(1)]);

    // Verify styles.xml preserved the unknown `<alignment>` and appended the new dxf.
    let pkg = XlsxPackage::from_bytes(&saved).expect("open saved package");
    let saved_styles_xml =
        std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).expect("styles.xml utf-8");
    let parsed_styles = Document::parse(saved_styles_xml).expect("parse styles.xml");
    let dxfs = parsed_styles
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dxfs" && n.tag_name().namespace() == Some(MAIN_NS))
        .expect("dxfs element");
    let dxfs: Vec<_> = dxfs
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "dxf" && n.tag_name().namespace() == Some(MAIN_NS))
        .collect();
    assert_eq!(dxfs.len(), 2, "expected new dxf to be appended");

    let alignment = dxfs[0]
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "alignment")
        .expect("expected alignment element to be preserved");
    assert_eq!(alignment.attribute("horizontal"), Some("center"));

    let fg_color = dxfs[1]
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "fgColor")
        .and_then(|n| n.attribute("rgb"));
    assert_eq!(
        fg_color,
        Some("FF0000FF"),
        "expected appended dxf to be the blue fill"
    );
}

