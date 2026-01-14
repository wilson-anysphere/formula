use formula_model::{
    format_render_plan, parse_range_a1, CellRef, CellValue, CellValueProvider, DataBarDirection,
    ConditionalFormattingEngine,
};
use formula_xlsx::{parse_worksheet_conditional_formatting, DxfProvider, Styles, XlsxPackage};
use roxmltree::Document;
use std::collections::HashMap;

struct SheetValues {
    values: HashMap<CellRef, CellValue>,
}

impl CellValueProvider for SheetValues {
    fn get_value(&self, cell: CellRef) -> Option<CellValue> {
        self.values.get(&cell).cloned()
    }
}

fn parse_sheet_values(xml: &str) -> SheetValues {
    let doc = Document::parse(xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let mut values = HashMap::new();
    for cell in doc.descendants().filter(|n| {
        n.is_element() && n.tag_name().name() == "c" && n.tag_name().namespace() == Some(main_ns)
    }) {
        let Some(r) = cell.attribute("r") else { continue };
        let Ok(cell_ref) = CellRef::from_a1(r) else { continue };
        let value = cell
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "v")
            .and_then(|n| n.text())
            .and_then(|t| t.parse::<f64>().ok())
            .map(CellValue::Number)
            .unwrap_or(CellValue::Empty);
        values.insert(cell_ref, value);
    }
    SheetValues { values }
}

fn fixture(path: &str) -> String {
    let dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path);
    dir.to_string_lossy().to_string()
}

fn load_package(path: &str) -> XlsxPackage {
    let bytes = std::fs::read(fixture(path)).expect("fixture exists");
    XlsxPackage::from_bytes(&bytes).expect("read xlsx")
}

#[test]
fn parses_office2007_conditional_formatting() {
    let pkg = load_package("conditional_formatting_2007.xlsx");
    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let sheet_cf = parse_worksheet_conditional_formatting(sheet_xml).unwrap();
    assert_eq!(sheet_cf.rules.len(), 4);
    assert!(sheet_cf
        .raw_blocks
        .iter()
        .any(|b| matches!(b.schema, formula_model::CfRuleSchema::Office2007)));
    let styles = Styles::parse(styles_xml).unwrap();
    assert_eq!(styles.dxfs.len(), 1);
}

#[test]
fn parses_x14_and_merges_extensions() {
    let pkg = load_package("conditional_formatting_x14.xlsx");
    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet_cf = parse_worksheet_conditional_formatting(sheet_xml).unwrap();
    assert_eq!(sheet_cf.rules.len(), 1, "x14 cfRule should merge into base rule");
    let rule = &sheet_cf.rules[0];
    assert_eq!(rule.schema, formula_model::CfRuleSchema::X14);
    match &rule.kind {
        formula_model::CfRuleKind::DataBar(db) => {
            assert_eq!(
                format!("{:08X}", db.color.unwrap().argb().unwrap_or(0)),
                "FF638EC6"
            );
            assert_eq!(db.min_length, Some(0));
            assert_eq!(db.max_length, Some(100));
            assert_eq!(db.gradient, Some(false));
            assert_eq!(
                format!("{:08X}", db.negative_fill_color.unwrap().argb().unwrap_or(0)),
                "FFFF0000"
            );
            assert_eq!(
                format!("{:08X}", db.axis_color.unwrap().argb().unwrap_or(0)),
                "FF000000"
            );
            assert_eq!(db.direction, Some(DataBarDirection::LeftToRight));
        }
        other => panic!("expected DataBar rule, got {other:?}"),
    }
    assert!(sheet_cf
        .raw_blocks
        .iter()
        .any(|b| matches!(b.schema, formula_model::CfRuleSchema::X14)));
}

#[test]
fn round_trip_preserves_conditional_formatting_xml() {
    let pkg = load_package("conditional_formatting_2007.xlsx");
    let written = pkg.write_to_bytes().unwrap();
    let reopened = XlsxPackage::from_bytes(&written).unwrap();
    assert_eq!(
        pkg.part("xl/worksheets/sheet1.xml"),
        reopened.part("xl/worksheets/sheet1.xml")
    );
    assert_eq!(
        pkg.part("xl/styles.xml"),
        reopened.part("xl/styles.xml")
    );
}

#[test]
fn dependencies_include_cfvo_formula_references() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="B1:B2">
    <cfRule type="dataBar" priority="1">
      <dataBar>
        <cfvo type="min"/>
        <cfvo type="formula" val="Sheet1!$A$1"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
</worksheet>"#;

    let parsed = parse_worksheet_conditional_formatting(xml).unwrap();
    assert_eq!(parsed.rules.len(), 1);
    let deps = &parsed.rules[0].dependencies;
    assert!(deps.contains(&parse_range_a1("B1:B2").unwrap()));
    assert!(deps.contains(&parse_range_a1("$A$1").unwrap()));
}

#[test]
fn evaluates_visible_range_and_renders_snapshot() {
    let pkg = load_package("conditional_formatting_2007.xlsx");
    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let values = parse_sheet_values(sheet_xml);
    let cf = parse_worksheet_conditional_formatting(sheet_xml).unwrap();
    let rules = cf.rules.clone();
    let mut engine = ConditionalFormattingEngine::new();
    let styles = Styles::parse(styles_xml).unwrap();
    let dxf_provider = DxfProvider { styles: &styles };

    let visible = parse_range_a1("A1:D3").unwrap();
    let eval = engine.evaluate_visible_range(&rules, visible, &values, None, Some(&dxf_provider));

    let snapshot = format_render_plan(visible, eval);
    insta::assert_snapshot!("conditional_formatting_render_plan", snapshot);
}
