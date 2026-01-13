use formula_xlsx::XlsxPackage;
use roxmltree::Document;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

const STYLE_PART: &str = "xl/charts/style1.xml";
const COLORS_PART: &str = "xl/charts/colors1.xml";

const STYLE_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="1"/>"#;

const COLORS_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="1"/>"#;

const REL_TYPE_CHART_STYLE: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const REL_TYPE_CHART_COLOR_STYLE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

const CT_CHART_STYLE: &str = "application/vnd.ms-office.chartstyle+xml";
const CT_CHART_COLOR_STYLE: &str = "application/vnd.ms-office.chartcolorstyle+xml";

fn build_workbook_with_optional_chart(include_chart: bool) -> Vec<u8> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Category").unwrap();
    worksheet.write_string(0, 1, "Value").unwrap();

    let categories = ["A", "B", "C", "D"];
    let values = [2.0, 4.0, 3.0, 5.0];

    for (i, (cat, val)) in categories.iter().zip(values).enumerate() {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, *cat).unwrap();
        worksheet.write_number(row, 1, val).unwrap();
    }

    let mut chart = Chart::new(XlsxChartType::Column);
    chart.title().set_name("Example Chart");

    let series = chart.add_series();
    series
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    if include_chart {
        worksheet.insert_chart(1, 3, &chart).unwrap();
    }

    workbook.save_to_buffer().unwrap()
}

fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

fn insert_before_closing_tag(xml: &str, closing_tag: &str, insert: &str) -> String {
    let idx = xml
        .rfind(closing_tag)
        .unwrap_or_else(|| panic!("missing closing tag {closing_tag} in:\n{xml}"));
    let mut updated = xml.to_string();
    updated.insert_str(idx, insert);
    updated
}

fn add_chart_style_and_colors_relationships(pkg: &mut XlsxPackage, chart_part: &str) -> String {
    let rels_part = rels_for_part(chart_part);

    let rels_xml = pkg
        .part(&rels_part)
        .map(|bytes| String::from_utf8(bytes.to_vec()).expect("chart rels should be utf-8"))
        .unwrap_or_else(|| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
                .to_string()
        });

    let insert = format!(
        r#"  <Relationship Id="rId9000" Type="{REL_TYPE_CHART_STYLE}" Target="style1.xml"/>
  <Relationship Id="rId9001" Type="{REL_TYPE_CHART_COLOR_STYLE}" Target="colors1.xml"/>
"#
    );
    let updated = insert_before_closing_tag(&rels_xml, "</Relationships>", &insert);
    pkg.set_part(rels_part.clone(), updated.into_bytes());

    rels_part
}

fn add_chart_style_and_colors_overrides(pkg: &mut XlsxPackage) {
    let ct_part = "[Content_Types].xml";
    let ct_xml = pkg
        .part(ct_part)
        .expect("[Content_Types].xml exists")
        .to_vec();
    let ct_xml = String::from_utf8(ct_xml).expect("content types should be utf-8");

    let insert = format!(
        r#"  <Override PartName="/{STYLE_PART}" ContentType="{CT_CHART_STYLE}"/>
  <Override PartName="/{COLORS_PART}" ContentType="{CT_CHART_COLOR_STYLE}"/>
"#
    );
    let updated = insert_before_closing_tag(&ct_xml, "</Types>", &insert);
    pkg.set_part(ct_part, updated.into_bytes());
}

fn assert_relationship_exists(xml: &str, rel_type: &str, target: &str) {
    let doc = Document::parse(xml).expect("parse .rels xml");
    let found = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .any(|n| n.attribute("Type") == Some(rel_type) && n.attribute("Target") == Some(target));
    assert!(
        found,
        "expected relationship Type={rel_type} Target={target} in:\n{xml}"
    );
}

fn assert_content_type_override_exists(xml: &str, part_name: &str, content_type: &str) {
    let doc = Document::parse(xml).expect("parse [Content_Types].xml");
    let found = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .any(|n| {
            n.attribute("PartName") == Some(part_name)
                && n.attribute("ContentType") == Some(content_type)
        });
    assert!(
        found,
        "expected Override PartName={part_name} ContentType={content_type} in:\n{xml}"
    );
}

#[test]
fn apply_preserved_drawing_parts_preserves_chart_style_and_color_parts_and_rels() {
    // 1) Generate a baseline workbook that contains a chart.
    let bytes_with_chart = build_workbook_with_optional_chart(true);
    let mut source_pkg = XlsxPackage::from_bytes(&bytes_with_chart).unwrap();

    // 2) Patch the chart relationships to include chart style/colors parts, and inject those parts.
    let chart_part = source_pkg
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = add_chart_style_and_colors_relationships(&mut source_pkg, &chart_part);

    source_pkg.set_part(STYLE_PART.to_string(), STYLE_XML.to_vec());
    source_pkg.set_part(COLORS_PART.to_string(), COLORS_XML.to_vec());
    add_chart_style_and_colors_overrides(&mut source_pkg);

    // Round-trip through zip writer to match the preservation pipeline's typical input.
    let source_bytes = source_pkg.write_to_bytes().unwrap();
    let source_pkg = XlsxPackage::from_bytes(&source_bytes).unwrap();

    // Sanity check: the injected parts exist in the source package.
    assert_eq!(source_pkg.part(STYLE_PART).unwrap(), STYLE_XML);
    assert_eq!(source_pkg.part(COLORS_PART).unwrap(), COLORS_XML);

    // 3) Preserve drawing parts from the patched workbook.
    let preserved = source_pkg.preserve_drawing_parts().unwrap();
    assert_eq!(preserved.parts.get(STYLE_PART).unwrap(), STYLE_XML);
    assert_eq!(preserved.parts.get(COLORS_PART).unwrap(), COLORS_XML);

    // 4) Apply to a regenerated workbook that doesn't contain the chart/style/color parts.
    let regenerated_bytes = build_workbook_with_optional_chart(false);
    let mut regenerated_pkg = XlsxPackage::from_bytes(&regenerated_bytes).unwrap();
    assert!(
        regenerated_pkg.part(STYLE_PART).is_none() && regenerated_pkg.part(COLORS_PART).is_none(),
        "regenerated workbook should not already contain chart style/color parts"
    );

    regenerated_pkg
        .apply_preserved_drawing_parts(&preserved)
        .unwrap();

    let merged_bytes = regenerated_pkg.write_to_bytes().unwrap();
    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).unwrap();

    // 5) Assertions: parts exist, chart .rels references them, and bytes are preserved exactly.
    assert_eq!(merged_pkg.part(STYLE_PART).unwrap(), STYLE_XML);
    assert_eq!(merged_pkg.part(COLORS_PART).unwrap(), COLORS_XML);

    let merged_chart_part = merged_pkg
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("merged workbook should contain a chart part")
        .to_string();
    let merged_chart_rels_part = rels_for_part(&merged_chart_part);
    assert_eq!(
        merged_chart_rels_part, chart_rels_part,
        "expected chart rels part name to be stable across preserve/apply"
    );

    let chart_rels_xml =
        std::str::from_utf8(merged_pkg.part(&merged_chart_rels_part).unwrap()).unwrap();
    assert_relationship_exists(chart_rels_xml, REL_TYPE_CHART_STYLE, "style1.xml");
    assert_relationship_exists(
        chart_rels_xml,
        REL_TYPE_CHART_COLOR_STYLE,
        "colors1.xml",
    );

    let content_types_xml = std::str::from_utf8(merged_pkg.part("[Content_Types].xml").unwrap())
        .expect("[Content_Types].xml should be utf-8");
    assert_content_type_override_exists(
        content_types_xml,
        "/xl/charts/style1.xml",
        CT_CHART_STYLE,
    );
    assert_content_type_override_exists(
        content_types_xml,
        "/xl/charts/colors1.xml",
        CT_CHART_COLOR_STYLE,
    );
}
