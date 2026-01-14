use std::io::Cursor;

use formula_xlsx::drawingml::preserve_drawing_parts_from_reader;
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

fn build_workbook_with_chart() -> Vec<u8> {
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

    worksheet.insert_chart(1, 3, &chart).unwrap();

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

#[test]
fn preserve_drawing_parts_from_reader_includes_chart_style_and_color_parts() {
    let bytes = build_workbook_with_chart();
    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    // Inject chart style/color style parts and relationships into the chart `.rels`.
    let chart_part = pkg
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    let rels_xml = pkg
        .part(&chart_rels_part)
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
    let updated_rels = insert_before_closing_tag(&rels_xml, "</Relationships>", &insert);
    pkg.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    pkg.set_part(STYLE_PART.to_string(), STYLE_XML.to_vec());
    pkg.set_part(COLORS_PART.to_string(), COLORS_XML.to_vec());

    // Ensure content types includes the necessary overrides.
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
    let updated_ct = insert_before_closing_tag(&ct_xml, "</Types>", &insert);
    pkg.set_part(ct_part, updated_ct.into_bytes());

    let bytes = pkg.write_to_bytes().unwrap();

    // Now run the streaming preservation pipeline and ensure it includes the injected parts.
    let preserved = preserve_drawing_parts_from_reader(Cursor::new(bytes)).unwrap();

    assert_eq!(preserved.parts.get(STYLE_PART).unwrap().as_slice(), STYLE_XML);
    assert_eq!(
        preserved.parts.get(COLORS_PART).unwrap().as_slice(),
        COLORS_XML
    );

    let preserved_chart_rels_xml = std::str::from_utf8(preserved.parts.get(&chart_rels_part).unwrap())
        .expect("preserved chart rels should be utf-8");
    assert_relationship_exists(
        preserved_chart_rels_xml,
        REL_TYPE_CHART_STYLE,
        "style1.xml",
    );
    assert_relationship_exists(
        preserved_chart_rels_xml,
        REL_TYPE_CHART_COLOR_STYLE,
        "colors1.xml",
    );
}

