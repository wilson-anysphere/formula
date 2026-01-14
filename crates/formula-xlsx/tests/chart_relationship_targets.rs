use formula_xlsx::drawingml::charts::ChartDiagnosticLevel;
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

fn build_workbook_with_two_charts() -> Vec<u8> {
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

    let mut chart1 = Chart::new(XlsxChartType::Column);
    chart1.title().set_name("Chart 1");
    chart1
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");
    worksheet.insert_chart(1, 3, &chart1).unwrap();

    let mut chart2 = Chart::new(XlsxChartType::Line);
    chart2.title().set_name("Chart 2");
    chart2
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");
    worksheet.insert_chart(16, 3, &chart2).unwrap();

    workbook.save_to_buffer().unwrap()
}

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

    chart
        .add_series()
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

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn write_relationships_xml(rels: &[formula_xlsx::openxml::Relationship]) -> Vec<u8> {
    let mut out = String::new();
    out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    out.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for rel in rels {
        out.push_str(&format!(
            r#"<Relationship Id="{}" Type="{}" Target="{}""#,
            xml_escape(&rel.id),
            xml_escape(&rel.type_uri),
            xml_escape(&rel.target)
        ));
        if let Some(mode) = &rel.target_mode {
            out.push_str(&format!(r#" TargetMode="{}""#, xml_escape(mode)));
        }
        out.push_str("/>");
    }
    out.push_str("</Relationships>");
    out.into_bytes()
}

#[test]
fn chart_object_extraction_strips_fragments_and_handles_external_chart_relationships() {
    let bytes = build_workbook_with_two_charts();
    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let drawing_part = pkg
        .part_names()
        .find(|p| p.starts_with("xl/drawings/drawing") && p.ends_with(".xml"))
        .expect("drawing part present")
        .to_string();
    let drawing_rels_part = rels_for_part(&drawing_part);
    let drawing_rels_xml = pkg
        .part(&drawing_rels_part)
        .expect("drawing rels present");

    let mut rels = formula_xlsx::openxml::parse_relationships(drawing_rels_xml)
        .expect("parse drawing relationships");

    let mut chart_rel_targets: Vec<_> = rels
        .iter()
        .filter(|rel| rel.type_uri.ends_with("/relationships/chart"))
        .map(|rel| rel.target.clone())
        .collect();
    chart_rel_targets.sort();

    assert!(
        chart_rel_targets.iter().any(|t| t.ends_with("chart1.xml")),
        "expected chart1 relationship target in drawing rels, got: {chart_rel_targets:?}"
    );
    assert!(
        chart_rel_targets.iter().any(|t| t.ends_with("chart2.xml")),
        "expected chart2 relationship target in drawing rels, got: {chart_rel_targets:?}"
    );

    let mut patched_fragment = false;
    let mut patched_external = false;
    for rel in &mut rels {
        if !rel.type_uri.ends_with("/relationships/chart") {
            continue;
        }
        if !patched_fragment && rel.target.ends_with("chart1.xml") {
            rel.target = format!("{}#foo", rel.target);
            patched_fragment = true;
            continue;
        }
        if !patched_external && rel.target.ends_with("chart2.xml") {
            rel.target_mode = Some("External".to_string());
            rel.target = "https://example.com/chart2.xml".to_string();
            patched_external = true;
        }
    }

    assert!(patched_fragment, "expected to patch chart1 relationship target");
    assert!(
        patched_external,
        "expected to patch chart2 relationship as TargetMode=External"
    );

    pkg.set_part(drawing_rels_part, write_relationships_xml(&rels));

    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = pkg.extract_chart_objects().unwrap();
    assert_eq!(
        chart_objects.len(),
        2,
        "expected external chart relationships to be preserved as empty chart objects"
    );

    let chart_object = chart_objects
        .iter()
        .find(|chart| chart.parts.chart.path == "xl/charts/chart1.xml")
        .expect("expected chart1 object to be extracted");
    assert!(
        chart_object
            .diagnostics
            .iter()
            .all(|d| d.level != ChartDiagnosticLevel::Error),
        "expected no error diagnostics, got: {:#?}",
        chart_object.diagnostics
    );

    let external_chart_object = chart_objects
        .iter()
        .find(|chart| chart.parts.chart.path.is_empty())
        .expect("expected external chart object to have empty part path");
    assert!(
        external_chart_object.parts.chart.bytes.is_empty(),
        "expected external chart object to have empty bytes"
    );
    assert!(
        external_chart_object.diagnostics.iter().any(|d| {
            d.level == ChartDiagnosticLevel::Warning
                && d.message.to_ascii_lowercase().contains("external")
        }),
        "expected warning diagnostic for external chart relationship, got: {:#?}",
        external_chart_object.diagnostics
    );

    let charts = pkg.extract_charts().unwrap();
    assert!(
        charts
            .iter()
            .any(|chart| chart.chart_part.as_deref() == Some("xl/charts/chart1.xml")),
        "expected extract_charts() to resolve the fragment chart target"
    );
}

#[test]
fn chart_object_extraction_normalizes_backslashes_in_chart_relationship_targets() {
    let bytes = build_workbook_with_chart();
    let mut package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_part = package
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    let mut updated_rels = package
        .part(&chart_rels_part)
        .map(|bytes| String::from_utf8(bytes.to_vec()).unwrap())
        .unwrap_or_else(|| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
        });
    let insert_idx = updated_rels
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    updated_rels.insert_str(
        insert_idx,
        r#"  <Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target=".\chartEx1.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let chart_ex_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:spPr/></cx:chartSpace>"#.to_vec();
    package.set_part("xl/charts/chartEx1.xml", chart_ex_xml.clone());
    package.set_part(
        "xl/charts/_rels/chartEx1.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
            .to_vec(),
    );

    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();
    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);

    let chart_object = &chart_objects[0];
    let chart_ex = chart_object
        .parts
        .chart_ex
        .as_ref()
        .expect("chartEx part detected via backslash target");
    assert_eq!(chart_ex.path, "xl/charts/chartEx1.xml");
    assert_eq!(chart_ex.bytes, chart_ex_xml);
}
