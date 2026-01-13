use formula_model::drawings::Anchor;
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

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

fn extract_chart_rid(drawing_frame_xml: &str) -> Option<String> {
    // Prefer `r:id="..."`, but fall back to `id="..."` to match the production parser.
    drawing_frame_xml
        .split("r:id=\"")
        .nth(1)
        .and_then(|s| s.split('"').next())
        .map(str::to_string)
        .or_else(|| {
            drawing_frame_xml
                .split("id=\"")
                .nth(1)
                .and_then(|s| s.split('"').next())
                .map(str::to_string)
        })
}

#[test]
fn extracts_chart_objects_with_anchor_and_parts() {
    let bytes = build_workbook_with_chart();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);

    let chart_object = &chart_objects[0];
    assert!(chart_object.parts.chart.path.starts_with("xl/charts/chart"));

    assert!(chart_object.drawing_frame_xml.contains("<c:chart"));
    let rid = extract_chart_rid(&chart_object.drawing_frame_xml).expect("chart r:id present");
    assert!(chart_object.drawing_frame_xml.contains(&rid));
    assert_eq!(chart_object.drawing_rel_id, rid);
    assert!(
        chart_object
            .drawing_object_name
            .as_deref()
            .is_some_and(|name| !name.trim().is_empty()),
        "expected drawing object name to be populated from xdr:cNvPr"
    );
    assert!(
        chart_object.drawing_object_id.is_some(),
        "expected drawing object id to be populated from xdr:cNvPr"
    );

    match chart_object.anchor {
        Anchor::TwoCell { from, to } => {
            assert!(to.cell.col > from.cell.col);
            assert!(to.cell.row > from.cell.row);
        }
        other => panic!("expected two-cell chart anchor, got {other:?}"),
    }
}

#[test]
fn detects_and_preserves_chart_ex_part_from_chart_relationships() {
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
        r#"  <Relationship Id="rId999" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="chartEx1.xml"/>"#,
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
    let chart_part = &chart_object.parts.chart;
    let chart_rels_path = chart_part
        .rels_path
        .as_deref()
        .expect("chart rels should be present");
    assert_eq!(
        chart_part.rels_bytes.as_deref(),
        package.part(chart_rels_path),
        "chart rels bytes should be stored on the OpcPart"
    );
    let chart_ex = chart_object
        .parts
        .chart_ex
        .as_ref()
        .expect("chartEx part detected");
    assert_eq!(chart_ex.path, "xl/charts/chartEx1.xml");
    assert_eq!(chart_ex.bytes, chart_ex_xml);
    assert_eq!(
        chart_ex.rels_path.as_deref(),
        Some("xl/charts/_rels/chartEx1.xml.rels")
    );
    let chart_ex_rels_path = chart_ex
        .rels_path
        .as_deref()
        .expect("chartEx rels path should be present");
    assert_eq!(
        chart_ex.rels_bytes.as_deref(),
        package.part(chart_ex_rels_path),
        "chartEx rels bytes should be stored on the OpcPart"
    );
}
