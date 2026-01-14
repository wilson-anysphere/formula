use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
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

#[test]
fn parses_external_data_in_real_xlsx_chart_part() {
    let bytes = build_workbook_with_chart();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("open xlsx");

    let chart_part = pkg
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    // Patch chartSpace XML to include `<c:externalData>`.
    let mut chart_xml =
        String::from_utf8(pkg.part(&chart_part).expect("chart xml present").to_vec())
            .expect("chart xml is utf-8");
    if !chart_xml.contains("xmlns:r=") {
        chart_xml = chart_xml.replacen(
            "<c:chartSpace",
            r#"<c:chartSpace xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            1,
        );
    }
    let external_data_xml =
        r#"<c:externalData r:id="rIdExternal"><c:autoUpdate val="1"/></c:externalData>"#;
    chart_xml = chart_xml.replacen(
        "</c:chartSpace>",
        &format!("{external_data_xml}</c:chartSpace>"),
        1,
    );
    pkg.set_part(chart_part.clone(), chart_xml.into_bytes());

    // Patch chart relationships to include the externalLink relationship referenced by r:id.
    let mut rels_xml = pkg
        .part(&chart_rels_part)
        .map(|bytes| String::from_utf8(bytes.to_vec()).unwrap())
        .unwrap_or_else(|| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
                .to_string()
        });
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"<Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="../externalLinks/externalLink1.xml"/>"#,
    );
    pkg.set_part(chart_rels_part, rels_xml.into_bytes());

    // Add a minimal externalLink part that points at an external workbook path.
    pkg.set_part(
        "xl/externalLinks/externalLink1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1"/>
</externalLink>"#
            .to_vec(),
    );
    pkg.set_part(
        "xl/externalLinks/_rels/externalLink1.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="external.xlsx" TargetMode="External"/>
</Relationships>"#
            .to_vec(),
    );

    let charts = pkg.extract_chart_objects().expect("extract chart objects");
    assert_eq!(charts.len(), 1);

    let model = charts[0].model.as_ref().expect("chart model present");
    assert_eq!(model.external_data_rel_id.as_deref(), Some("rIdExternal"));
    assert_eq!(model.external_data_auto_update, Some(true));
}
