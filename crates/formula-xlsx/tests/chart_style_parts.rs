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

#[test]
fn detects_and_preserves_chart_style_and_color_style_parts() {
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
        r#"  <Relationship Id="rId998" Target="style1.xml"/>"#,
    );
    updated_rels.insert_str(
        insert_idx,
        r#"  <Relationship Id="rId997" Target="colors1.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let style_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="10"/>"#
        .to_vec();
    let colors_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="11">
  <a:srgbClr val="112233"/>
  <a:schemeClr val="accent1"/>
  <a:sysClr val="windowText" lastClr="445566"/>
  <a:prstClr val="red"/>
</cs:colorStyle>"#
        .to_vec();

    package.set_part("xl/charts/style1.xml", style_xml.clone());
    package.set_part("xl/charts/colors1.xml", colors_xml.clone());

    // Reload so that we exercise the real extraction path.
    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);

    let chart_object = &chart_objects[0];
    let style_part = chart_object
        .parts
        .style
        .as_ref()
        .expect("style part detected");
    assert_eq!(style_part.path, "xl/charts/style1.xml");
    assert_eq!(style_part.bytes, style_xml);

    let colors_part = chart_object
        .parts
        .colors
        .as_ref()
        .expect("colors part detected");
    assert_eq!(colors_part.path, "xl/charts/colors1.xml");
    assert_eq!(colors_part.bytes, colors_xml);

    let model = chart_object.model.as_ref().expect("chart model parsed");
    let style_model = model.style_part.as_ref().expect("style attached to model");
    assert_eq!(style_model.id, Some(10));
    assert_eq!(
        style_model.raw_xml,
        std::str::from_utf8(&style_part.bytes).unwrap()
    );

    let colors_model = model
        .colors_part
        .as_ref()
        .expect("colors attached to model");
    assert_eq!(colors_model.id, Some(11));
    assert_eq!(
        colors_model.colors,
        vec![
            "112233".to_string(),
            "scheme:accent1".to_string(),
            "sys:lastClr:445566".to_string(),
            "prst:red".to_string(),
        ]
    );
    assert_eq!(
        colors_model.raw_xml,
        std::str::from_utf8(&colors_part.bytes).unwrap()
    );

    // Ensure round-tripping preserves style/color parts byte-for-byte.
    let roundtrip = package.write_to_bytes().expect("round-trip write");
    let package2 = XlsxPackage::from_bytes(&roundtrip).expect("parse round-trip package");
    assert_eq!(
        package.part(&style_part.path),
        package2.part(&style_part.path),
        "chartStyle xml should round-trip losslessly"
    );
    assert_eq!(
        package.part(&colors_part.path),
        package2.part(&colors_part.path),
        "chartColorStyle xml should round-trip losslessly"
    );
}
