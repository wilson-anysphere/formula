use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType, Workbook};

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

    let mut chart = Chart::new(ChartType::Column);
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
fn detects_chart_user_shapes_part_from_chart_relationships_and_roundtrips() {
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
        r#"  <Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="../drawings/drawing99.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let user_shapes_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></cdr:wsDr>"#.to_vec();
    package.set_part("xl/drawings/drawing99.xml", user_shapes_xml.clone());

    package.set_part(
        "xl/drawings/_rels/drawing99.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
            .to_vec(),
    );

    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);

    let chart_object = &chart_objects[0];
    let user_shapes = chart_object
        .parts
        .user_shapes
        .as_ref()
        .expect("chart userShapes part detected");
    assert_eq!(user_shapes.path, "xl/drawings/drawing99.xml");
    assert_eq!(user_shapes.bytes, user_shapes_xml);
    assert_eq!(
        user_shapes.rels_path.as_deref(),
        Some("xl/drawings/_rels/drawing99.xml.rels")
    );
    let user_shapes_rels_path = user_shapes
        .rels_path
        .as_deref()
        .expect("userShapes rels path should be present");
    assert_eq!(
        user_shapes.rels_bytes.as_deref(),
        package.part(user_shapes_rels_path),
        "userShapes rels bytes should be embedded on the extracted OpcPart"
    );

    // Ensure round-tripping preserves the userShapes part byte-for-byte.
    let roundtrip = package.write_to_bytes().expect("round-trip write");
    let pkg2 = XlsxPackage::from_bytes(&roundtrip).expect("parse round-trip package");
    assert_eq!(
        package.part(&user_shapes.path),
        pkg2.part(&user_shapes.path),
        "chart userShapes xml should round-trip losslessly"
    );
    assert_eq!(
        package.part(user_shapes.rels_path.as_deref().unwrap()),
        pkg2.part(user_shapes.rels_path.as_deref().unwrap()),
        "chart userShapes rels should round-trip losslessly"
    );
}

#[test]
fn detects_chart_user_shapes_part_via_target_heuristic_when_type_missing() {
    let bytes = build_workbook_with_chart();
    let mut package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_part = package
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    // Some producers omit the relationship type; ensure we still pick up the
    // chart userShapes drawing by filename heuristic.
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
        r#"  <Relationship Id="rId999" Type="" Target="../drawings/userShapes99.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let user_shapes_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></cdr:wsDr>"#.to_vec();
    package.set_part("xl/drawings/userShapes99.xml", user_shapes_xml.clone());

    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);

    let chart_object = &chart_objects[0];
    let user_shapes = chart_object
        .parts
        .user_shapes
        .as_ref()
        .expect("chart userShapes part detected");
    assert_eq!(user_shapes.path, "xl/drawings/userShapes99.xml");
    assert_eq!(user_shapes.bytes, user_shapes_xml);
}
