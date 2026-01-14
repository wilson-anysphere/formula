use formula_xlsx::XlsxPackage;
use formula_xlsx::drawingml::charts::ChartDiagnosticLevel;
use rust_xlsxwriter::{Chart, ChartType, Workbook};

const FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/charts/chart-user-shapes.xlsx");

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
fn ignores_external_chart_user_shapes_relationships() {
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
        r#"  <Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" Target="https://example.com" TargetMode="External"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);
    let chart_object = &chart_objects[0];
    assert!(
        chart_object.parts.user_shapes.is_none(),
        "external chartUserShapes relationship should be ignored"
    );
    assert!(
        chart_object
            .diagnostics
            .iter()
            .all(|d| !d.message.to_ascii_lowercase().contains("usershapes")),
        "expected no userShapes-related diagnostics for external relationship"
    );
}

#[test]
fn warns_when_chart_user_shapes_relationship_targets_missing_part() {
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

    let bytes = package.write_to_bytes().unwrap();
    let package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_objects = package.extract_chart_objects().unwrap();
    assert_eq!(chart_objects.len(), 1);
    let chart_object = &chart_objects[0];
    assert!(
        chart_object.parts.user_shapes.is_none(),
        "missing chartUserShapes part should not crash extraction"
    );
    assert!(
        chart_object.diagnostics.iter().any(|d| {
            d.level == ChartDiagnosticLevel::Warning
                && d.message.contains("missing chart userShapes part")
        }),
        "expected warning about missing chart userShapes part"
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

#[test]
fn detects_chart_user_shapes_part_via_drawing_target_heuristic_when_type_missing() {
    let bytes = build_workbook_with_chart();
    let mut package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_part = package
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    // Some producers omit the relationship type; ensure we still pick up the
    // chart userShapes drawing by filename heuristic (`drawing*.xml`).
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
        r#"  <Relationship Id="rId999" Type="" Target="../drawings/drawing99.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let user_shapes_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></cdr:wsDr>"#.to_vec();
    package.set_part("xl/drawings/drawing99.xml", user_shapes_xml.clone());

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
}

#[test]
fn detects_chart_user_shapes_part_when_target_has_query_string_and_type_missing() {
    let bytes = build_workbook_with_chart();
    let mut package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_part = package
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    // Some producers include URI query strings in relationship targets. These are not part of the
    // actual OPC part name, so we should ignore them when applying filename heuristics.
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
        r#"  <Relationship Id="rId999" Type="" Target="../drawings/drawing99.xml?foo=bar"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    let user_shapes_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></cdr:wsDr>"#.to_vec();
    package.set_part("xl/drawings/drawing99.xml", user_shapes_xml.clone());

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
}

#[test]
fn does_not_misclassify_other_xml_in_drawings_dir_as_chart_user_shapes_when_type_missing() {
    let bytes = build_workbook_with_chart();
    let mut package = XlsxPackage::from_bytes(&bytes).unwrap();

    let chart_part = package
        .part_names()
        .find(|p| p.starts_with("xl/charts/chart") && p.ends_with(".xml"))
        .expect("chart part present")
        .to_string();
    let chart_rels_part = rels_for_part(&chart_part);

    // Ensure that *any* `../drawings/*.xml` target doesn't trigger the heuristic: the old heuristic
    // matched on the substring `"drawing"`, which appears in the directory name `drawings/` and
    // could therefore misclassify unrelated drawing XML parts.
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
        r#"  <Relationship Id="rId999" Type="" Target="../drawings/drawing99.xml"/>"#,
    );
    updated_rels.insert_str(
        insert_idx,
        r#"  <Relationship Id="rId998" Type="" Target="../drawings/other99.xml"/>"#,
    );
    package.set_part(chart_rels_part.clone(), updated_rels.into_bytes());

    package.set_part("xl/drawings/other99.xml", br#"<root/>"#.to_vec());

    let user_shapes_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></cdr:wsDr>"#.to_vec();
    package.set_part("xl/drawings/drawing99.xml", user_shapes_xml.clone());

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
}

#[test]
fn extracts_chart_user_shapes_part_from_fixture_and_preserves_bytes() {
    let pkg = XlsxPackage::from_bytes(FIXTURE).expect("open xlsx fixture");
    let charts = pkg.extract_chart_objects().expect("extract chart objects");
    assert!(
        !charts.is_empty(),
        "expected fixture to contain at least one chart object"
    );

    let chart = &charts[0];
    let user_shapes = chart
        .parts
        .user_shapes
        .as_ref()
        .expect("expected chartUserShapes part to be detected");
    assert_eq!(user_shapes.path, "xl/drawings/drawing2.xml");
    assert!(
        !user_shapes.bytes.is_empty(),
        "expected chartUserShapes part to have bytes"
    );
    assert_eq!(
        user_shapes.rels_path.as_deref(),
        Some("xl/drawings/_rels/drawing2.xml.rels"),
        "expected chartUserShapes rels path to be captured"
    );
    let user_shapes_rels_path = user_shapes.rels_path.as_deref().unwrap();
    assert_eq!(
        user_shapes.rels_bytes.as_deref(),
        pkg.part(user_shapes_rels_path),
        "expected chartUserShapes rels bytes to be embedded on the extracted OpcPart"
    );

    // No-op round-trip should preserve the part bytes byte-for-byte.
    let roundtrip = pkg.write_to_bytes().expect("round-trip write");
    let pkg2 = XlsxPackage::from_bytes(&roundtrip).expect("parse round-trip package");
    assert_eq!(
        pkg.part(&user_shapes.path),
        pkg2.part(&user_shapes.path),
        "chartUserShapes xml should round-trip losslessly"
    );
    assert_eq!(
        pkg.part(user_shapes_rels_path),
        pkg2.part(user_shapes_rels_path),
        "chartUserShapes rels should round-trip losslessly"
    );

    // The drawing-parts preservation pipeline should also keep chartUserShapes reachable from the
    // chart's relationship graph.
    let preserved = pkg.preserve_drawing_parts().expect("preserve drawing parts");
    assert!(
        preserved.parts.contains_key(&user_shapes.path),
        "expected preserved drawing parts to include chartUserShapes part"
    );
    assert!(
        preserved.parts.contains_key(user_shapes_rels_path),
        "expected preserved drawing parts to include chartUserShapes rels part"
    );

    // Apply preserved parts to a regenerated workbook and verify the userShapes part survives.
    let mut workbook = Workbook::new();
    workbook.add_worksheet();
    let regenerated_bytes = workbook.save_to_buffer().expect("write regenerated workbook");
    let mut regenerated_pkg =
        XlsxPackage::from_bytes(&regenerated_bytes).expect("parse regenerated workbook");
    regenerated_pkg
        .apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved drawing parts");

    let merged_bytes = regenerated_pkg
        .write_to_bytes()
        .expect("write merged workbook");
    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).expect("parse merged workbook");
    let merged_charts = merged_pkg
        .extract_chart_objects()
        .expect("extract charts from merged workbook");
    assert!(
        !merged_charts.is_empty(),
        "expected merged workbook to contain charts"
    );
    let merged_user_shapes = merged_charts[0]
        .parts
        .user_shapes
        .as_ref()
        .expect("expected merged workbook to contain chartUserShapes part");
    assert_eq!(
        merged_pkg.part(&merged_user_shapes.path),
        pkg.part(&user_shapes.path),
        "expected merged workbook to preserve chartUserShapes bytes"
    );
    assert_eq!(
        merged_pkg.part(merged_user_shapes.rels_path.as_deref().unwrap()),
        pkg.part(user_shapes_rels_path),
        "expected merged workbook to preserve chartUserShapes rels bytes"
    );
}
