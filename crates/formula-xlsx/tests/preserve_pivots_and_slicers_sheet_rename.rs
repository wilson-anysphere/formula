use formula_xlsx::XlsxPackage;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");

#[test]
fn preserved_pivots_and_drawings_survive_sheet_rename_by_index() {
    let mut source = XlsxPackage::from_bytes(FIXTURE).expect("load fixture");

    // Ensure the source sheet XML actually references the drawing + pivot table so the preserve
    // routines can extract the relationship IDs.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let sheet_xml = std::str::from_utf8(source.part(sheet_part).expect("sheet1.xml present"))
        .expect("sheet xml utf-8");
    let patched_sheet_xml = sheet_xml.replace(
        "<sheetData/>",
        "<sheetData/><drawing r:id=\"rId1\"/><pivotTables><pivotTable r:id=\"rId2\"/></pivotTables>",
    );
    source.set_part(sheet_part, patched_sheet_xml.into_bytes());

    let preserved_drawings = source
        .preserve_drawing_parts()
        .expect("preserve drawings");
    assert!(!preserved_drawings.is_empty());
    let preserved_pivots = source.preserve_pivot_parts().expect("preserve pivots");
    assert!(!preserved_pivots.is_empty());

    let mut dest = XlsxPackage::from_bytes(FIXTURE).expect("load dest fixture");

    // Simulate a user renaming a sheet in-app before saving.
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = std::str::from_utf8(dest.part(workbook_part).expect("workbook.xml present"))
        .expect("workbook xml utf-8");
    dest.set_part(
        workbook_part,
        workbook_xml
            .replace("name=\"Sheet1\"", "name=\"Renamed\"")
            .into_bytes(),
    );

    // Simulate a regenerated workbook that dropped drawing/pivot attachments and parts.
    dest.parts_map_mut()
        .remove("xl/worksheets/_rels/sheet1.xml.rels");

    let removed_prefixes = [
        "xl/drawings/",
        "xl/charts/",
        "xl/media/",
        "xl/pivotTables/",
        "xl/pivotCache/",
        "xl/slicers/",
        "xl/slicerCaches/",
        "xl/timelines/",
        "xl/timelineCaches/",
    ];
    let removed = dest
        .part_names()
        .filter(|name| removed_prefixes.iter().any(|prefix| name.starts_with(prefix)))
        .map(str::to_string)
        .collect::<Vec<_>>();
    for name in removed {
        dest.parts_map_mut().remove(&name);
    }

    dest.apply_preserved_drawing_parts(&preserved_drawings)
        .expect("apply drawings");
    dest.apply_preserved_pivot_parts(&preserved_pivots)
        .expect("apply pivots");

    let workbook_xml = std::str::from_utf8(dest.part(workbook_part).expect("workbook.xml exists"))
        .expect("workbook xml utf-8");
    assert!(
        workbook_xml.contains("name=\"Renamed\""),
        "expected workbook.xml to contain renamed sheet"
    );

    let sheet_xml = std::str::from_utf8(dest.part(sheet_part).expect("sheet exists"))
        .expect("sheet xml utf-8");
    assert!(
        sheet_xml.contains("<drawing r:id=\"rId1\""),
        "expected <drawing> to be re-attached to renamed sheet"
    );
    assert!(
        sheet_xml.contains("<pivotTables"),
        "expected <pivotTables> to be re-attached to renamed sheet"
    );
    assert!(
        sheet_xml.contains("<pivotTable r:id=\"rId2\""),
        "expected <pivotTable> entry to be re-attached to renamed sheet"
    );

    let sheet_rels = std::str::from_utf8(
        dest.part("xl/worksheets/_rels/sheet1.xml.rels")
            .expect("sheet rels exist"),
    )
    .expect("sheet rels utf-8");
    assert!(sheet_rels.contains("Relationship Id=\"rId1\""));
    assert!(sheet_rels.contains(
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\""
    ));
    assert!(sheet_rels.contains("Target=\"../drawings/drawing1.xml\""));
    assert!(sheet_rels.contains("Relationship Id=\"rId2\""));
    assert!(sheet_rels.contains(
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\""
    ));
    assert!(sheet_rels.contains("Target=\"../pivotTables/pivotTable1.xml\""));

    // Slicers/timelines are attached via the sheet drawing relationships, so ensure the drawing
    // `.rels` still points at them and the parts exist.
    let drawing_rels = std::str::from_utf8(
        dest.part("xl/drawings/_rels/drawing1.xml.rels")
            .expect("drawing rels exist"),
    )
    .expect("drawing rels utf-8");
    assert!(drawing_rels.contains("Target=\"../slicers/slicer1.xml\""));
    assert!(drawing_rels.contains("Target=\"../timelines/timeline1.xml\""));

    for part in [
        "xl/pivotTables/pivotTable1.xml",
        "xl/slicers/slicer1.xml",
        "xl/slicers/_rels/slicer1.xml.rels",
        "xl/slicerCaches/slicerCache1.xml",
        "xl/slicerCaches/_rels/slicerCache1.xml.rels",
        "xl/timelines/timeline1.xml",
        "xl/timelines/_rels/timeline1.xml.rels",
        "xl/timelineCaches/timelineCacheDefinition1.xml",
        "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
    ] {
        assert!(dest.part(part).is_some(), "expected part {part} to be present");
    }
}
