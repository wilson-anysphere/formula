use formula_xlsx::XlsxPackage;
use formula_xlsx::WorkbookSheetInfo;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");

fn add_sheet2(pkg: &mut XlsxPackage) {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = std::str::from_utf8(pkg.part(workbook_part).expect("workbook.xml present"))
        .expect("workbook xml utf-8");
    if workbook_xml.contains("name=\"Sheet2\"") {
        return;
    }
    pkg.set_part(
        workbook_part,
        workbook_xml
            .replace(
                "</sheets>",
                "    <sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n  </sheets>",
            )
            .into_bytes(),
    );

    let rels_part = "xl/_rels/workbook.xml.rels";
    let rels_xml = std::str::from_utf8(pkg.part(rels_part).expect("workbook rels present"))
        .expect("rels xml utf-8");
    if !rels_xml.contains("Id=\"rId2\"") {
        pkg.set_part(
            rels_part,
            rels_xml
                .replace(
                    "</Relationships>",
                    "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n</Relationships>",
                )
                .into_bytes(),
        );
    }

    let content_types_part = "[Content_Types].xml";
    let content_types = std::str::from_utf8(pkg.part(content_types_part).expect("ct present"))
        .expect("ct utf-8");
    if !content_types.contains("/xl/worksheets/sheet2.xml") {
        pkg.set_part(
            content_types_part,
            content_types
                .replace(
                    "</Types>",
                    "  <Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n</Types>",
                )
                .into_bytes(),
        );
    }

    pkg.set_part(
        "xl/worksheets/sheet2.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#
        .to_vec(),
    );
}

fn resolve_sheet_part(pkg: &XlsxPackage, sheet: &WorkbookSheetInfo) -> String {
    let rels_part = "xl/_rels/workbook.xml.rels";
    let rels_xml = std::str::from_utf8(pkg.part(rels_part).expect("workbook rels exist"))
        .expect("rels xml utf-8");
    let doc = roxmltree::Document::parse(rels_xml).expect("parse workbook rels");
    let target = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Id") == Some(sheet.rel_id.as_str())
        })
        .and_then(|n| n.attribute("Target"))
        .expect("sheet relationship target");

    if let Some(target) = target.strip_prefix('/') {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
}

#[test]
fn preserved_pivots_and_drawings_survive_sheet_rename_by_index() {
    let mut source = XlsxPackage::from_bytes(FIXTURE).expect("load fixture");
    add_sheet2(&mut source);

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
    add_sheet2(&mut dest);

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

    let sheets = dest.workbook_sheets().expect("parse sheets");
    let renamed = sheets
        .iter()
        .find(|s| s.name == "Renamed")
        .expect("renamed sheet exists");
    let renamed_part = resolve_sheet_part(&dest, renamed);

    let sheet_xml = std::str::from_utf8(dest.part(&renamed_part).expect("renamed sheet exists"))
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

    let sheet2 = sheets.iter().find(|s| s.name == "Sheet2").expect("Sheet2 exists");
    let sheet2_part = resolve_sheet_part(&dest, sheet2);
    let sheet2_xml = std::str::from_utf8(dest.part(&sheet2_part).expect("Sheet2 xml exists"))
        .expect("sheet2 xml utf-8");
    assert!(
        !sheet2_xml.contains("<drawing"),
        "did not expect preserved drawings to attach to Sheet2"
    );
    assert!(
        !sheet2_xml.contains("<pivotTables"),
        "did not expect preserved pivot tables to attach to Sheet2"
    );
}
