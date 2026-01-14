use formula_xlsx::XlsxPackage;
use std::io::{Cursor, Write};

use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    let bytes = zip.finish().unwrap().into_inner();
    XlsxPackage::from_bytes(&bytes).expect("read test pkg")
}

#[test]
fn pivot_slicer_parts_tolerates_malformed_slicer_cache_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        name="TestSlicer">
  <slicerCache r:id="rId1"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             name="Cache1" sourceName="Field1">
  <slicerCachePivotTable r:id="rIdBroken"/>
</slicerCache>"#;

    // Intentionally malformed XML: mismatched tags (`Relationship` is never closed).
    let malformed_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBroken" Type="urn:example:pivotTable" Target="../pivotTables/pivotTable1.xml">
</Relationships>"#;

    let package = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        ("xl/slicerCaches/_rels/slicerCache1.xml.rels", malformed_cache_rels),
    ]);

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.slicers.len(), 1);

    let slicer = &parts.slicers[0];
    assert_eq!(slicer.part_name, "xl/slicers/slicer1.xml");
    assert!(slicer.connected_pivot_tables.is_empty());

    Ok(())
}

#[test]
fn pivot_slicer_parts_tolerates_malformed_timeline_cache_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="TestTimeline">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:timelineCache" Target="../timelineCaches/timelineCache1.xml"/>
</Relationships>"#;

    let timeline_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                         name="TimelineCache1" sourceName="Field1">
  <pivotTable r:id="rIdBroken"/>
</timelineCacheDefinition>"#;

    // Intentionally malformed XML: mismatched tags (`Relationship` is never closed).
    let malformed_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBroken" Type="urn:example:pivotTable" Target="../pivotTables/pivotTable1.xml">
</Relationships>"#;

    let package = build_package(&[
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels),
        ("xl/timelineCaches/timelineCache1.xml", timeline_cache_xml),
        ("xl/timelineCaches/_rels/timelineCache1.xml.rels", malformed_cache_rels),
    ]);

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.timelines.len(), 1);

    let timeline = &parts.timelines[0];
    assert_eq!(timeline.part_name, "xl/timelines/timeline1.xml");
    assert!(timeline.connected_pivot_tables.is_empty());

    Ok(())
}
