use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn pivot_slicer_parts_ignores_malformed_drawing_rels() {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer1" uid="{11111111-1111-1111-1111-111111111111}">
  <slicerCache id="rId1"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCache name="SlicerCache1" sourceName="PivotTable1">
    <slicerCachePivotTables>
      <slicerCachePivotTable id="rId1"/>
    </slicerCachePivotTables>
  </slicerCache>
</slicerCacheDefinition>"#;

    let slicer_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline1" uid="{22222222-2222-2222-2222-222222222222}">
  <timelineCache id="rId1"/>
</timeline>"#;

    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition" Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    let timeline_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
  name="TimelineCache1" sourceName="PivotTable1" baseField="0" level="0">
  <pivotTables>
    <pivotTable id="rId1"/>
  </pivotTables>
</timelineCacheDefinition>"#;

    let timeline_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let drawing1_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>"#;

    // Invalid attribute syntax (missing quote) to exercise best-effort `.rels` parsing.
    let broken_drawing_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml" TargetMode="External/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        (
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            slicer_cache_rels,
        ),
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache_xml,
        ),
        (
            "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
            timeline_cache_rels,
        ),
        ("xl/drawings/_rels/drawing1.xml.rels", drawing1_rels),
        ("xl/drawings/_rels/drawing2.xml.rels", broken_drawing_rels),
        // Optional, but makes the resolved target concrete.
        ("xl/pivotTables/pivotTable1.xml", br#"<pivotTableDefinition/>"#),
    ]);

    let package = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
    let parts = package
        .pivot_slicer_parts()
        .expect("pivot slicer discovery should be best-effort");

    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(parts.timelines.len(), 1);

    let slicer = &parts.slicers[0];
    assert_eq!(slicer.placed_on_drawings, vec!["xl/drawings/drawing1.xml"]);
    assert_eq!(
        slicer.connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml".to_string()]
    );

    let timeline = &parts.timelines[0];
    assert_eq!(timeline.placed_on_drawings, vec!["xl/drawings/drawing1.xml"]);
    assert_eq!(
        timeline.connected_pivot_tables,
        vec!["xl/pivotTables/pivotTable1.xml".to_string()]
    );
}

#[test]
fn pivot_slicer_parts_ignores_malformed_cache_rels() {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer1">
  <slicerCache id="rId1"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicerCache name="SlicerCache1" sourceName="PivotTable1">
    <slicerCachePivotTables>
      <slicerCachePivotTable id="rId1"/>
    </slicerCachePivotTables>
  </slicerCache>
</slicerCacheDefinition>"#;

    // Invalid attribute syntax (missing quote) in the cache `.rels` file.
    let broken_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml" TargetMode="Internal/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        (
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            broken_cache_rels,
        ),
    ]);

    let package = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
    let parts = package
        .pivot_slicer_parts()
        .expect("pivot slicer discovery should be best-effort");

    assert_eq!(parts.slicers.len(), 1);
    let slicer = &parts.slicers[0];
    assert_eq!(slicer.cache_name.as_deref(), Some("SlicerCache1"));
    assert_eq!(slicer.source_name.as_deref(), Some("PivotTable1"));
    assert!(slicer.connected_pivot_tables.is_empty());
}

