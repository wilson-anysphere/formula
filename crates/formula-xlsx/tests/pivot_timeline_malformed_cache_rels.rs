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
fn pivot_slicer_parts_tolerates_malformed_timeline_cache_rels() {
    // Timeline referencing `rId1` for its cache definition.
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    // Valid relationship from the timeline part to the cache definition.
    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:timelineCache" Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    // Cache definition that references a pivot table via `r:id="rIdBroken"`.
    let timeline_cache_definition = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <pivotTables>
    <pivotTable r:id="rIdBroken"/>
  </pivotTables>
</timelineCacheDefinition>"#;

    // Malformed relationships part for the cache definition. This should not prevent parsing the
    // timeline; instead, it should result in zero connected pivot tables.
    let malformed_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBroken" Type="urn:example:pivotTable" Target="../pivotTables/pivotTable1.xml">
</Relationships>"#;

    let pkg = build_package(&[
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache_definition,
        ),
        (
            "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
            malformed_cache_rels,
        ),
    ]);

    let parts = pkg
        .pivot_slicer_parts()
        .expect("should parse pivot slicer parts despite malformed timeline cache .rels");

    assert_eq!(parts.timelines.len(), 1);
    assert!(parts.timelines[0].connected_pivot_tables.is_empty());
}

