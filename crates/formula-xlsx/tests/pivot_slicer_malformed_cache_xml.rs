use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;
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
fn pivot_slicer_parts_tolerates_malformed_slicer_cache_xml() {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer1">
  <slicerCache r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    // Malformed XML (missing close angle bracket).
    let slicer_cache_xml =
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><slicerCache"#;

    // Keep `.rels` valid so only the cache XML is broken.
    let slicer_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>"#;

    let pkg = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        (
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            slicer_cache_rels,
        ),
    ]);

    let parsed = pkg
        .pivot_slicer_parts()
        .expect("should tolerate malformed cache xml");

    assert_eq!(parsed.slicers.len(), 1);
    let slicer = &parsed.slicers[0];

    assert_eq!(
        slicer.cache_part.as_deref(),
        Some("xl/slicerCaches/slicerCache1.xml")
    );
    assert_eq!(slicer.cache_name, None);
    assert_eq!(slicer.source_name, None);
    assert!(slicer.connected_pivot_tables.is_empty());
    assert_eq!(slicer.connected_tables, vec!["xl/tables/table1.xml".to_string()]);
}

#[test]
fn pivot_slicer_parts_tolerates_malformed_timeline_cache_xml() {
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="Timeline1">
  <timelineCache r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</timeline>"#;

    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:timelineCache" Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    // Malformed XML (missing close tag).
    let timeline_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <pivotTables>"#;

    let pkg = build_package(&[
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache_xml,
        ),
    ]);

    let parsed = pkg
        .pivot_slicer_parts()
        .expect("should tolerate malformed timeline cache xml");

    assert_eq!(parsed.timelines.len(), 1);
    let timeline = &parsed.timelines[0];

    assert_eq!(
        timeline.cache_part.as_deref(),
        Some("xl/timelineCaches/timelineCacheDefinition1.xml")
    );
    assert_eq!(timeline.cache_name, None);
    assert_eq!(timeline.source_name, None);
    assert_eq!(timeline.base_field, None);
    assert_eq!(timeline.level, None);
    assert!(timeline.connected_pivot_tables.is_empty());
}

