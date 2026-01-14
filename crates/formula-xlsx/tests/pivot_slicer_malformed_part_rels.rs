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
fn pivot_slicer_parts_tolerates_malformed_slicer_part_rels() {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <slicerCache r:id="rId1"/>
</slicer>"#;

    // Intentionally malformed XML: truncated `<Relationship` element.
    let slicer_rels_malformed = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:slicerCache" Target="../slicerCaches/slicerCache1.xml"
</Relationships>"#;

    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>"#;

    let bytes = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels_malformed),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
    ]);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let parts = pkg.pivot_slicer_parts().expect("pivot_slicer_parts should be best-effort");
    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(parts.timelines.len(), 0);
    assert_eq!(parts.slicers[0].part_name, "xl/slicers/slicer1.xml");
    assert_eq!(parts.slicers[0].cache_part, None);
}

#[test]
fn pivot_slicer_parts_tolerates_malformed_timeline_part_rels() {
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    // Intentionally malformed XML: truncated `<Relationship` element.
    let timeline_rels_malformed = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:timelineCache" Target="../timelineCaches/timelineCache1.xml"
</Relationships>"#;

    // Unused in this test (the malformed `.rels` prevents resolution), but included to mirror a
    // realistic package layout.
    let timeline_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>"#;

    let bytes = build_package(&[
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels_malformed),
        ("xl/timelineCaches/timelineCache1.xml", timeline_cache_xml),
    ]);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let parts = pkg.pivot_slicer_parts().expect("pivot_slicer_parts should be best-effort");
    assert_eq!(parts.slicers.len(), 0);
    assert_eq!(parts.timelines.len(), 1);
    assert_eq!(parts.timelines[0].part_name, "xl/timelines/timeline1.xml");
    assert_eq!(parts.timelines[0].cache_part, None);
}

