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
fn pivot_slicer_parts_discovers_noncanonical_slicer_and_timeline_part_names() {
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <slicerCache r:id="rId1"/>
</slicer>"#;

    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    let bytes = build_package(&[
        // Non-canonical ZIP entry names:
        // - Windows-style separators (`\`)
        // - Different casing
        (r#"XL\slicers\slicer1.xml"#, slicer_xml),
        ("/XL/TIMELINES/TIMELINE1.XML", timeline_xml),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let parts = pkg
        .pivot_slicer_parts()
        .expect("pivot_slicer_parts should tolerate noncanonical part names");

    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(parts.timelines.len(), 1);
    assert_eq!(parts.slicers[0].part_name, "xl/slicers/slicer1.xml");
    assert_eq!(parts.slicers[0].cache_part, None);
    assert_eq!(parts.timelines[0].part_name, "xl/timelines/timeline1.xml");
    assert_eq!(parts.timelines[0].cache_part, None);
}

