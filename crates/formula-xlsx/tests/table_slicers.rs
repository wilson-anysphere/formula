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
fn parses_table_slicer_connected_tables_from_cache_relationships() -> Result<(), Box<dyn std::error::Error>>
{
    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  name="TableSlicer1" uid="{00000000-0000-0000-0000-000000000001}">
  <slicerCache r:id="rId1"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache"
    Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    // Table slicer caches often do not have explicit pivot-table references; they instead point to
    // the table part through a relationship of type `.../table`.
    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  name="TableSlicerCache1" sourceName="Table1"/>"#;

    let slicer_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
    Target="../tables/table1.xml"/>
</Relationships>"#;

    let table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  id="1" name="Table1" displayName="Table1" ref="A1:A1" totalsRowCount="0">
  <autoFilter ref="A1:A1"/>
  <tableColumns count="1">
    <tableColumn id="1" name="Column1"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
    showRowStripes="1" showColumnStripes="0"/>
</table>"#;

    let package = build_package(&[
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        ("xl/slicerCaches/_rels/slicerCache1.xml.rels", slicer_cache_rels),
        ("xl/tables/table1.xml", table_xml),
    ]);

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(
        parts.slicers[0].connected_tables,
        vec!["xl/tables/table1.xml".to_string()]
    );

    Ok(())
}

