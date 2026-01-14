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
fn resolves_timeline_field_name_from_base_field_and_pivot_cache() -> Result<(), Box<dyn std::error::Error>>
{
    // Synthetic package:
    // - Pivot cache has 2 fields: Date and Region.
    // - Timeline cache sets baseField="0" (Date) and is connected to pivotTable1.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let pivot_table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" cacheId="1" name="PivotTable1"/>"#;

    let pivot_cache_def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B1" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Date"/>
    <cacheField name="Region"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition"
                Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    let timeline_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         name="TimelineCache1"
                         sourceName="PivotTable1"
                         baseField="0"
                         level="0">
  <pivotTable id="rId1"/>
</timelineCacheDefinition>"#;

    let timeline_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
                Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let package = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
        ("xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def_xml),
        ("xl/timelines/timeline1.xml", timeline_xml),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels),
        ("xl/timelineCaches/timelineCacheDefinition1.xml", timeline_cache_xml),
        (
            "xl/timelineCaches/_rels/timelineCacheDefinition1.xml.rels",
            timeline_cache_rels,
        ),
    ]);

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.timelines.len(), 1);
    assert_eq!(parts.timelines[0].field_name.as_deref(), Some("Date"));

    Ok(())
}

