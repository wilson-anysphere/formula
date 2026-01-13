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
fn infers_slicer_field_name_from_pivot_cache_values() -> Result<(), Box<dyn std::error::Error>> {
    // Synthetic package:
    // - Pivot cache has 2 text fields: Color and Shape.
    // - Slicer cache items are `Red`/`Blue`, which match only the Color field.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let pivot_table_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" cacheId="1" name="PivotTable1"/>"#;

    let pivot_cache_def_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="3">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B4" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Color"/>
    <cacheField name="Shape"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    let pivot_cache_records_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <r><s v="Red"/><s v="Circle"/></r>
  <r><s v="Blue"/><s v="Square"/></r>
  <r><s v="Red"/><s v="Triangle"/></r>
</pivotCacheRecords>"#;

    let slicer_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="ColorSlicer">
  <slicerCache r:id="rId1"/>
</slicer>"#;

    let slicer_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="../slicerCaches/slicerCache1.xml"/>
</Relationships>"#;

    let slicer_cache_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="ColorSlicerCache" sourceName="PivotTable1">
  <slicerCachePivotTables>
    <slicerCachePivotTable r:id="rId1"/>
  </slicerCachePivotTables>
  <slicerCacheData>
    <slicerCacheItem n="Red" s="1"/>
    <slicerCacheItem n="Blue" s="1"/>
  </slicerCacheData>
</slicerCache>"#;

    let slicer_cache_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>"#;

    let package = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/pivotTables/pivotTable1.xml", pivot_table_xml),
        ("xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def_xml),
        ("xl/pivotCache/pivotCacheRecords1.xml", pivot_cache_records_xml),
        ("xl/slicers/slicer1.xml", slicer_xml),
        ("xl/slicers/_rels/slicer1.xml.rels", slicer_rels),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache_xml),
        (
            "xl/slicerCaches/_rels/slicerCache1.xml.rels",
            slicer_cache_rels,
        ),
    ]);

    let parts = package.pivot_slicer_parts()?;
    assert_eq!(parts.slicers.len(), 1);
    assert_eq!(parts.slicers[0].field_name.as_deref(), Some("Color"));

    Ok(())
}

