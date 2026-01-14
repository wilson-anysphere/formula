use std::io::{Cursor, Write};

use formula_xlsx::pivots::preserve_pivot_parts_from_reader;
use formula_xlsx::XlsxPackage;
use roxmltree::Document;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn count(haystack: &str, needle: &str) -> usize {
    haystack.match_indices(needle).count()
}

fn build_source_package() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCacheDefinition1.xml" ContentType="application/vnd.ms-excel.timelineCacheDefinition+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches count="1">
    <pivotCache cacheId="1" r:id="rId5"/>
  </pivotCaches>
  <slicerCaches count="1">
    <slicerCache r:id="rId10"/>
  </slicerCaches>
  <timelineCaches count="1">
    <timelineCache r:id="rId11"/>
  </timelineCaches>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId10" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rId11" Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition" Target="timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let pivot_cache_def = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1"/>"#;

    let slicer_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="SlicerCache1"/>"#;

    let timeline_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="TimelineCache1"/>"#;

    build_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache,
        ),
    ])
}

fn build_source_package_with_invalid_workbook_cache_relationships() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCacheDefinition1.xml" ContentType="application/vnd.ms-excel.timelineCacheDefinition+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    // Workbook references rId10/rId11 for slicer/timeline caches, but workbook.xml.rels is
    // intentionally missing/mismatching those relationships. Preservation should skip the
    // corresponding `<slicerCaches>`/`<timelineCaches>` blocks to avoid re-applying broken r:id
    // references.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <pivotCaches count="1">
    <pivotCache cacheId="1" r:id="rId5"/>
  </pivotCaches>
  <slicerCaches count="1">
    <slicerCache r:id="rId10"/>
  </slicerCaches>
  <timelineCaches count="1">
    <timelineCache r:id="rId11"/>
  </timelineCaches>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let pivot_cache_def = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1"/>"#;

    let slicer_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCache xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="SlicerCache1"/>"#;

    let timeline_cache = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="TimelineCache1"/>"#;

    build_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def),
        ("xl/slicerCaches/slicerCache1.xml", slicer_cache),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache,
        ),
    ])
}

fn build_destination_package() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <extLst><ext/></extLst>
</workbook>"#;

    // Deliberately consume rId10/rId11 for unrelated relationships so that applying preserved
    // slicer/timeline cache relationships must allocate fresh IDs and rewrite r:id values inside
    // the preserved `<slicerCaches>`/`<timelineCaches>` fragments.
    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#;

    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let styles_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    build_zip(&[
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/styles.xml", styles_xml),
    ])
}

#[test]
fn preserves_and_reapplies_workbook_slicer_and_timeline_cache_lists() {
    let source_bytes = build_source_package();

    let source_pkg = XlsxPackage::from_bytes(&source_bytes).expect("read source package");
    let preserved = source_pkg
        .preserve_pivot_parts()
        .expect("preserve pivot parts");

    assert!(
        preserved.workbook_slicer_caches.is_some(),
        "expected workbook <slicerCaches> subtree to be preserved"
    );
    assert!(
        preserved.workbook_timeline_caches.is_some(),
        "expected workbook <timelineCaches> subtree to be preserved"
    );
    assert_eq!(preserved.workbook_slicer_cache_rels.len(), 1);
    assert_eq!(preserved.workbook_slicer_cache_rels[0].rel_id, "rId10");
    assert_eq!(
        preserved.workbook_slicer_cache_rels[0].target,
        "slicerCaches/slicerCache1.xml"
    );
    assert_eq!(preserved.workbook_timeline_cache_rels.len(), 1);
    assert_eq!(preserved.workbook_timeline_cache_rels[0].rel_id, "rId11");
    assert_eq!(
        preserved.workbook_timeline_cache_rels[0].target,
        "timelineCaches/timelineCacheDefinition1.xml"
    );

    // Streaming variant should capture the same workbook-level cache lists.
    let preserved_streaming = preserve_pivot_parts_from_reader(Cursor::new(source_bytes.clone()))
        .expect("preserve pivot parts (streaming)");
    assert!(preserved_streaming.workbook_slicer_caches.is_some());
    assert!(preserved_streaming.workbook_timeline_caches.is_some());
    assert_eq!(
        preserved_streaming.workbook_slicer_cache_rels,
        preserved.workbook_slicer_cache_rels
    );
    assert_eq!(
        preserved_streaming.workbook_timeline_cache_rels,
        preserved.workbook_timeline_cache_rels
    );

    let dest_bytes = build_destination_package();
    let mut dest_pkg = XlsxPackage::from_bytes(&dest_bytes).expect("read destination");
    dest_pkg
        .apply_preserved_pivot_parts(&preserved)
        .expect("apply preserved pivot parts");

    let workbook_xml = std::str::from_utf8(dest_pkg.part("xl/workbook.xml").unwrap()).unwrap();
    Document::parse(workbook_xml).expect("output workbook.xml should be parseable XML");

    assert_eq!(count(workbook_xml, "<slicerCaches"), 1);
    assert_eq!(count(workbook_xml, "<timelineCaches"), 1);

    // rId10 and rId11 were consumed in the destination; preserved r:id values should have been
    // rewritten to avoid conflicts.
    assert!(workbook_xml.contains(r#"<slicerCache r:id="rId12""#));
    assert!(workbook_xml.contains(r#"<timelineCache r:id="rId13""#));
    assert!(!workbook_xml.contains(r#"<slicerCache r:id="rId10""#));

    // Best-effort ordering: insert cache lists before <extLst>.
    let slicer_pos = workbook_xml.find("<slicerCaches").unwrap();
    let timeline_pos = workbook_xml.find("<timelineCaches").unwrap();
    let ext_pos = workbook_xml.find("<extLst").unwrap();
    assert!(slicer_pos < timeline_pos);
    assert!(timeline_pos < ext_pos);

    let workbook_rels = std::str::from_utf8(dest_pkg.part("xl/_rels/workbook.xml.rels").unwrap())
        .unwrap();
    assert!(
        workbook_rels.contains(
            r#"Relationship Id="rId12" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml""#
        ),
        "missing slicerCache relationship: {workbook_rels}"
    );
    assert!(
        workbook_rels.contains(
            r#"Relationship Id="rId13" Type="http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition" Target="timelineCaches/timelineCacheDefinition1.xml""#
        ),
        "missing timelineCacheDefinition relationship: {workbook_rels}"
    );

    // Ensure the destination relationships we conflicted with remain intact.
    assert!(
        workbook_rels.contains(r#"Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""#),
        "expected destination styles relationship to remain on rId10"
    );
    assert!(
        workbook_rels.contains(r#"Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme""#),
        "expected destination theme relationship to remain on rId11"
    );
}

#[test]
fn skips_workbook_cache_lists_when_relationships_are_missing_or_wrong_type() {
    let source_bytes = build_source_package_with_invalid_workbook_cache_relationships();
    let source_pkg = XlsxPackage::from_bytes(&source_bytes).expect("read source package");

    let preserved = source_pkg
        .preserve_pivot_parts()
        .expect("preserve pivot parts");
    assert!(
        preserved.workbook_pivot_caches.is_some(),
        "pivotCaches should still be preserved when its rels are valid"
    );
    assert!(
        preserved.workbook_slicer_caches.is_none(),
        "slicerCaches should be skipped when workbook rels are invalid"
    );
    assert!(
        preserved.workbook_timeline_caches.is_none(),
        "timelineCaches should be skipped when workbook rels are invalid"
    );
    assert!(preserved.workbook_slicer_cache_rels.is_empty());
    assert!(preserved.workbook_timeline_cache_rels.is_empty());

    // Raw parts are still preserved; we only avoid re-applying a broken workbook.xml subtree.
    assert!(preserved.parts.contains_key("xl/slicerCaches/slicerCache1.xml"));
    assert!(preserved
        .parts
        .contains_key("xl/timelineCaches/timelineCacheDefinition1.xml"));

    let preserved_streaming =
        preserve_pivot_parts_from_reader(Cursor::new(source_bytes)).expect("streaming preserve");
    assert!(preserved_streaming.workbook_pivot_caches.is_some());
    assert!(preserved_streaming.workbook_slicer_caches.is_none());
    assert!(preserved_streaming.workbook_timeline_caches.is_none());
    assert!(preserved_streaming.workbook_slicer_cache_rels.is_empty());
    assert!(preserved_streaming.workbook_timeline_cache_rels.is_empty());
}
