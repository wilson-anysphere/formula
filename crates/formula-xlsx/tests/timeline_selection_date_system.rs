use formula_xlsx::XlsxPackage;
use std::io::{Cursor, Write};
use zip::write::FileOptions;
use zip::ZipWriter;

fn parse_timeline_selection_from_parts(
    workbook_xml: &[u8],
    start_date: &str,
    end_date: &str,
) -> Result<(Option<String>, Option<String>), Box<dyn std::error::Error>> {
    let timeline_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timeline xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          name="Timeline1">
  <timelineCache r:id="rId1"/>
</timeline>"#;

    let timeline_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache"
                Target="../timelineCaches/timelineCacheDefinition1.xml"/>
</Relationships>"#;

    let timeline_cache_definition_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <selection startDate="{start_date}" endDate="{end_date}"/>
</timelineCacheDefinition>"#
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("xl/workbook.xml", workbook_xml),
        ("xl/timelines/timeline1.xml", timeline_xml.as_slice()),
        ("xl/timelines/_rels/timeline1.xml.rels", timeline_rels.as_slice()),
        (
            "xl/timelineCaches/timelineCacheDefinition1.xml",
            timeline_cache_definition_xml.as_bytes(),
        ),
    ] {
        zip.start_file(name, options)?;
        zip.write_all(bytes)?;
    }

    let bytes = zip.finish()?.into_inner();
    let package = XlsxPackage::from_bytes(&bytes)?;
    let parts = package.pivot_slicer_parts()?;

    assert_eq!(parts.timelines.len(), 1);
    let timeline = &parts.timelines[0];
    Ok((timeline.selection.start.clone(), timeline.selection.end.clone()))
}

#[test]
fn timeline_selection_respects_workbook_date_system_1904() -> Result<(), Box<dyn std::error::Error>>
{
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <workbookPr date1904="1"/>
</workbook>"#;

    let (start, end) = parse_timeline_selection_from_parts(workbook_xml.as_slice(), "1", "2")?;
    assert_eq!(start.as_deref(), Some("1904-01-02"));
    assert_eq!(end.as_deref(), Some("1904-01-03"));

    Ok(())
}

#[test]
fn timeline_selection_defaults_to_1900_date_system() -> Result<(), Box<dyn std::error::Error>> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <workbookPr/>
</workbook>"#;

    let (start, end) = parse_timeline_selection_from_parts(workbook_xml.as_slice(), "1", "2")?;
    assert_eq!(start.as_deref(), Some("1900-01-01"));
    assert_eq!(end.as_deref(), Some("1900-01-02"));

    Ok(())
}

#[test]
fn timeline_selection_preserves_iso_and_compact_ymd_inputs() -> Result<(), Box<dyn std::error::Error>>
{
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <workbookPr date1904="1"/>
</workbook>"#;

    let (start, end) = parse_timeline_selection_from_parts(
        workbook_xml.as_slice(),
        "2024-02-03T12:34:56Z",
        "20240204",
    )?;
    assert_eq!(start.as_deref(), Some("2024-02-03"));
    assert_eq!(end.as_deref(), Some("2024-02-04"));

    Ok(())
}
