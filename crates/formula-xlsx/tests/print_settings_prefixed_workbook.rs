use std::io::{Cursor, Read, Write};

use formula_xlsx::print::{
    read_workbook_print_settings, write_workbook_print_settings, CellRange, ColRange, PrintTitles,
    RowRange,
};

fn build_prefixed_workbook_xlsx(workbook_xml: &str) -> Vec<u8> {
    // Prefixed `<Relationship>` elements.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_prefixed_workbook_xlsx() -> Vec<u8> {
    // Prefix-only SpreadsheetML (no default `xmlns`) with a non-`r` relationships prefix.
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;
    build_prefixed_workbook_xlsx(workbook_xml)
}

#[test]
fn print_settings_roundtrip_with_prefix_only_workbook_and_prefixed_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let original = build_minimal_prefixed_workbook_xlsx();

    let mut settings = read_workbook_print_settings(&original)?;
    assert_eq!(settings.sheets.len(), 1);
    assert_eq!(settings.sheets[0].sheet_name, "Sheet1");

    settings.sheets[0].print_area = Some(vec![CellRange {
        start_row: 1,
        end_row: 2,
        start_col: 1,
        end_col: 3,
    }]);
    settings.sheets[0].print_titles = Some(PrintTitles {
        repeat_rows: Some(RowRange { start: 1, end: 1 }),
        repeat_cols: Some(ColRange { start: 2, end: 2 }),
    });

    let rewritten = write_workbook_print_settings(&original, &settings)?;
    let reread = read_workbook_print_settings(&rewritten)?;

    assert_eq!(
        reread.sheets[0].print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 2,
                start_col: 1,
                end_col: 3
            }][..]
        )
    );
    assert_eq!(
        reread.sheets[0].print_titles,
        settings.sheets[0].print_titles
    );

    // Verify we didn't introduce namespace-less `<definedNames>`/`<definedName>` into a prefix-only workbook.
    let mut zip = zip::ZipArchive::new(Cursor::new(&rewritten))?;
    let mut workbook_file = zip.by_name("xl/workbook.xml")?;
    let mut workbook_xml = String::new();
    workbook_file.read_to_string(&mut workbook_xml)?;

    roxmltree::Document::parse(&workbook_xml)?;
    assert!(
        workbook_xml.contains("<x:definedNames") && workbook_xml.contains("<x:definedName"),
        "expected prefixed definedNames/definedName, got:\n{workbook_xml}"
    );
    assert!(
        !workbook_xml.contains("<definedNames") && !workbook_xml.contains("<definedName"),
        "should not introduce unprefixed definedNames/definedName, got:\n{workbook_xml}"
    );

    Ok(())
}

#[test]
fn print_settings_update_existing_defined_names_preserves_prefix(
) -> Result<(), Box<dyn std::error::Error>> {
    // Workbook already contains a prefixed <x:definedNames> block with a prefixed <x:definedName>.
    // This exercises the "update existing definedName" path and ensures we don't emit mismatched
    // end tags like </definedName> for a <x:definedName> start tag.
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
  <x:definedNames>
    <x:definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</x:definedName>
  </x:definedNames>
</x:workbook>"#;

    let original = build_prefixed_workbook_xlsx(workbook_xml);
    let original_settings = read_workbook_print_settings(&original)?;
    assert_eq!(
        original_settings.sheets[0].print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 2,
                start_col: 1,
                end_col: 2,
            }][..]
        )
    );

    // Update print area (existing definedName) and add print titles (new definedName).
    let mut updated_settings = original_settings.clone();
    updated_settings.sheets[0].print_area = Some(vec![CellRange {
        start_row: 3,
        end_row: 4,
        start_col: 2,
        end_col: 5,
    }]);
    updated_settings.sheets[0].print_titles = Some(PrintTitles {
        repeat_rows: Some(RowRange { start: 1, end: 1 }),
        repeat_cols: Some(ColRange { start: 2, end: 2 }),
    });

    let rewritten = write_workbook_print_settings(&original, &updated_settings)?;
    let reread = read_workbook_print_settings(&rewritten)?;
    assert_eq!(
        reread.sheets[0].print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 3,
                end_row: 4,
                start_col: 2,
                end_col: 5
            }][..]
        )
    );
    assert_eq!(
        reread.sheets[0].print_titles,
        updated_settings.sheets[0].print_titles
    );

    // Verify workbook.xml is valid XML and contains only prefixed definedNames/definedName tags.
    let mut zip = zip::ZipArchive::new(Cursor::new(&rewritten))?;
    let mut workbook_file = zip.by_name("xl/workbook.xml")?;
    let mut rewritten_workbook_xml = String::new();
    workbook_file.read_to_string(&mut rewritten_workbook_xml)?;

    roxmltree::Document::parse(&rewritten_workbook_xml)?;
    assert!(
        rewritten_workbook_xml.contains("name=\"_xlnm.Print_Area\"")
            && rewritten_workbook_xml.contains("name=\"_xlnm.Print_Titles\""),
        "expected both print-related definedNames, got:\n{rewritten_workbook_xml}"
    );
    assert!(
        rewritten_workbook_xml.contains("<x:definedNames")
            && rewritten_workbook_xml.contains("<x:definedName")
            && rewritten_workbook_xml.contains("</x:definedName>"),
        "expected prefixed definedNames/definedName with matching end tag, got:\n{rewritten_workbook_xml}"
    );
    assert!(
        !rewritten_workbook_xml.contains("<definedNames")
            && !rewritten_workbook_xml.contains("<definedName")
            && !rewritten_workbook_xml.contains("</definedName>"),
        "should not introduce unprefixed definedNames/definedName tags, got:\n{rewritten_workbook_xml}"
    );

    Ok(())
}
