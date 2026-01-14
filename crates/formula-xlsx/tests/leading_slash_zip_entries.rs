use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue, Range};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming, read_workbook_model_from_bytes, worksheet_parts_from_reader,
    WorksheetCellPatch, XlsxPackage,
};
use formula_xlsx::print::{read_workbook_print_settings, write_workbook_print_settings, CellRange};
use formula_xlsx::pivots::preserve_pivot_parts_from_reader;
use formula_xlsx::drawingml::preserve_drawing_parts_from_reader;
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

fn build_minimal_xlsx_with_leading_slash_entries() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <mergeCells count="1">
    <mergeCell ref="A1:B2"/>
  </mergeCells>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "/xl/workbook.xml", workbook_xml);
    add_file(
        &mut zip,
        options,
        "/xl/_rels/workbook.xml.rels",
        workbook_rels,
    );
    add_file(&mut zip, options, "/xl/worksheets/sheet1.xml", worksheet_xml);

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_leading_slash_pivot_part() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

    let pivot_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "/[Content_Types].xml", content_types);
    add_file(&mut zip, options, "/xl/workbook.xml", workbook_xml);
    add_file(
        &mut zip,
        options,
        "/xl/pivotTables/pivotTable1.xml",
        pivot_xml,
    );

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_leading_slash_drawing_part() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>"#;

    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing1.xml"/>
</Relationships>"#;

    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "/[Content_Types].xml", content_types);
    add_file(&mut zip, options, "/xl/workbook.xml", workbook_xml);
    add_file(
        &mut zip,
        options,
        "/xl/_rels/workbook.xml.rels",
        workbook_rels,
    );
    add_file(&mut zip, options, "/xl/worksheets/sheet1.xml", worksheet_xml);
    add_file(
        &mut zip,
        options,
        "/xl/worksheets/_rels/sheet1.xml.rels",
        sheet_rels,
    );
    add_file(&mut zip, options, "/xl/drawings/drawing1.xml", drawing_xml);

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_noncanonical_entries() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    // Non-canonical entry names:
    // - Different case (`XL/Workbook.xml`)
    // - Windows-style path separators (`xl\worksheets\sheet1.xml`)
    add_file(&mut zip, options, "XL/Workbook.xml", workbook_xml);
    add_file(&mut zip, options, "xl/_rels/workbook.xml.rels", workbook_rels);
    add_file(
        &mut zip,
        options,
        r#"xl\worksheets\sheet1.xml"#,
        worksheet_xml,
    );

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_noncanonical_sheet_part_and_rels() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1"/>
  </hyperlinks>
</worksheet>"#;

    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com"
    TargetMode="External"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "XL/Workbook.xml", workbook_xml);
    add_file(&mut zip, options, "xl/_rels/workbook.xml.rels", workbook_rels);
    add_file(
        &mut zip,
        options,
        r#"xl\worksheets\sheet1.xml"#,
        worksheet_xml,
    );
    add_file(
        &mut zip,
        options,
        "xl/worksheets/_rels/sheet1.xml.rels",
        sheet_rels,
    );

    zip.finish().unwrap().into_inner()
}

// The CI filter uses `-- leading_slash_zip_entries`; wrapping tests in this module ensures the
// substring matches and the intended subset runs.
mod leading_slash_zip_entries_tests {
    use super::*;

    #[test]
    fn worksheet_parts_from_reader_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let parts = worksheet_parts_from_reader(Cursor::new(bytes)).expect("worksheet parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].name, "Sheet1");
        assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn worksheet_parts_from_reader_tolerates_noncanonical_zip_entries() {
        let bytes = build_minimal_xlsx_with_noncanonical_entries();
        let parts = worksheet_parts_from_reader(Cursor::new(bytes)).expect("worksheet parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].name, "Sheet1");
        assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn load_from_bytes_tolerates_noncanonical_worksheet_part_names_when_sheet_rels_are_required() {
        use formula_model::HyperlinkTarget;

        let bytes = build_minimal_xlsx_with_noncanonical_sheet_part_and_rels();
        let doc = load_from_bytes(&bytes).expect("load xlsx document");
        assert_eq!(doc.workbook.sheets.len(), 1);
        let sheet = &doc.workbook.sheets[0];
        assert_eq!(sheet.hyperlinks.len(), 1);
        assert_eq!(sheet.hyperlinks[0].range.to_string(), "A1");
        match &sheet.hyperlinks[0].target {
            HyperlinkTarget::ExternalUrl { uri } => assert_eq!(uri, "https://example.com"),
            other => panic!("expected external hyperlink target (got {other:?})"),
        }
    }

    #[test]
    fn read_workbook_model_from_bytes_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
    }

    #[test]
    fn load_from_bytes_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let doc = load_from_bytes(&bytes).expect("load xlsx document");
        assert_eq!(doc.workbook.sheets.len(), 1);
        assert_eq!(doc.workbook.sheets[0].name, "Sheet1");
    }

    #[test]
    fn merge_cells_reader_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("zip");
        let merges = formula_xlsx::merge_cells::read_merge_cells_from_xlsx(
            &mut archive,
            "xl/worksheets/sheet1.xml",
        )
        .expect("merge cells");
        assert_eq!(merges, vec![Range::from_a1("A1:B2").unwrap()]);
    }

    #[test]
    fn streaming_patcher_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let patch = WorksheetCellPatch::new(
            "xl/worksheets/sheet1.xml",
            CellRef::new(0, 0),
            CellValue::Number(42.0),
            None,
        );
        let mut out = Cursor::new(Vec::new());
        patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch]).expect("streaming patch");

        let pkg = XlsxPackage::from_bytes(&out.into_inner()).expect("read patched package");
        let sheet_xml = std::str::from_utf8(
            pkg.part("xl/worksheets/sheet1.xml")
                .expect("worksheet part present"),
        )
        .expect("worksheet xml utf-8");
        assert!(
            sheet_xml.contains("<v>42</v>") || sheet_xml.contains("<v>42.0</v>"),
            "expected patched worksheet XML to contain cell value 42 (got {sheet_xml:?})"
        );
    }

    #[test]
    fn print_settings_writer_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_entries();
        let mut settings = read_workbook_print_settings(&bytes).expect("read print settings");
        assert_eq!(settings.sheets.len(), 1);

        settings.sheets[0].print_area = Some(vec![CellRange {
            start_row: 1,
            end_row: 2,
            start_col: 1,
            end_col: 2,
        }]);

        let rewritten =
            write_workbook_print_settings(&bytes, &settings).expect("write print settings");
        let reread = read_workbook_print_settings(&rewritten).expect("re-read print settings");
        assert_eq!(
            reread.sheets[0].print_area.as_deref(),
            settings.sheets[0].print_area.as_deref()
        );
    }

    #[test]
    fn preserve_pivot_parts_from_reader_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_pivot_part();
        let preserved = preserve_pivot_parts_from_reader(Cursor::new(bytes)).expect("preserve pivots");
        assert!(
            preserved
                .parts
                .contains_key("xl/pivotTables/pivotTable1.xml"),
            "expected preserved pivot parts to include xl/pivotTables/pivotTable1.xml"
        );
    }

    #[test]
    fn preserve_drawing_parts_from_reader_tolerates_leading_slash_entries() {
        let bytes = build_minimal_xlsx_with_leading_slash_drawing_part();
        let preserved = preserve_drawing_parts_from_reader(Cursor::new(bytes)).expect("preserve drawings");
        assert!(
            preserved.parts.contains_key("xl/drawings/drawing1.xml"),
            "expected preserved drawing parts to include xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn read_workbook_model_from_bytes_tolerates_noncanonical_zip_entries() {
        let bytes = build_minimal_xlsx_with_noncanonical_entries();
        let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
        assert_eq!(workbook.sheets.len(), 1);
        assert_eq!(workbook.sheets[0].name, "Sheet1");
    }

    #[test]
    fn load_from_bytes_tolerates_noncanonical_zip_entries() {
        let bytes = build_minimal_xlsx_with_noncanonical_entries();
        let doc = load_from_bytes(&bytes).expect("load xlsx document");
        assert_eq!(doc.workbook.sheets.len(), 1);
        assert_eq!(doc.workbook.sheets[0].name, "Sheet1");
    }

    #[test]
    fn xlsx_package_part_lookup_tolerates_noncanonical_zip_entries() {
        let bytes = build_minimal_xlsx_with_noncanonical_entries();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
        let workbook_xml =
            std::str::from_utf8(pkg.part("xl/workbook.xml").expect("workbook.xml part"))
                .expect("workbook.xml utf-8");
        assert!(
            workbook_xml.contains("<sheet name=\"Sheet1\""),
            "expected workbook.xml to contain sheet, got:\n{workbook_xml}"
        );
    }
}
