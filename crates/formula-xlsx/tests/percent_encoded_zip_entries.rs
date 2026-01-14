use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    patch_xlsx_streaming, worksheet_parts_from_reader, worksheet_parts_from_reader_limited,
    patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches, WorksheetCellPatch,
    StreamingXlsxPackage, XlsxPackage,
};
use zip::ZipArchive;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_minimal_xlsx_with_percent_encoded_sheet_part() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    // Intentionally use an unescaped space in the relationship Target (producer bug), while the
    // stored ZIP entry name is percent-encoded.
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
    Target="worksheets/sheet 1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types).unwrap();
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();
    // Store the worksheet part name percent-encoded.
    zip.start_file("xl/worksheets/sheet%201.xml", options)
        .unwrap();
    zip.write_all(worksheet_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsx_with_percent_encoded_drawing_part() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
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
    Target="worksheets/sheet 1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rIdDrawing"/>
</worksheet>"#;

    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDrawing"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing 1.xml"/>
</Relationships>"#;

    // Store the drawing part name percent-encoded.
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types).unwrap();
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();
    zip.start_file("xl/worksheets/sheet%201.xml", options)
        .unwrap();
    zip.write_all(worksheet_xml).unwrap();
    zip.start_file("xl/worksheets/_rels/sheet%201.xml.rels", options)
        .unwrap();
    zip.write_all(sheet_rels).unwrap();
    zip.start_file("xl/drawings/drawing%201.xml", options)
        .unwrap();
    zip.write_all(drawing_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_minimal_xlsm_with_percent_encoded_dependency() -> Vec<u8> {
    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
    Target="../embeddings/oleObject 1.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(b"dummy-vba").unwrap();
    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(vba_rels).unwrap();
    // Store the target part percent-encoded, but reference it unescaped in the relationship.
    zip.start_file("xl/embeddings/oleObject%201.bin", options)
        .unwrap();
    zip.write_all(b"dummy-ole").unwrap();

    zip.finish().unwrap().into_inner()
}

// The CI filter uses `-- percent_encoded_zip_entries`; wrapping tests in this module ensures the
// substring matches and the intended subset runs.
mod percent_encoded_zip_entries_tests {
    use super::*;

    #[test]
    fn worksheet_parts_from_reader_finds_percent_encoded_sheet_entries() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();
        let parts = worksheet_parts_from_reader(Cursor::new(bytes)).expect("worksheet parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].name, "Sheet1");
        assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet%201.xml");
    }

    #[test]
    fn worksheet_parts_from_reader_limited_finds_percent_encoded_sheet_entries() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();
        let parts = worksheet_parts_from_reader_limited(Cursor::new(bytes), 1024)
            .expect("worksheet parts");
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].name, "Sheet1");
        assert_eq!(parts[0].worksheet_part, "xl/worksheets/sheet%201.xml");
    }

    #[test]
    fn streaming_patcher_accepts_percent_encoded_sheet_part_names() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();
        let parts = worksheet_parts_from_reader(Cursor::new(bytes.clone())).expect("worksheet parts");
        let part = parts[0].worksheet_part.clone();

        let patch = WorksheetCellPatch::new(
            &part,
            CellRef::new(0, 0),
            CellValue::Number(42.0),
            None,
        );
        let mut out = Cursor::new(Vec::new());
        patch_xlsx_streaming(Cursor::new(bytes), &mut out, &[patch]).expect("streaming patch");

        let pkg = XlsxPackage::from_bytes(&out.into_inner()).expect("read patched package");
        let sheet_xml = std::str::from_utf8(pkg.part(&part).expect("worksheet part present"))
            .expect("worksheet xml utf-8");
        assert!(
            sheet_xml.contains("<v>42</v>") || sheet_xml.contains("<v>42.0</v>"),
            "expected patched worksheet XML to contain cell value 42 (got {sheet_xml:?})"
        );
    }

    #[test]
    fn workbook_cell_patches_resolve_to_percent_encoded_sheet_entries() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();

        let mut patches = WorkbookCellPatches::default();
        patches.set_cell(
            "Sheet1",
            CellRef::new(0, 0),
            CellPatch::set_value(CellValue::Number(42.0)),
        );

        let mut out = Cursor::new(Vec::new());
        patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)
            .expect("workbook streaming patch");

        let pkg = XlsxPackage::from_bytes(&out.into_inner()).expect("read patched package");
        let sheet_xml =
            std::str::from_utf8(pkg.part("xl/worksheets/sheet%201.xml").expect("worksheet part"))
                .expect("worksheet xml utf-8");
        assert!(
            sheet_xml.contains("<v>42</v>") || sheet_xml.contains("<v>42.0</v>"),
            "expected patched worksheet XML to contain cell value 42 (got {sheet_xml:?})"
        );
    }

    #[test]
    fn preserve_pivot_parts_from_reader_discovers_sheet_with_percent_encoded_zip_name() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();
        let preserved =
            formula_xlsx::pivots::preserve_pivot_parts_from_reader(Cursor::new(bytes))
                .expect("preserve pivot parts");
        assert_eq!(preserved.workbook_sheets.len(), 1);
        assert_eq!(preserved.workbook_sheets[0].name, "Sheet1");
        assert_eq!(preserved.workbook_sheets[0].index, 0);
    }

    #[test]
    fn streaming_xlsx_package_set_part_maps_to_percent_encoded_zip_entry_name() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();
        let mut pkg =
            StreamingXlsxPackage::from_reader(Cursor::new(bytes)).expect("open streaming package");

        let replacement = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>99</v></c></row>
  </sheetData>
</worksheet>"#;
        // Intentionally use the unescaped name; the source ZIP entry uses `%20`.
        pkg.set_part("xl/worksheets/sheet 1.xml", replacement.to_vec());

        let mut out = Cursor::new(Vec::new());
        pkg.write_to(&mut out).expect("write streaming package");

        let out_bytes = out.into_inner();
        let mut archive = ZipArchive::new(Cursor::new(out_bytes)).expect("open output zip");
        let names = archive
            .file_names()
            .map(|name| name.to_string())
            .collect::<Vec<_>>();
        assert!(
            names.iter().any(|n| n == "xl/worksheets/sheet%201.xml"),
            "expected output to contain the original percent-encoded sheet entry, got {names:?}"
        );
        assert!(
            !names.iter().any(|n| n == "xl/worksheets/sheet 1.xml"),
            "expected output to not contain a new unescaped sheet entry, got {names:?}"
        );

        let mut sheet = archive
            .by_name("xl/worksheets/sheet%201.xml")
            .expect("open patched sheet");
        let mut xml = String::new();
        sheet.read_to_string(&mut xml).expect("read sheet xml");
        assert!(
            xml.contains("<v>99</v>") || xml.contains("<v>99.0</v>"),
            "expected patched sheet XML to contain 99 (got {xml:?})"
        );
    }

    #[test]
    fn preserve_drawing_parts_from_reader_finds_percent_encoded_drawing_entries() {
        let bytes = build_minimal_xlsx_with_percent_encoded_drawing_part();
        let preserved = formula_xlsx::drawingml::preserve_drawing_parts_from_reader(Cursor::new(bytes))
            .expect("preserve drawing parts");
        assert!(
            preserved.parts.contains_key("xl/drawings/drawing%201.xml"),
            "expected preserved parts to include percent-encoded drawing1.xml; got keys: {:?}",
            preserved.parts.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn macro_strip_streaming_deletes_percent_encoded_dependency_targets() {
        let bytes = build_minimal_xlsm_with_percent_encoded_dependency();
        let mut out = Cursor::new(Vec::new());
        formula_xlsx::strip_vba_project_streaming(Cursor::new(bytes), &mut out)
            .expect("strip macros");
        let out_bytes = out.into_inner();
        let archive = ZipArchive::new(Cursor::new(out_bytes)).expect("open stripped zip");
        assert_eq!(
            archive.len(),
            0,
            "expected all parts to be removed from the minimal macro test workbook"
        );
    }

    #[test]
    fn write_workbook_print_settings_updates_percent_encoded_sheet_entries() {
        let bytes = build_minimal_xlsx_with_percent_encoded_sheet_part();

        let settings = formula_xlsx::print::WorkbookPrintSettings {
            sheets: vec![formula_xlsx::print::SheetPrintSettings {
                sheet_name: "Sheet1".to_string(),
                print_area: None,
                print_titles: None,
                page_setup: formula_xlsx::print::PageSetup {
                    orientation: formula_xlsx::print::Orientation::Landscape,
                    ..Default::default()
                },
                manual_page_breaks: Default::default(),
            }],
        };

        let out_bytes = formula_xlsx::print::write_workbook_print_settings(&bytes, &settings)
            .expect("write print settings");

        let mut archive = ZipArchive::new(Cursor::new(out_bytes)).expect("open output zip");
        let mut sheet = archive
            .by_name("xl/worksheets/sheet%201.xml")
            .expect("open sheet");
        let mut xml = String::new();
        sheet.read_to_string(&mut xml).expect("read sheet xml");
        assert!(
            xml.contains("orientation=\"landscape\""),
            "expected worksheet XML to include landscape orientation (got {xml:?})"
        );
    }
}
