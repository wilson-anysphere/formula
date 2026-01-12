use std::io::{Cursor, Read, Write};

use formula_model::{CellValue, ErrorValue, Workbook};
use formula_xlsx::{read_workbook_model_from_bytes, XlsxDocument};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

fn worksheet_cell_type_and_value(xml: &str, a1: &str) -> Option<(Option<String>, Option<String>)> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_target_cell = false;
    let mut in_v = false;
    let mut t_attr: Option<String> = None;
    let mut v_text: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf).ok()? {
            Event::Start(e) if e.local_name().as_ref() == b"c" => {
                in_target_cell = false;
                in_v = false;
                t_attr = None;
                v_text = None;

                let mut r_attr: Option<String> = None;
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"r" {
                        r_attr = Some(attr.unescape_value().ok()?.into_owned());
                    }
                    if attr.key.as_ref() == b"t" {
                        t_attr = Some(attr.unescape_value().ok()?.into_owned());
                    }
                }

                if r_attr.as_deref() == Some(a1) {
                    in_target_cell = true;
                } else {
                    t_attr = None;
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                // Empty cell record (no children). We still want to capture its `t=` attribute.
                let mut r_attr: Option<String> = None;
                let mut t: Option<String> = None;
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"r" {
                        r_attr = Some(attr.unescape_value().ok()?.into_owned());
                    }
                    if attr.key.as_ref() == b"t" {
                        t = Some(attr.unescape_value().ok()?.into_owned());
                    }
                }
                if r_attr.as_deref() == Some(a1) {
                    return Some((t, None));
                }
            }
            Event::Start(e) if in_target_cell && e.local_name().as_ref() == b"v" => {
                in_v = true;
            }
            Event::Text(e) if in_target_cell && in_v => {
                v_text = Some(e.unescape().ok()?.into_owned());
            }
            Event::End(e) if in_target_cell && e.local_name().as_ref() == b"v" => {
                in_v = false;
            }
            Event::End(e) if in_target_cell && e.local_name().as_ref() == b"c" => {
                return Some((t_attr, v_text));
            }
            Event::Eof => return None,
            _ => {}
        }
        buf.clear();
    }
}

#[test]
fn xlsx_error_values_roundtrip_for_newer_excel_errors() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet
            .set_value_a1("A1", CellValue::Error(ErrorValue::GettingData))
            .expect("set A1");
        sheet
            .set_value_a1("B1", CellValue::Error(ErrorValue::Field))
            .expect("set B1");
        sheet
            .set_value_a1("C1", CellValue::Error(ErrorValue::Connect))
            .expect("set C1");
        sheet
            .set_value_a1("D1", CellValue::Error(ErrorValue::Blocked))
            .expect("set D1");
        sheet
            .set_value_a1("E1", CellValue::Error(ErrorValue::Unknown))
            .expect("set E1");
    }

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("write xlsx");

    // Spot-check on-disk XML uses the error cell type and correct `<v>` text.
    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open zip");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml")
        .read_to_string(&mut sheet_xml)
        .expect("read sheet1.xml");

    let (t_attr, v_text) =
        worksheet_cell_type_and_value(&sheet_xml, "A1").expect("A1 should exist in sheet xml");
    assert_eq!(t_attr.as_deref(), Some("e"));
    assert_eq!(v_text.as_deref(), Some("#GETTING_DATA"));

    // Round-trip through the fast reader.
    let loaded = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    let sheet = &loaded.sheets[0];

    assert_eq!(
        sheet.value_a1("A1").expect("A1"),
        CellValue::Error(ErrorValue::GettingData)
    );
    assert_eq!(
        sheet.value_a1("B1").expect("B1"),
        CellValue::Error(ErrorValue::Field)
    );
    assert_eq!(
        sheet.value_a1("C1").expect("C1"),
        CellValue::Error(ErrorValue::Connect)
    );
    assert_eq!(
        sheet.value_a1("D1").expect("D1"),
        CellValue::Error(ErrorValue::Blocked)
    );
    assert_eq!(
        sheet.value_a1("E1").expect("E1"),
        CellValue::Error(ErrorValue::Unknown)
    );
}

#[test]
fn xlsx_error_value_na_exclamation_alias_parses_as_na() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value_a1("A1", CellValue::Error(ErrorValue::NA))
        .expect("set A1");

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("write xlsx");

    // Excel sometimes serializes `#N/A` as `#N/A!` in SpreadsheetML. Rewrite the
    // sheet XML in-place to simulate a workbook produced by Excel.
    let mut archive = ZipArchive::new(Cursor::new(&bytes)).expect("open zip");
    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("read zip entry");
        if file.is_dir() {
            continue;
        }

        let name = file.name().to_string();
        let mut data = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut data).expect("read zip entry bytes");

        if name == "xl/worksheets/sheet1.xml" {
            let xml = String::from_utf8(data).expect("sheet xml should be utf-8");
            assert!(
                xml.contains("<v>#N/A</v>"),
                "expected canonical #N/A error in baseline sheet xml"
            );
            let patched = xml.replace("<v>#N/A</v>", "<v>#N/A!</v>");
            data = patched.into_bytes();
        }

        writer
            .start_file(name, options)
            .expect("start zip file");
        writer.write_all(&data).expect("write zip file bytes");
    }
    let patched_bytes = writer.finish().expect("finish zip").into_inner();

    let loaded = read_workbook_model_from_bytes(&patched_bytes).expect("read workbook model");
    let sheet = &loaded.sheets[0];
    assert_eq!(
        sheet.value_a1("A1").expect("A1"),
        CellValue::Error(ErrorValue::NA)
    );
}
