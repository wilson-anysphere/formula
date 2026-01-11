use std::io::Read;

use formula_model::CellRef;
use formula_xlsx::{load_from_path, read_workbook, write_workbook};
use quick_xml::events::Event;
use quick_xml::Reader;
use tempfile::tempdir;
use zip::ZipArchive;

fn worksheet_formula_texts_from_xlsx(bytes: &[u8], part_name: &str) -> Vec<String> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(part_name).expect("worksheet part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read worksheet xml");

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut formulas = Vec::new();

    let mut in_f = false;
    let mut current = String::new();
    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) if e.name().as_ref() == b"f" => {
                in_f = true;
                current.clear();
            }
            Event::Empty(e) if e.name().as_ref() == b"f" => {
                formulas.push(String::new());
            }
            Event::Text(t) if in_f => {
                current.push_str(&t.unescape().expect("unescape").into_owned());
            }
            Event::End(e) if e.name().as_ref() == b"f" => {
                in_f = false;
                formulas.push(current.clone());
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    formulas
}

#[test]
fn xlsx_import_stores_formulas_without_leading_equals() {
    let wb = read_workbook("tests/fixtures/table.xlsx").expect("read workbook");
    let sheet = wb.sheet_by_name("Sheet1").expect("Sheet1 missing");

    let e1 = CellRef::from_a1("E1").unwrap();
    assert_eq!(sheet.formula(e1), Some("SUM(Table1[Total])"));

    let d2 = CellRef::from_a1("D2").unwrap();
    assert_eq!(sheet.formula(d2), Some("[@Qty]*[@Price]"));
}

#[test]
fn xlsx_write_does_not_emit_leading_equals_in_f() {
    let wb = read_workbook("tests/fixtures/table.xlsx").expect("read workbook");
    let dir = tempdir().unwrap();
    let out_path = dir.path().join("formula-policy.xlsx");
    write_workbook(&wb, &out_path).expect("write workbook");

    let bytes = std::fs::read(&out_path).expect("read written xlsx");
    let formulas = worksheet_formula_texts_from_xlsx(&bytes, "xl/worksheets/sheet1.xml");
    for f in formulas.into_iter().filter(|f| !f.is_empty()) {
        assert!(
            !f.trim_start().starts_with('='),
            "SpreadsheetML <f> text must not start with '=' (got {f:?})"
        );
    }
}

#[test]
fn xlsx_document_write_does_not_emit_leading_equals_in_f() {
    let doc = load_from_path("tests/fixtures/table.xlsx").expect("load xlsx doc");
    let bytes = doc.save_to_vec().expect("write xlsx doc");
    let formulas = worksheet_formula_texts_from_xlsx(&bytes, "xl/worksheets/sheet1.xml");
    for f in formulas.into_iter().filter(|f| !f.is_empty()) {
        assert!(
            !f.trim_start().starts_with('='),
            "SpreadsheetML <f> text must not start with '=' (got {f:?})"
        );
    }
}
