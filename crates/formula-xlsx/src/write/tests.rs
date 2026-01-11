use super::*;

use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Read;
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
fn writes_spreadsheetml_formula_text_without_leading_equals() {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    let a1 = formula_model::CellRef::from_a1("A1").unwrap();
    sheet.set_formula(a1, Some("1+1".to_string()));

    let mut doc = crate::XlsxDocument::new(workbook);

    // Simulate stale/incorrect `FormulaMeta` coming from a caller: the `<f>` text must
    // never contain a leading '='.
    doc.meta.cell_meta.insert(
        (sheet_id, a1),
        crate::CellMeta {
            formula: Some(crate::FormulaMeta {
                file_text: "=1+1".to_string(),
                ..Default::default()
            }),
            ..Default::default()
        },
    );

    let bytes = write_to_vec(&doc).expect("write doc");
    let formulas = worksheet_formula_texts_from_xlsx(&bytes, "xl/worksheets/sheet1.xml");
    for f in formulas.into_iter().filter(|f| !f.is_empty()) {
        assert!(
            !f.starts_with('='),
            "SpreadsheetML <f> text must not start with '=' (got {f:?})"
        );
    }
}
