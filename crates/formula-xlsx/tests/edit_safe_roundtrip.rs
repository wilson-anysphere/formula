use std::io::{Cursor, Read};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::ZipArchive;

const FIXTURE: &[u8] = include_bytes!("fixtures/rt_simple.xlsx");

#[test]
fn direct_model_edits_write_correct_types_and_formulas() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");

    // Do not touch `doc.xlsx_meta_mut().cell_meta` to ensure writer fallbacks handle stale metadata.
    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(123.0));
    sheet.set_value(
        CellRef::from_a1("B1").unwrap(),
        CellValue::String("World".to_string()),
    );
    sheet.set_formula(CellRef::from_a1("C1").unwrap(), Some("SEQUENCE(2)".to_string()));
    sheet.set_formula(
        CellRef::from_a1("D1").unwrap(),
        Some("FORECAST.ETS(1,2,3)".to_string()),
    );

    let saved = doc.save_to_vec().expect("save");

    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let shared_strings = zip_part(&saved, "xl/sharedStrings.xml");

    let a1 = parse_cell(&sheet_xml, "A1");
    assert_eq!(a1.t.as_deref(), None);
    assert_eq!(a1.v.as_deref(), Some("123"));

    let b1 = parse_cell(&sheet_xml, "B1");
    assert_eq!(b1.t.as_deref(), Some("s"));
    assert_eq!(b1.v.as_deref(), Some("1"));

    let shared = parse_shared_strings(&shared_strings);
    assert_eq!(shared.get(1).map(String::as_str), Some("World"));

    let c1 = parse_cell(&sheet_xml, "C1");
    assert_eq!(c1.f.as_deref(), Some("_xlfn.SEQUENCE(2)"));

    let d1 = parse_cell(&sheet_xml, "D1");
    assert_eq!(d1.f.as_deref(), Some("_xlfn.FORECAST.ETS(1,2,3)"));
}

#[test]
fn editing_api_updates_formula_meta_file_text() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    let sheet_id = doc.workbook.sheets[0].id;

    assert!(doc.set_cell_formula(
        sheet_id,
        CellRef::from_a1("C1").unwrap(),
        Some("=SEQUENCE(3)".to_string()),
    ));
    assert!(doc.set_cell_formula(
        sheet_id,
        CellRef::from_a1("D1").unwrap(),
        Some("=FORECAST.ETS(1,2,3)".to_string()),
    ));

    let saved = doc.save_to_vec().expect("save");
    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let c1 = parse_cell(&sheet_xml, "C1");
    assert_eq!(c1.f.as_deref(), Some("_xlfn.SEQUENCE(3)"));

    let d1 = parse_cell(&sheet_xml, "D1");
    assert_eq!(d1.f.as_deref(), Some("_xlfn.FORECAST.ETS(1,2,3)"));
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[derive(Default)]
struct ParsedCell {
    t: Option<String>,
    v: Option<String>,
    f: Option<String>,
}

fn parse_cell(sheet_xml: &[u8], target_a1: &str) -> ParsedCell {
    let mut reader = Reader::from_reader(sheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_cell = false;
    let mut in_v = false;
    let mut in_f = false;
    let mut out = ParsedCell::default();

    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) if e.name().as_ref() == b"c" => {
                let mut r = None;
                let mut t = None;
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"r" => r = Some(val),
                        b"t" => t = Some(val),
                        _ => {}
                    }
                }
                in_cell = r.as_deref() == Some(target_a1);
                if in_cell {
                    out.t = t;
                }
            }
            Event::End(e) if e.name().as_ref() == b"c" && in_cell => break,

            Event::Start(e) if in_cell && e.name().as_ref() == b"v" => in_v = true,
            Event::End(e) if in_cell && e.name().as_ref() == b"v" => in_v = false,
            Event::Text(e) if in_cell && in_v => out.v = Some(e.unescape().expect("text").into_owned()),

            Event::Start(e) if in_cell && e.name().as_ref() == b"f" => in_f = true,
            Event::End(e) if in_cell && e.name().as_ref() == b"f" => in_f = false,
            Event::Text(e) if in_cell && in_f => out.f = Some(e.unescape().expect("text").into_owned()),
            Event::Empty(e) if in_cell && e.name().as_ref() == b"f" => {
                out.f = Some(String::new());
            }

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn parse_shared_strings(xml: &[u8]) -> Vec<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_si = false;
    let mut in_t = false;
    let mut current = String::new();
    let mut strings = Vec::new();

    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) if e.name().as_ref() == b"si" => {
                in_si = true;
                current.clear();
            }
            Event::End(e) if e.name().as_ref() == b"si" => {
                if in_si {
                    strings.push(current.clone());
                }
                in_si = false;
            }

            Event::Start(e) if in_si && e.name().as_ref() == b"t" => in_t = true,
            Event::End(e) if in_si && e.name().as_ref() == b"t" => in_t = false,
            Event::Text(e) if in_si && in_t => current.push_str(&e.unescape().expect("text")),

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    strings
}
