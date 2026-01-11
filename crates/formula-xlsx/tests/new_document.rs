use formula_model::{Cell, CellRef, CellValue, Font, Style, Workbook};
use formula_xlsx::{load_from_bytes, XlsxDocument};
use quick_xml::events::Event;
use quick_xml::Reader;

#[test]
fn new_document_saves_and_loads_with_multiple_sheets() {
    let mut workbook = Workbook::new();
    let sheet1 = workbook.add_sheet("First").unwrap();
    let sheet2 = workbook.add_sheet("Second").unwrap();

    let italic_style_id = workbook.intern_style(Style {
        font: Some(Font {
            italic: true,
            ..Default::default()
        }),
        ..Default::default()
    });

    {
        let ws1 = workbook.sheet_mut(sheet1).unwrap();
        ws1.set_cell(CellRef::from_a1("A1").unwrap(), Cell::new(CellValue::Number(1.0)));
        ws1.set_cell(
            CellRef::from_a1("B1").unwrap(),
            Cell::new(CellValue::String("Hello".to_string())),
        );
    }

    {
        let ws2 = workbook.sheet_mut(sheet2).unwrap();
        let mut cell = Cell::new(CellValue::Boolean(true));
        cell.style_id = italic_style_id;
        ws2.set_cell(CellRef::from_a1("C3").unwrap(), cell);
    }

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("write xlsx");

    let loaded = load_from_bytes(&bytes).expect("read xlsx");

    assert_eq!(loaded.workbook.sheets.len(), 2);
    assert_eq!(loaded.workbook.sheets[0].name, "First");
    assert_eq!(loaded.workbook.sheets[1].name, "Second");

    let ws1 = loaded.workbook.sheet_by_name("First").unwrap();
    assert_eq!(ws1.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        ws1.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );

    let ws2 = loaded.workbook.sheet_by_name("Second").unwrap();
    assert_eq!(ws2.value_a1("C3").unwrap(), CellValue::Boolean(true));
    let styled_cell = ws2.cell_a1("C3").unwrap().unwrap();
    let style = loaded
        .workbook
        .styles
        .get(styled_cell.style_id)
        .expect("style id valid");
    assert!(style.font.as_ref().is_some_and(|font| font.italic));

    let rels = std::str::from_utf8(
        loaded
            .parts()
            .get("xl/_rels/workbook.xml.rels")
            .expect("workbook rels part present"),
    )
    .unwrap();
    assert!(rels.contains("worksheets/sheet1.xml"));
    assert!(rels.contains("worksheets/sheet2.xml"));

    let content_types = std::str::from_utf8(
        loaded
            .parts()
            .get("[Content_Types].xml")
            .expect("content types part present"),
    )
    .unwrap();
    assert!(content_types.contains("/xl/worksheets/sheet1.xml"));
    assert!(content_types.contains("/xl/worksheets/sheet2.xml"));

    let styles = loaded
        .parts()
        .get("xl/styles.xml")
        .expect("styles part present");
    assert!(
        cell_xfs_count(styles) >= 2,
        "styles.xml should contain at least one xf for the workbook default and one for our style"
    );
}

fn cell_xfs_count(xml: &[u8]) -> u32 {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"cellXfs" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"count" {
                        return attr
                            .unescape_value()
                            .expect("count")
                            .parse()
                            .unwrap_or(0);
                    }
                }
                return 0;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    0
}
