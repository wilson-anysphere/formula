use base64::{engine::general_purpose::STANDARD, Engine as _};
use formula_model::{ColRange, Orientation, PageMargins, Range, RowRange, Scaling};
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use std::io::{Cursor, Write};

fn load_fixture_xlsx() -> Vec<u8> {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/print-settings.xlsx.base64");
    let data = std::fs::read_to_string(&fixture_path).expect("fixture base64 should be readable");
    let cleaned: String = data.lines().map(str::trim).collect();
    STANDARD
        .decode(cleaned.as_bytes())
        .expect("fixture base64 should decode")
}

fn build_xlsx_without_defined_names() -> Vec<u8> {
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

    // Exercise `<pageSetUpPr fitToPage="1"/>` even when the `<pageSetup>` element does not
    // explicitly set `fitToWidth`/`fitToHeight`.
    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr>
    <pageSetUpPr fitToPage="1"/>
  </sheetPr>
  <sheetData/>
  <pageMargins left="1.1" right="1.2" top="1.3" bottom="1.4" header="0.5" footer="0.6"/>
  <pageSetup paperSize="9" orientation="landscape"/>
  <rowBreaks count="1" manualBreakCount="1">
    <brk id="1" max="16383" man="1"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="1">
    <brk id="3" max="1048575" man="1"/>
  </colBreaks>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_xlsx_fit_to_page_disabled_but_fit_dimensions_present() -> Vec<u8> {
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

    // Regression: `pageSetUpPr/@fitToPage` controls whether `pageSetup/@fitToWidth` and
    // `pageSetup/@fitToHeight` are active. When `fitToPage="0"`, Excel uses `pageSetup/@scale`
    // percent scaling instead (even if fit dimensions are present).
    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr>
    <pageSetUpPr fitToPage="0"/>
  </sheetPr>
  <sheetData/>
  <pageSetup paperSize="9" orientation="portrait" scale="77" fitToWidth="2" fitToHeight="3"/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn print_settings_imports_page_setup_breaks_print_area_and_titles_into_workbook_model() {
    let bytes = load_fixture_xlsx();

    // Lightweight reader.
    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.print_settings.sheets.len(), 1);
    let sheet = &workbook.print_settings.sheets[0];
    assert_eq!(sheet.sheet_name, "Sheet1");

    assert_eq!(
        sheet.print_area.as_deref(),
        Some(&[Range::from_a1("A1:D10").unwrap()][..])
    );
    assert_eq!(
        sheet.print_titles,
        Some(formula_model::PrintTitles {
            repeat_rows: Some(RowRange { start: 0, end: 0 }),
            repeat_cols: Some(ColRange { start: 0, end: 1 }),
        })
    );

    assert_eq!(sheet.page_setup.orientation, Orientation::Landscape);
    assert_eq!(sheet.page_setup.paper_size.code, 9);
    assert_eq!(sheet.page_setup.margins, PageMargins::default());
    assert_eq!(
        sheet.page_setup.scaling,
        Scaling::FitTo {
            width: 1,
            height: 0
        }
    );

    // Manual breaks are stored in XLSX as 1-based row/col numbers after which the break occurs.
    // The model stores these indices as 0-based.
    assert!(sheet.manual_page_breaks.row_breaks_after.contains(&4));
    assert!(sheet.manual_page_breaks.col_breaks_after.contains(&1));

    // Full-fidelity loader should populate the same model field.
    let doc = load_from_bytes(&bytes).expect("load xlsx document");
    assert_eq!(
        doc.workbook.print_settings.sheets,
        workbook.print_settings.sheets
    );
}

#[test]
fn print_settings_imports_worksheet_page_setup_even_without_defined_names() {
    let bytes = build_xlsx_without_defined_names();

    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.print_settings.sheets.len(), 1);
    let sheet = &workbook.print_settings.sheets[0];
    assert_eq!(sheet.sheet_name, "Sheet1");

    assert_eq!(sheet.print_area, None);
    assert_eq!(sheet.print_titles, None);
    assert_eq!(sheet.page_setup.orientation, Orientation::Landscape);
    assert_eq!(sheet.page_setup.paper_size.code, 9);
    assert_eq!(
        sheet.page_setup.margins,
        PageMargins {
            left: 1.1,
            right: 1.2,
            top: 1.3,
            bottom: 1.4,
            header: 0.5,
            footer: 0.6,
        }
    );
    assert_eq!(
        sheet.page_setup.scaling,
        Scaling::FitTo {
            width: 0,
            height: 0
        }
    );

    // brk/@id values are 1-based; model stores 0-based indices.
    assert!(sheet.manual_page_breaks.row_breaks_after.contains(&0));
    assert!(sheet.manual_page_breaks.col_breaks_after.contains(&2));
}

#[test]
fn print_settings_scale_wins_when_fit_to_page_is_explicitly_disabled() {
    let bytes = build_xlsx_fit_to_page_disabled_but_fit_dimensions_present();

    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.print_settings.sheets.len(), 1);
    let sheet = &workbook.print_settings.sheets[0];

    assert_eq!(sheet.sheet_name, "Sheet1");
    assert_eq!(sheet.page_setup.paper_size.code, 9);
    assert_eq!(sheet.page_setup.orientation, Orientation::Portrait);
    assert_eq!(sheet.page_setup.scaling, Scaling::Percent(77));
}
