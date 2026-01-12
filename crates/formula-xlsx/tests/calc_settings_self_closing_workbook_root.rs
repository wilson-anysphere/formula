use std::io::{Cursor, Write};

use formula_model::calc_settings::CalculationMode;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn set_calc_settings_expands_prefixed_self_closing_workbook_root() {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    let bytes = build_package(&[("xl/workbook.xml", workbook_xml.as_slice())]);

    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let mut settings = pkg.calc_settings().unwrap();
    settings.calculation_mode = CalculationMode::Manual;
    pkg.set_calc_settings(&settings).unwrap();

    let updated = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(updated).unwrap();
    let root = doc.root_element();
    assert_eq!(root.tag_name().name(), "workbook");
    assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));
    assert!(
        updated.contains("<x:calcPr"),
        "expected inserted calcPr to use SpreadsheetML prefix; got:\n{updated}"
    );
    assert!(
        updated.contains("</x:workbook>"),
        "expected expanded root to include a closing tag; got:\n{updated}"
    );
}

#[test]
fn set_calc_settings_expands_default_ns_self_closing_workbook_root() {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    let bytes = build_package(&[("xl/workbook.xml", workbook_xml.as_slice())]);

    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let mut settings = pkg.calc_settings().unwrap();
    settings.calculation_mode = CalculationMode::Manual;
    pkg.set_calc_settings(&settings).unwrap();

    let updated = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    let doc = roxmltree::Document::parse(updated).unwrap();
    let root = doc.root_element();
    assert_eq!(root.tag_name().name(), "workbook");
    assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));
    assert!(
        updated.contains("<calcPr"),
        "expected inserted calcPr to be unprefixed under default SpreadsheetML namespace; got:\n{updated}"
    );
    assert!(
        updated.contains("</workbook>"),
        "expected expanded root to include a closing tag; got:\n{updated}"
    );
    assert!(
        !updated.contains(":calcPr"),
        "must not introduce a prefixed calcPr when SpreadsheetML is the default namespace; got:\n{updated}"
    );
}

