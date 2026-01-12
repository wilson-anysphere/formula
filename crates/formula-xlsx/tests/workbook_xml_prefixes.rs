use std::io::{Cursor, Read, Write};

use formula_model::calc_settings::CalculationMode;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, CellPatch, RecalcPolicy, WorkbookCellPatches, XlsxPackage};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn build_prefixed_workbook_xlsx() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    // NOTE: SpreadsheetML namespace is *not* the default namespace here. All SpreadsheetML elements
    // are prefixed with `x:` and the workbook uses a non-`r` relationships prefix (`rel:`).
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:calcPr calcMode="manual" calcOnSave="0" fullCalcOnLoad="0" iterative="1" iterateCount="5" iterateDelta="0.5" fullPrecision="0"/>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    // Minimal valid styles.xml.
    let styles = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

    // Keep worksheet XML unprefixed (with a default SpreadsheetML namespace) so the patch pipeline
    // doesn't need to deal with worksheet prefixing in this test.
    let sheet1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", root_rels.as_slice()),
        ("xl/workbook.xml", workbook_xml.as_slice()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_slice()),
        ("xl/styles.xml", styles.as_slice()),
        ("xl/worksheets/sheet1.xml", sheet1.as_slice()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_prefixed_workbook_xlsx_missing_calc_pr() -> Vec<u8> {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    // NOTE: SpreadsheetML namespace is *not* the default namespace here. All SpreadsheetML elements
    // are prefixed with `x:` and the workbook uses a non-`r` relationships prefix (`rel:`).
    //
    // This workbook intentionally omits `<calcPr>` so callers must insert it.
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    // Minimal valid styles.xml.
    let styles = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"#;

    // Keep worksheet XML unprefixed (with a default SpreadsheetML namespace) so the patch pipeline
    // doesn't need to deal with worksheet prefixing in this test.
    let sheet1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", root_rels.as_slice()),
        ("xl/workbook.xml", workbook_xml.as_slice()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_slice()),
        ("xl/styles.xml", styles.as_slice()),
        ("xl/worksheets/sheet1.xml", sheet1.as_slice()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = zip.by_name(name).unwrap();
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

fn assert_workbook_xml_namespace_correct(workbook_xml: &[u8]) {
    let workbook_xml = std::str::from_utf8(workbook_xml).unwrap();
    let doc = roxmltree::Document::parse(workbook_xml).unwrap();
    let root = doc.root_element();
    assert_eq!(root.tag_name().name(), "workbook");
    assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

    let calc_pr = root
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "calcPr")
        .expect("workbook.xml should contain <calcPr>");
    assert_eq!(calc_pr.tag_name().namespace(), Some(SPREADSHEETML_NS));
}

#[test]
fn workbook_xml_prefixes_are_roundtrip_safe() {
    let bytes = build_prefixed_workbook_xlsx();

    // A) calc_settings() reads from a prefixed <x:calcPr>.
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let settings = pkg.calc_settings().unwrap();
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
    assert!(!settings.calculate_before_save);
    assert!(!settings.full_calc_on_load);
    assert!(!settings.full_precision);
    assert!(settings.iterative.enabled);
    assert_eq!(settings.iterative.max_iterations, 5);
    assert_eq!(settings.iterative.max_change, 0.5);

    // B) Formula edit triggers recalc policy; workbook.xml must remain namespace-correct.
    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::new(0, 0),
        CellPatch::set_value_with_formula(CellValue::Number(0.0), "=1+1"),
    );
    pkg.apply_cell_patches_with_recalc_policy(&patches, RecalcPolicy::default())
        .unwrap();
    assert_workbook_xml_namespace_correct(pkg.part("xl/workbook.xml").unwrap());

    // C) load_from_bytes + save_to_vec preserves namespace correctness.
    let doc = load_from_bytes(&bytes).unwrap();
    let saved = doc.save_to_vec().unwrap();
    let saved_workbook_xml = read_zip_part(&saved, "xl/workbook.xml");
    assert_workbook_xml_namespace_correct(&saved_workbook_xml);
}

#[test]
fn workbook_xml_prefixes_calc_settings_inserts_calc_pr_with_prefix() {
    let bytes = build_prefixed_workbook_xlsx_missing_calc_pr();
    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    // With no calcPr present, reading should fall back to defaults.
    let settings = pkg.calc_settings().unwrap();
    assert_eq!(settings, formula_model::calc_settings::CalcSettings::default());

    // Writing should insert a prefixed calcPr tag (SpreadsheetML is prefix-only in this workbook).
    let mut updated = settings.clone();
    updated.calculation_mode = CalculationMode::Manual;
    pkg.set_calc_settings(&updated).unwrap();

    let updated_workbook_xml = pkg.part("xl/workbook.xml").unwrap();
    assert_workbook_xml_namespace_correct(updated_workbook_xml);

    let updated_workbook_xml_str = std::str::from_utf8(updated_workbook_xml).unwrap();
    assert!(
        updated_workbook_xml_str.contains("<x:calcPr"),
        "expected inserted calcPr to use workbook's SpreadsheetML prefix"
    );
}

#[test]
fn workbook_xml_prefixes_recalc_policy_inserts_calc_pr_with_prefix() {
    let bytes = build_prefixed_workbook_xlsx_missing_calc_pr();
    let mut pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    // Trigger recalc policy by changing a formula; this must insert calcPr fullCalcOnLoad="1".
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::new(0, 0),
        CellPatch::set_value_with_formula(CellValue::Number(0.0), "=1+1"),
    );
    pkg.apply_cell_patches_with_recalc_policy(&patches, RecalcPolicy::default())
        .unwrap();

    let workbook_xml = pkg.part("xl/workbook.xml").unwrap();
    assert_workbook_xml_namespace_correct(workbook_xml);
    let workbook_xml_str = std::str::from_utf8(workbook_xml).unwrap();
    assert!(workbook_xml_str.contains("<x:calcPr"));
    assert!(
        workbook_xml_str.contains("fullCalcOnLoad=\"1\""),
        "expected inserted calcPr to request full calc on load"
    );
}
