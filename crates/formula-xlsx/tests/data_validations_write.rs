use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::{Cell, CellRef, CellValue, DataValidation, DataValidationKind, Range, Workbook};
use formula_xlsx::{load_from_bytes, load_from_path, XlsxDocument};
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::{ZipArchive, ZipWriter};

fn read_part(bytes: &[u8], part: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut text = String::new();
    archive.by_name(part)?.read_to_string(&mut text)?;
    Ok(text)
}

fn extract_data_validations_subtree(xml: &str) -> Option<String> {
    let start = xml.find("<dataValidations")?;

    if let Some(end_rel) = xml[start..].find("</dataValidations>") {
        let end = start + end_rel + "</dataValidations>".len();
        return Some(xml[start..end].to_string());
    }

    // Self-closing block (rare, but valid).
    if let Some(end_rel) = xml[start..].find("/>") {
        let end = start + end_rel + "/>".len();
        return Some(xml[start..end].to_string());
    }

    None
}

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    // Minimal styles part: only a default xf.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn writes_data_validations_section() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.set_cell(
            CellRef::from_a1("A1")?,
            Cell::new(CellValue::String("Pick".to_string())),
        );

        let rule = DataValidation {
            kind: DataValidationKind::List,
            operator: None,
            // The model convention is to store formulas without a leading '=' but we accept it
            // defensively and strip it when writing.
            formula1: "=\"Yes,No\"".to_string(),
            formula2: None,
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            // Excel shows the in-cell dropdown arrow by default for list validations.
            show_drop_down: true,
            input_message: None,
            error_alert: None,
        };
        sheet.add_data_validation(vec![Range::from_a1("A1")?], rule);
    }

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec()?;

    let sheet_xml = read_part(&bytes, "xl/worksheets/sheet1.xml")?;

    assert!(
        sheet_xml.contains("<dataValidations"),
        "expected `<dataValidations>` section, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<dataValidation") && sheet_xml.contains("type=\"list\""),
        "expected list `<dataValidation>` element, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("sqref=\"A1\""),
        "expected sqref=\"A1\", got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<formula1>\"Yes,No\"</formula1>"),
        "expected literal list formula1, got:\n{sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("showDropDown=\"1\""),
        "expected dropdown arrow to be shown when show_drop_down=true, got:\n{sheet_xml}"
    );

    Ok(())
}

#[test]
fn writes_data_validation_formulas_add_xlfn_prefixes() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");

    let rule = DataValidation {
        kind: DataValidationKind::Custom,
        operator: None,
        // Writer should strip the leading '=' and restore `_xlfn.` prefixes.
        formula1: "=SEQUENCE(1)".to_string(),
        formula2: None,
        allow_blank: true,
        show_input_message: false,
        show_error_message: false,
        show_drop_down: false,
        input_message: None,
        error_alert: None,
    };
    sheet.add_data_validation(vec![Range::from_a1("B2")?], rule);

    let bytes = XlsxDocument::new(workbook).save_to_vec()?;
    let sheet_xml = read_part(&bytes, "xl/worksheets/sheet1.xml")?;

    assert!(
        sheet_xml.contains("type=\"custom\"") && sheet_xml.contains("sqref=\"B2\""),
        "expected custom data validation on B2, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<formula1>_xlfn.SEQUENCE(1)</formula1>"),
        "expected `_xlfn.`-prefixed formula, got:\n{sheet_xml}"
    );

    Ok(())
}

#[test]
fn writes_list_data_validation_hides_drop_down_arrow_when_disabled(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");

    let rule = DataValidation {
        kind: DataValidationKind::List,
        operator: None,
        formula1: "\"Yes,No\"".to_string(),
        formula2: None,
        allow_blank: true,
        show_input_message: false,
        show_error_message: false,
        // UI semantics: false => suppress the in-cell dropdown arrow.
        show_drop_down: false,
        input_message: None,
        error_alert: None,
    };
    sheet.add_data_validation(vec![Range::from_a1("C3")?], rule);

    let bytes = XlsxDocument::new(workbook).save_to_vec()?;
    let sheet_xml = read_part(&bytes, "xl/worksheets/sheet1.xml")?;

    assert!(
        sheet_xml.contains("type=\"list\"") && sheet_xml.contains("sqref=\"C3\""),
        "expected list data validation on C3, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("showDropDown=\"1\""),
        "expected showDropDown=\"1\" when show_drop_down=false, got:\n{sheet_xml}"
    );

    Ok(())
}

fn strip_data_validations(xml: &str) -> Result<String, Box<dyn std::error::Error>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut writer = quick_xml::Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut skip_depth = 0usize;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"dataValidations" => {
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"dataValidations" => {
                // Skip.
            }
            _ => writer.write_event(event.to_owned())?,
        }
        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn write_fixture_without_data_validations(
    fixture_path: &Path,
    out_path: &Path,
) -> Result<(), Box<dyn std::error::Error>> {
    let src_file = std::fs::File::open(fixture_path)?;
    let mut archive = ZipArchive::new(src_file)?;

    let dst_file = std::fs::File::create(out_path)?;
    let mut writer = ZipWriter::new(dst_file);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        // Avoid pre-allocating based on attacker-controlled ZIP metadata.
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;

        if name == "xl/worksheets/sheet1.xml" {
            let xml = std::str::from_utf8(&buf)?;
            let stripped = strip_data_validations(xml)?;
            buf = stripped.into_bytes();
        }

        writer.start_file(name, options)?;
        writer.write_all(&buf)?;
    }

    writer.finish()?;
    Ok(())
}

#[test]
fn clearing_data_validations_removes_data_validations_block() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/data-validation-list.xlsx");

    let mut doc = load_from_path(&fixture)?;
    doc.workbook.sheets[0].data_validations.clear();

    let out_bytes = doc.save_to_vec()?;

    let sheet_xml = read_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    assert!(
        !sheet_xml.contains("dataValidations"),
        "expected `<dataValidations>` removal, got:\n{sheet_xml}"
    );

    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, &out_bytes)?;

    // Build an "expected" workbook that is byte-for-byte the fixture except for the removal of
    // `<dataValidations>` from `xl/worksheets/sheet1.xml`. The output should match that with no
    // additional critical diffs.
    let expected_path = tmpdir.path().join("expected.xlsx");
    write_fixture_without_data_validations(&fixture, &expected_path)?;

    let report = xlsx_diff::diff_workbooks(&expected_path, &out_path)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!("Critical diffs detected after clearing data validations");
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("unexpected critical diffs after removing dataValidations block");
    }

    Ok(())
}

#[test]
fn no_op_roundtrip_preserves_data_validations_subtree() -> Result<(), Box<dyn std::error::Error>> {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/data-validation-list.xlsx");
    let fixture_bytes = std::fs::read(&fixture)?;
    let original_sheet_xml = read_part(&fixture_bytes, "xl/worksheets/sheet1.xml")?;
    let original_dv = extract_data_validations_subtree(&original_sheet_xml)
        .expect("fixture should contain a <dataValidations> block");

    let doc = load_from_path(&fixture)?;
    let out_bytes = doc.save_to_vec()?;
    let out_sheet_xml = read_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    let out_dv = extract_data_validations_subtree(&out_sheet_xml)
        .expect("output should contain a <dataValidations> block");

    assert_eq!(out_dv, original_dv);
    Ok(())
}

#[test]
fn no_op_roundtrip_preserves_data_validations_with_xlfn_formulas(
) -> Result<(), Box<dyn std::error::Error>> {
    // Ensure our semantic-change detector treats `_xlfn.`-prefixed formulas as equivalent to the
    // normalized model representation (which strips `_xlfn.` on read). If it doesn't, we would
    // spuriously rewrite `<dataValidations>` on no-op saves.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="1">
    <dataValidation sqref="A1" type="custom" allowBlank="1" showInputMessage="0" showErrorMessage="0">
      <formula1>=_xlfn.SEQUENCE(1)</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let original_sheet_xml = read_part(&bytes, "xl/worksheets/sheet1.xml")?;
    let original_dv = extract_data_validations_subtree(&original_sheet_xml)
        .expect("fixture should contain a <dataValidations> block");

    let doc = load_from_bytes(&bytes)?;
    let out_bytes = doc.save_to_vec()?;
    let out_sheet_xml = read_part(&out_bytes, "xl/worksheets/sheet1.xml")?;
    let out_dv = extract_data_validations_subtree(&out_sheet_xml)
        .expect("output should contain a <dataValidations> block");

    assert_eq!(out_dv, original_dv);
    Ok(())
}
