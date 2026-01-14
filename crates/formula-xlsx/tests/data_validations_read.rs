use std::io::{Cursor, Write};
use std::path::Path;

use formula_model::{DataValidationErrorStyle, DataValidationKind, DataValidationOperator, Range};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../")
        .join(rel)
}

#[test]
fn reads_list_data_validation_fixture() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read(fixture_path(
        "fixtures/xlsx/metadata/data-validation-list.xlsx",
    ))?;
    let doc = formula_xlsx::load_from_bytes(&bytes)?;

    let sheet = doc
        .workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("Sheet1 should exist");

    assert_eq!(sheet.data_validations.len(), 1);
    let dv = &sheet.data_validations[0];

    assert_eq!(dv.id, 1, "import should allocate stable ids starting at 1");
    assert_eq!(dv.validation.kind, DataValidationKind::List);
    assert_eq!(dv.validation.allow_blank, true);
    assert_eq!(dv.validation.show_input_message, true);
    assert_eq!(dv.validation.show_error_message, true);
    assert_eq!(
        dv.validation.show_drop_down, true,
        "list validations should show the in-cell dropdown arrow by default"
    );
    assert_eq!(dv.validation.formula1, r#""Yes,No""#);
    assert_eq!(dv.ranges, vec![Range::from_a1("A1")?]);

    Ok(())
}

fn build_data_validation_xlsx() -> Vec<u8> {
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
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <dataValidations count="2">
    <dataValidation type="whole" operator="between" allowBlank="0" showInputMessage="1" showErrorMessage="1" showDropDown="1" sqref="A1 B2:C3" promptTitle="Pick a number" prompt="Enter a value between 1 and 10" errorStyle="warning" errorTitle="Nope" error="Out of range">
      <formula1>=1</formula1>
      <formula2>=10</formula2>
    </dataValidation>
    <dataValidation type="custom" allowBlank="1" sqref="D4">
      <formula1>=_xlfn.SEQUENCE(1)</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn reads_synthetic_data_validations_ranges_operator_and_messages(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_data_validation_xlsx();

    // Use the fast reader path here so this test exercises the streaming worksheet parser.
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)?;
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("Sheet1 should exist");

    assert_eq!(sheet.data_validations.len(), 2);

    let first = &sheet.data_validations[0];
    assert_eq!(first.id, 1);
    assert_eq!(
        first.ranges,
        vec![Range::from_a1("A1")?, Range::from_a1("B2:C3")?]
    );
    assert_eq!(first.validation.kind, DataValidationKind::Whole);
    assert_eq!(
        first.validation.operator,
        Some(DataValidationOperator::Between)
    );
    assert_eq!(first.validation.formula1, "1");
    assert_eq!(first.validation.formula2.as_deref(), Some("10"));
    assert_eq!(first.validation.show_drop_down, false);
    assert_eq!(first.validation.show_input_message, true);
    assert_eq!(first.validation.show_error_message, true);
    assert_eq!(
        first
            .validation
            .input_message
            .as_ref()
            .and_then(|m| m.title.as_deref()),
        Some("Pick a number")
    );
    assert_eq!(
        first
            .validation
            .input_message
            .as_ref()
            .and_then(|m| m.body.as_deref()),
        Some("Enter a value between 1 and 10")
    );
    assert_eq!(
        first.validation.error_alert.as_ref().map(|a| a.style),
        Some(DataValidationErrorStyle::Warning)
    );
    assert_eq!(
        first
            .validation
            .error_alert
            .as_ref()
            .and_then(|a| a.title.as_deref()),
        Some("Nope")
    );
    assert_eq!(
        first
            .validation
            .error_alert
            .as_ref()
            .and_then(|a| a.body.as_deref()),
        Some("Out of range")
    );

    let second = &sheet.data_validations[1];
    assert_eq!(second.id, 2);
    assert_eq!(second.ranges, vec![Range::from_a1("D4")?]);
    assert_eq!(second.validation.kind, DataValidationKind::Custom);
    assert_eq!(second.validation.allow_blank, true);
    assert_eq!(
        second.validation.formula1, "_xlfn.SEQUENCE(1)",
        "import should strip a single leading '=' but preserve formula text otherwise"
    );

    Ok(())
}
