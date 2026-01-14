use std::io::{Cursor, Read};

use formula_model::{
    DataValidation, DataValidationErrorAlert, DataValidationErrorStyle, DataValidationInputMessage,
    DataValidationKind, DataValidationOperator, Range, Workbook,
};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn exports_worksheet_data_validations() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        let validation = DataValidation {
            kind: DataValidationKind::Decimal,
            operator: Some(DataValidationOperator::Between),
            formula1: "SEQUENCE(1)".to_string(),
            formula2: Some("SEQUENCE(2)".to_string()),
            allow_blank: true,
            show_input_message: true,
            show_error_message: true,
            show_drop_down: false,
            input_message: Some(DataValidationInputMessage {
                title: Some("Pick & choose".to_string()),
                body: Some("Enter a value <= 10".to_string()),
            }),
            error_alert: Some(DataValidationErrorAlert {
                style: DataValidationErrorStyle::Warning,
                title: Some("Bad <value>".to_string()),
                body: Some("Must be between 1 & 2".to_string()),
            }),
        };

        sheet.add_data_validation(
            vec![Range::from_a1("A1:A5")?, Range::from_a1("C1:C5")?],
            validation,
        );
    }

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let sheet_xml_bytes = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;

    let data_validations = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "dataValidations")
        .expect("expected <dataValidations> in worksheet xml");
    assert_eq!(data_validations.attribute("count"), Some("1"));

    let dv = data_validations
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "dataValidation")
        .expect("expected <dataValidation> child");

    assert_eq!(dv.attribute("type"), Some("decimal"));
    assert_eq!(dv.attribute("operator"), Some("between"));
    assert_eq!(dv.attribute("allowBlank"), Some("1"));
    assert_eq!(dv.attribute("showInputMessage"), Some("1"));
    assert_eq!(dv.attribute("showErrorMessage"), Some("1"));
    assert_eq!(dv.attribute("showDropDown"), None);

    assert_eq!(dv.attribute("promptTitle"), Some("Pick & choose"));
    assert_eq!(dv.attribute("prompt"), Some("Enter a value <= 10"));

    assert_eq!(dv.attribute("errorStyle"), Some("warning"));
    assert_eq!(dv.attribute("errorTitle"), Some("Bad <value>"));
    assert_eq!(dv.attribute("error"), Some("Must be between 1 & 2"));

    assert_eq!(dv.attribute("sqref"), Some("A1:A5 C1:C5"));

    let formula1 = dv
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "formula1")
        .and_then(|n| n.text())
        .unwrap_or_default();
    let formula2 = dv
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "formula2")
        .and_then(|n| n.text())
        .unwrap_or_default();

    assert_eq!(formula1, "_xlfn.SEQUENCE(1)");
    assert_eq!(formula2, "_xlfn.SEQUENCE(2)");

    Ok(())
}
