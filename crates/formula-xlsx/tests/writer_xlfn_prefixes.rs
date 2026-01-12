use std::io::{Cursor, Read};

use formula_model::{CellRef, CellValue};
use zip::ZipArchive;

#[test]
fn write_workbook_to_writer_prefixes_xlfn_functions_in_cells_and_defined_names(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(CellRef::from_a1("A1")?, CellValue::Number(1.0));
    sheet.set_formula(CellRef::from_a1("B1")?, Some("SEQUENCE(2)".to_string()));
    sheet.set_formula(
        CellRef::from_a1("C1")?,
        Some("FORECAST.ETS(1,2,3)".to_string()),
    );

    workbook.create_defined_name(
        formula_model::DefinedNameScope::Workbook,
        "MySeq",
        "SEQUENCE(3)",
        None,
        false,
        None,
    )?;
    workbook.create_defined_name(
        formula_model::DefinedNameScope::Workbook,
        "MyForecast",
        "FORECAST.ETS(1,2,3)",
        None,
        false,
        None,
    )?;

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let mut zip = ZipArchive::new(Cursor::new(bytes))?;

    let mut sheet_xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    assert!(
        sheet_xml.contains("<f>_xlfn.SEQUENCE(2)</f>"),
        "expected cell formula to be stored with _xlfn. prefix, got sheet xml: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<f>_xlfn.FORECAST.ETS(1,2,3)</f>"),
        "expected FORECAST.ETS cell formula to be stored with _xlfn. prefix, got sheet xml: {sheet_xml}"
    );

    let mut workbook_xml = String::new();
    zip.by_name("xl/workbook.xml")?
        .read_to_string(&mut workbook_xml)?;
    assert!(
        workbook_xml.contains(r#"<definedName name="MySeq">_xlfn.SEQUENCE(3)</definedName>"#),
        "expected defined name formula to be stored with _xlfn. prefix, got workbook xml: {workbook_xml}"
    );
    assert!(
        workbook_xml.contains(
            r#"<definedName name="MyForecast">_xlfn.FORECAST.ETS(1,2,3)</definedName>"#
        ),
        "expected FORECAST.ETS defined name formula to be stored with _xlfn. prefix, got workbook xml: {workbook_xml}"
    );

    Ok(())
}
