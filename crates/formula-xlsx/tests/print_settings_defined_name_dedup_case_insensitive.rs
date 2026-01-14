use std::io::Cursor;

use formula_model::{DefinedNameScope, Range, Workbook};

#[test]
fn writer_dedups_print_defined_names_case_insensitively() -> Result<(), Box<dyn std::error::Error>>
{
    let mut workbook = Workbook::new();
    let sheet1 = workbook.add_sheet("Sheet1")?;

    // Some producers may spell built-in defined names with different casing. Ensure the semantic
    // writer does not emit duplicates when `Workbook::print_settings` is present.
    workbook.create_defined_name(
        DefinedNameScope::Sheet(sheet1),
        "_xlnm.print_area",
        "Sheet1!$A$1:$A$1",
        None,
        false,
        Some(0),
    )?;

    workbook.set_sheet_print_area(sheet1, Some(vec![Range::from_a1("A1")?]));

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let mut zip = zip::ZipArchive::new(Cursor::new(bytes))?;
    let mut workbook_file = zip.by_name("xl/workbook.xml")?;
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut workbook_file, &mut xml)?;

    let doc = roxmltree::Document::parse(&xml)?;
    let count = doc
        .descendants()
        .filter(|n| {
            n.is_element()
                && n.tag_name().name() == "definedName"
                && n.attribute("name")
                    .is_some_and(|v| v.eq_ignore_ascii_case("_xlnm.Print_Area"))
        })
        .count();
    assert_eq!(count, 1, "expected exactly one Print_Area defined name, got {count}\n{xml}");

    Ok(())
}

