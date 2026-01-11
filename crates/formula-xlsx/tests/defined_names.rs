use formula_model::{DefinedNameScope, Workbook};

/// Ensure `formula-xlsx` round-trips workbook/sheet scoped defined names through `workbook.xml`.
#[test]
fn defined_names_roundtrip_through_workbook_xml() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet1 = workbook.add_sheet("Sheet1")?;
    let sheet2 = workbook.add_sheet("Sheet2")?;

    workbook.create_defined_name(
        DefinedNameScope::Workbook,
        "MyX",
        "Sheet1!A1",
        None,
        false,
        None,
    )?;
    workbook.create_defined_name(
        DefinedNameScope::Sheet(sheet2),
        "LocalFoo",
        // Leading '=' should be stripped on read/write.
        "=Sheet2!B2",
        Some("example comment".to_string()),
        true,
        None,
    )?;

    let mut buf = std::io::Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let doc = formula_xlsx::load_from_bytes(&bytes)?;
    let roundtripped = doc.workbook;

    let myx = roundtripped
        .get_defined_name(DefinedNameScope::Workbook, "MyX")
        .expect("workbook scoped name missing");
    assert_eq!(myx.refers_to, "Sheet1!A1");

    let sheet2_rt = roundtripped
        .sheet_by_name("Sheet2")
        .expect("Sheet2 missing")
        .id;
    let local = roundtripped
        .get_defined_name(DefinedNameScope::Sheet(sheet2_rt), "LocalFoo")
        .expect("sheet scoped name missing");
    assert_eq!(local.refers_to, "Sheet2!B2");
    assert_eq!(local.comment.as_deref(), Some("example comment"));
    assert!(local.hidden);
    assert_eq!(local.xlsx_local_sheet_id, Some(1));

    // Ensure the initial sheet ids are not used by the loaded workbook, but scope mapping still works.
    assert_ne!(sheet1, sheet2_rt);

    Ok(())
}

