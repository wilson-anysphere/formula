use formula_model::{CellRef, DeleteSheetError, Workbook};

#[test]
fn delete_sheet_invalidates_formula_references() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet("Data").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("=Data!A1".to_string()))
        .unwrap();
    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A2", Some("=SUM(Data!A1:B2)".to_string()))
        .unwrap();

    wb.delete_sheet(data).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(calc_sheet.formula(CellRef::new(0, 0)), Some("#REF!"));
    assert_eq!(calc_sheet.formula(CellRef::new(1, 0)), Some("SUM(#REF!)"));
}

#[test]
fn delete_sheet_invalidates_quoted_sheet_references() {
    let mut wb = Workbook::new();
    let data = wb.add_sheet("My Sheet").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("='My Sheet'!A1".to_string()))
        .unwrap();

    wb.delete_sheet(data).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(calc_sheet.formula(CellRef::new(0, 0)), Some("#REF!"));
}

#[test]
fn delete_sheet_does_not_rewrite_external_workbook_references() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();

    wb.sheet_mut(sheet2)
        .unwrap()
        .set_formula_a1("A1", Some("=[Book.xlsx]Sheet1!A1+Sheet1!A1".to_string()))
        .unwrap();

    wb.delete_sheet(sheet1).unwrap();

    let sheet2 = wb.sheet(sheet2).unwrap();
    assert_eq!(
        sheet2.formula(CellRef::new(0, 0)),
        Some("[Book.xlsx]Sheet1!A1+#REF!")
    );
}

#[test]
fn delete_sheet_adjusts_3d_reference_boundaries() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let _sheet2 = wb.add_sheet("Sheet2").unwrap();
    let _sheet3 = wb.add_sheet("Sheet3").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("=SUM(Sheet1:Sheet3!A1)".to_string()))
        .unwrap();

    wb.delete_sheet(sheet1).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(
        calc_sheet.formula(CellRef::new(0, 0)),
        Some("SUM(Sheet2:Sheet3!A1)")
    );
}

#[test]
fn delete_sheet_keeps_3d_reference_when_deleting_inside_span() {
    let mut wb = Workbook::new();
    let _sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();
    let _sheet3 = wb.add_sheet("Sheet3").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("=SUM(Sheet1:Sheet3!A1)".to_string()))
        .unwrap();

    wb.delete_sheet(sheet2).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(
        calc_sheet.formula(CellRef::new(0, 0)),
        Some("SUM(Sheet1:Sheet3!A1)")
    );
}

#[test]
fn delete_sheet_simplifies_3d_reference_to_single_sheet() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let _sheet2 = wb.add_sheet("Sheet2").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("=SUM(Sheet1:Sheet2!A1)".to_string()))
        .unwrap();

    wb.delete_sheet(sheet1).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(
        calc_sheet.formula(CellRef::new(0, 0)),
        Some("SUM(Sheet2!A1)")
    );
}

#[test]
fn delete_sheet_invalidates_3d_reference_that_only_contains_deleted_sheet() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let calc = wb.add_sheet("Calc").unwrap();

    wb.sheet_mut(calc)
        .unwrap()
        .set_formula_a1("A1", Some("=SUM(Sheet1:Sheet1!A1)".to_string()))
        .unwrap();

    wb.delete_sheet(sheet1).unwrap();

    let calc_sheet = wb.sheet(calc).unwrap();
    assert_eq!(calc_sheet.formula(CellRef::new(0, 0)), Some("SUM(#REF!)"));
}

#[test]
fn delete_last_sheet_fails() {
    let mut wb = Workbook::new();
    let only = wb.add_sheet("Only").unwrap();
    assert_eq!(
        wb.delete_sheet(only),
        Err(DeleteSheetError::CannotDeleteLastSheet)
    );
}

#[test]
fn delete_sheet_preserves_xlsx_ids_on_remaining_sheets() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();
    let sheet3 = wb.add_sheet("Sheet3").unwrap();

    {
        let s1 = wb.sheet_mut(sheet1).unwrap();
        s1.xlsx_sheet_id = Some(10);
        s1.xlsx_rel_id = Some("rId10".to_string());
    }
    {
        let s2 = wb.sheet_mut(sheet2).unwrap();
        s2.xlsx_sheet_id = Some(20);
        s2.xlsx_rel_id = Some("rId20".to_string());
    }
    {
        let s3 = wb.sheet_mut(sheet3).unwrap();
        s3.xlsx_sheet_id = Some(30);
        s3.xlsx_rel_id = Some("rId30".to_string());
    }

    wb.delete_sheet(sheet2).unwrap();

    let s1 = wb.sheet(sheet1).unwrap();
    assert_eq!(s1.xlsx_sheet_id, Some(10));
    assert_eq!(s1.xlsx_rel_id.as_deref(), Some("rId10"));

    let s3 = wb.sheet(sheet3).unwrap();
    assert_eq!(s3.xlsx_sheet_id, Some(30));
    assert_eq!(s3.xlsx_rel_id.as_deref(), Some("rId30"));
}
