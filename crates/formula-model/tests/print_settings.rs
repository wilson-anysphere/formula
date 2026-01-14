use formula_model::{CellRef, ManualPageBreaks, Orientation, PageSetup, Range, Scaling, Workbook};

#[test]
fn default_print_settings_exist_for_sheet() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();

    let settings = wb.sheet_print_settings(sheet_id);
    assert_eq!(settings.sheet_name, "Sheet1");
    assert!(settings.print_area.is_none());
    assert!(settings.print_titles.is_none());
    assert_eq!(settings.page_setup, PageSetup::default());
    assert_eq!(settings.manual_page_breaks, ManualPageBreaks::default());
}

#[test]
fn setting_and_clearing_print_area_is_deterministic() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();

    let area1 = vec![Range::new(CellRef::new(0, 0), CellRef::new(1, 1))];
    let area2 = vec![Range::new(CellRef::new(2, 2), CellRef::new(3, 3))];

    // Set out of order; storage should still be ordered by workbook sheet order.
    assert!(wb.set_sheet_print_area(sheet2, Some(area2.clone())));
    assert!(wb.set_sheet_print_area(sheet1, Some(area1.clone())));

    assert_eq!(wb.print_settings.sheets.len(), 2);
    assert_eq!(wb.print_settings.sheets[0].sheet_name, "Sheet1");
    assert_eq!(wb.print_settings.sheets[0].print_area, Some(area1));
    assert_eq!(wb.print_settings.sheets[1].sheet_name, "Sheet2");
    assert_eq!(wb.print_settings.sheets[1].print_area, Some(area2));

    // Clearing should remove the stored overrides entirely.
    assert!(wb.set_sheet_print_area(sheet1, None));
    assert!(wb.set_sheet_print_area(sheet2, None));
    assert!(wb.print_settings.sheets.is_empty());
}

#[test]
fn print_settings_roundtrip_through_serde() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();

    let print_area = vec![Range::from_a1("B2:C3").unwrap()];
    assert!(wb.set_sheet_print_area(sheet_id, Some(print_area)));

    assert!(wb.set_sheet_page_setup(
        sheet_id,
        PageSetup {
            orientation: Orientation::Landscape,
            paper_size: formula_model::PaperSize::A4,
            margins: formula_model::PageMargins::default(),
            scaling: Scaling::FitTo {
                width: 1,
                height: 2
            },
        },
    ));

    let mut breaks = ManualPageBreaks::default();
    breaks.row_breaks_after.insert(10);
    breaks.col_breaks_after.insert(3);
    assert!(wb.set_manual_page_breaks(sheet_id, breaks.clone()));

    let json = serde_json::to_string(&wb).unwrap();
    let deserialized: Workbook = serde_json::from_str(&json).unwrap();

    assert_eq!(deserialized.print_settings, wb.print_settings);
    assert_eq!(
        deserialized.sheet_print_settings(sheet_id),
        wb.sheet_print_settings(sheet_id)
    );
}
