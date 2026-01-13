use std::collections::BTreeSet;
use std::io::Cursor;

use formula_model::{
    ManualPageBreaks, PageMargins, PageSetup, PaperSize, PrintTitles, Range, RowRange, Scaling,
    SheetPrintSettings, Workbook,
};
use formula_xlsx::print::{
    read_workbook_print_settings, CellRange, ColRange, Orientation, PageSetup as XlsxPageSetup,
    PaperSize as XlsxPaperSize, PrintTitles as XlsxPrintTitles, RowRange as XlsxRowRange,
    Scaling as XlsxScaling, SheetPrintSettings as XlsxSheetPrintSettings,
    WorkbookPrintSettings as XlsxWorkbookPrintSettings,
};

#[test]
fn exports_print_settings_from_model() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1")?;

    let mut sheet_settings = SheetPrintSettings::new("Sheet1");
    sheet_settings.print_area = Some(vec![Range::from_a1("A1:B2")?]);
    sheet_settings.print_titles = Some(PrintTitles {
        repeat_rows: Some(RowRange { start: 0, end: 0 }),
        repeat_cols: Some(formula_model::ColRange { start: 0, end: 1 }),
    });
    sheet_settings.page_setup = PageSetup {
        orientation: formula_model::Orientation::Landscape,
        paper_size: PaperSize::A4,
        margins: PageMargins {
            left: 0.25,
            right: 0.5,
            top: 0.75,
            bottom: 1.0,
            header: 0.3,
            footer: 0.4,
        },
        scaling: Scaling::FitTo {
            width: 1,
            height: 0,
        },
    };
    sheet_settings.manual_page_breaks = ManualPageBreaks {
        row_breaks_after: BTreeSet::from([4]),
        col_breaks_after: BTreeSet::new(),
    };
    workbook.print_settings.sheets = vec![sheet_settings];

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let read = read_workbook_print_settings(&bytes)?;

    let expected = XlsxWorkbookPrintSettings {
        sheets: vec![XlsxSheetPrintSettings {
            sheet_name: "Sheet1".to_string(),
            print_area: Some(vec![CellRange {
                start_row: 1,
                end_row: 2,
                start_col: 1,
                end_col: 2,
            }]),
            print_titles: Some(XlsxPrintTitles {
                repeat_rows: Some(XlsxRowRange { start: 1, end: 1 }),
                repeat_cols: Some(ColRange { start: 1, end: 2 }),
            }),
            page_setup: XlsxPageSetup {
                orientation: Orientation::Landscape,
                paper_size: XlsxPaperSize { code: 9 },
                margins: formula_xlsx::print::PageMargins {
                    left: 0.25,
                    right: 0.5,
                    top: 0.75,
                    bottom: 1.0,
                    header: 0.3,
                    footer: 0.4,
                },
                scaling: XlsxScaling::FitTo {
                    width: 1,
                    height: 0,
                },
            },
            manual_page_breaks: formula_xlsx::print::ManualPageBreaks {
                row_breaks_after: BTreeSet::from([5]),
                col_breaks_after: BTreeSet::new(),
            },
        }],
    };

    assert_eq!(read, expected);

    Ok(())
}
