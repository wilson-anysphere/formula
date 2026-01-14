use base64::{engine::general_purpose::STANDARD, Engine as _};
use formula_model::{
    CellRef, ManualPageBreaks, PageMargins, PageSetup, PaperSize, PrintTitles, Range, Scaling,
    WorkbookPrintSettings,
};
use formula_xlsx::print::{
    CellRange, Orientation, SheetPrintSettings as XlsxSheetPrintSettings,
    WorkbookPrintSettings as XlsxWorkbookPrintSettings,
};
use std::io::Cursor;

fn load_fixture_xlsx() -> Vec<u8> {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/print-settings.xlsx.base64");
    let data = std::fs::read_to_string(&fixture_path).expect("fixture base64 should be readable");
    let cleaned: String = data.lines().map(str::trim).collect();
    STANDARD
        .decode(cleaned.as_bytes())
        .expect("fixture base64 should decode")
}

fn xlsx_print_to_model_settings(print: &XlsxWorkbookPrintSettings) -> WorkbookPrintSettings {
    let mut out = WorkbookPrintSettings::default();
    for sheet in &print.sheets {
        let mut model = formula_model::SheetPrintSettings::new(sheet.sheet_name.clone());

        model.print_area = sheet.print_area.as_ref().map(|ranges| {
            ranges
                .iter()
                .map(|r| {
                    Range::new(
                        CellRef::new(r.start_row.saturating_sub(1), r.start_col.saturating_sub(1)),
                        CellRef::new(r.end_row.saturating_sub(1), r.end_col.saturating_sub(1)),
                    )
                })
                .collect()
        });

        model.print_titles = sheet.print_titles.map(|t| PrintTitles {
            repeat_rows: t.repeat_rows.map(|r| formula_model::RowRange {
                start: r.start.saturating_sub(1),
                end: r.end.saturating_sub(1),
            }),
            repeat_cols: t.repeat_cols.map(|c| formula_model::ColRange {
                start: c.start.saturating_sub(1),
                end: c.end.saturating_sub(1),
            }),
        });

        model.page_setup = PageSetup {
            orientation: match sheet.page_setup.orientation {
                Orientation::Portrait => formula_model::Orientation::Portrait,
                Orientation::Landscape => formula_model::Orientation::Landscape,
            },
            paper_size: PaperSize {
                code: sheet.page_setup.paper_size.code,
            },
            margins: PageMargins {
                left: sheet.page_setup.margins.left,
                right: sheet.page_setup.margins.right,
                top: sheet.page_setup.margins.top,
                bottom: sheet.page_setup.margins.bottom,
                header: sheet.page_setup.margins.header,
                footer: sheet.page_setup.margins.footer,
            },
            scaling: match sheet.page_setup.scaling {
                formula_xlsx::print::Scaling::Percent(pct) => Scaling::Percent(pct),
                formula_xlsx::print::Scaling::FitTo { width, height } => {
                    Scaling::FitTo { width, height }
                }
            },
        };

        let mut breaks = ManualPageBreaks::default();
        for id in &sheet.manual_page_breaks.row_breaks_after {
            breaks.row_breaks_after.insert(id.saturating_sub(1));
        }
        for id in &sheet.manual_page_breaks.col_breaks_after {
            breaks.col_breaks_after.insert(id.saturating_sub(1));
        }
        model.manual_page_breaks = breaks;

        if !model.is_default() {
            out.sheets.push(model);
        }
    }

    out
}

#[test]
fn print_settings_xlsx_document_writeback_updates_workbook_and_worksheets(
) -> Result<(), Box<dyn std::error::Error>> {
    let original = load_fixture_xlsx();
    let baseline = formula_xlsx::print::read_workbook_print_settings(&original)?;
    assert_eq!(
        baseline.sheets.len(),
        1,
        "expected fixture to contain a single sheet"
    );

    // Build an updated print settings struct (xlsx-print types), then convert to model types.
    let mut expected = baseline.clone();
    let sheet: &mut XlsxSheetPrintSettings = &mut expected.sheets[0];

    // Update print area to B2:C5.
    sheet.print_area = Some(vec![CellRange {
        start_row: 2,
        end_row: 5,
        start_col: 2,
        end_col: 3,
    }]);

    // Update manual page breaks.
    sheet.manual_page_breaks.row_breaks_after.clear();
    sheet.manual_page_breaks.row_breaks_after.insert(3);
    sheet.manual_page_breaks.col_breaks_after.clear();
    sheet.manual_page_breaks.col_breaks_after.insert(4);

    // Keep print titles/page setup as-is (ensures we don't regress those fields while writing).

    let mut doc = formula_xlsx::load_from_bytes(&original)?;
    doc.workbook.print_settings = xlsx_print_to_model_settings(&expected);

    let saved = doc.save_to_vec()?;
    let reread = formula_xlsx::print::read_workbook_print_settings(&saved)?;
    assert_eq!(reread, expected);

    Ok(())
}

#[test]
fn print_settings_xlsx_document_noop_roundtrip_has_no_diffs(
) -> Result<(), Box<dyn std::error::Error>> {
    let original = load_fixture_xlsx();
    let doc = formula_xlsx::load_from_bytes(&original)?;
    let saved = doc.save_to_vec()?;

    let tmpdir = tempfile::tempdir()?;
    let original_path = tmpdir.path().join("original.xlsx");
    let saved_path = tmpdir.path().join("saved.xlsx");
    std::fs::write(&original_path, &original)?;
    std::fs::write(&saved_path, &saved)?;

    let report = xlsx_diff::diff_workbooks(&original_path, &saved_path)?;
    assert!(
        report.is_empty(),
        "expected no diffs on no-op roundtrip, got:\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<Vec<_>>()
            .join("\n")
    );

    Ok(())
}

#[test]
fn print_settings_xlsx_document_writeback_matches_unicode_sheet_names_case_insensitive_like_excel(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Straße")?;

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let original = buf.into_inner();

    let mut doc = formula_xlsx::load_from_bytes(&original)?;
    let mut settings = formula_model::SheetPrintSettings::new("STRASSE");
    settings.print_area = Some(vec![Range::from_a1("A1")?]);
    doc.workbook.print_settings.sheets = vec![settings];

    let saved = doc.save_to_vec()?;
    let reread = formula_xlsx::print::read_workbook_print_settings(&saved)?;
    assert_eq!(reread.sheets.len(), 1);
    assert_eq!(reread.sheets[0].sheet_name, "Straße");
    assert_eq!(
        reread.sheets[0].print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 1,
                start_col: 1,
                end_col: 1
            }][..]
        )
    );

    Ok(())
}
