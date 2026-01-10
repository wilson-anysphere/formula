use formula_xlsx::print::{
    calculate_pages, CellRange, ManualPageBreaks, PageMargins, PageSetup, PaperSize, Scaling,
};

fn format_pages(pages: &[formula_xlsx::print::Page]) -> String {
    pages
        .iter()
        .enumerate()
        .map(|(idx, p)| {
            format!(
                "{:02}: R{}-{} C{}-{}\n",
                idx + 1,
                p.start_row,
                p.end_row,
                p.start_col,
                p.end_col
            )
        })
        .collect()
}

#[test]
fn calculate_pages_snapshot_letter_no_margins() {
    // Letter = 612x792 points.
    // With 200pt columns => 3 cols/page.
    // With 100pt rows => 7 rows/page.
    let col_widths = vec![200.0; 10];
    let row_heights = vec![100.0; 20];

    let setup = PageSetup {
        paper_size: PaperSize::LETTER,
        margins: PageMargins {
            left: 0.0,
            right: 0.0,
            top: 0.0,
            bottom: 0.0,
            header: 0.0,
            footer: 0.0,
        },
        scaling: Scaling::Percent(100),
        ..PageSetup::default()
    };

    let pages = calculate_pages(
        CellRange {
            start_row: 1,
            end_row: 20,
            start_col: 1,
            end_col: 10,
        },
        &col_widths,
        &row_heights,
        &setup,
        &ManualPageBreaks::default(),
    );

    let snapshot = format_pages(&pages);
    assert_eq!(
        snapshot,
        "\
01: R1-7 C1-3
02: R1-7 C4-6
03: R1-7 C7-9
04: R1-7 C10-10
05: R8-14 C1-3
06: R8-14 C4-6
07: R8-14 C7-9
08: R8-14 C10-10
09: R15-20 C1-3
10: R15-20 C4-6
11: R15-20 C7-9
12: R15-20 C10-10
"
    );
}
