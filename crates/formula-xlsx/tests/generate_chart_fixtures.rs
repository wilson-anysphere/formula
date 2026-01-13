use std::path::{Path, PathBuf};

use rust_xlsxwriter::{Chart, ChartType, Workbook};

fn write_chart_fixture(path: &Path, chart_type: ChartType) {
    // Avoid overwriting fixtures that may have been manually curated (or exported
    // from Excel) and have committed golden images / models.
    if path.exists() {
        return;
    }
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Category").unwrap();
    worksheet.write_string(0, 1, "Value").unwrap();

    let categories = ["A", "B", "C", "D"];
    let values = [2.0, 4.0, 3.0, 5.0];

    for (i, (cat, val)) in categories.iter().zip(values).enumerate() {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, *cat).unwrap();
        worksheet.write_number(row, 1, val).unwrap();
    }

    let mut chart = Chart::new(chart_type);
    chart.title().set_name("Example Chart");

    let series = chart.add_series();
    series
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    worksheet.insert_chart(1, 3, &chart).unwrap();

    workbook.save(path).unwrap();
}

fn write_combo_chart_fixture(path: &Path) {
    if path.exists() {
        return;
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Category").unwrap();
    worksheet.write_string(0, 1, "Bar").unwrap();
    worksheet.write_string(0, 2, "Line").unwrap();

    let categories = ["A", "B", "C", "D"];
    let bar_values = [2.0, 4.0, 3.0, 5.0];
    let line_values = [1.0, 2.0, 2.5, 4.0];

    for (i, ((cat, bar), line)) in categories
        .iter()
        .zip(bar_values)
        .zip(line_values)
        .enumerate()
    {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, *cat).unwrap();
        worksheet.write_number(row, 1, bar).unwrap();
        worksheet.write_number(row, 2, line).unwrap();
    }

    let mut chart = Chart::new(ChartType::Column);
    chart.title().set_name("Combo Bar + Line");
    let bar_series = chart.add_series();
    bar_series
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    let mut line_chart = Chart::new(ChartType::Line);
    let line_series = line_chart.add_series();
    line_series
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$C$2:$C$5");

    chart.combine(&line_chart);

    worksheet.insert_chart(1, 4, &chart).unwrap();

    workbook.save(path).unwrap();
}

/// Generates small chart fixtures under `fixtures/charts/xlsx/`.
///
/// This test is ignored by default because it writes files to the repository.
#[test]
#[ignore]
fn generate_chart_fixtures() {
    let root: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/charts/xlsx");
    std::fs::create_dir_all(&root).unwrap();

    write_chart_fixture(&root.join("bar.xlsx"), ChartType::Column);
    write_chart_fixture(&root.join("line.xlsx"), ChartType::Line);
    write_chart_fixture(&root.join("pie.xlsx"), ChartType::Pie);
    write_chart_fixture(&root.join("scatter.xlsx"), ChartType::Scatter);

    write_chart_fixture(&root.join("area.xlsx"), ChartType::Area);
    write_chart_fixture(&root.join("doughnut.xlsx"), ChartType::Doughnut);
    write_combo_chart_fixture(&root.join("combo-bar-line.xlsx"));

    // Additional priority chart types from docs/17-charts.md that are useful to
    // keep around as parsing regressions, even if we don't fully model them yet.
    write_chart_fixture(&root.join("bar-horizontal.xlsx"), ChartType::Bar);
    write_chart_fixture(&root.join("radar.xlsx"), ChartType::Radar);
}
