use formula_model::charts::{ComboChartEntry, PlotAreaModel};
use formula_xlsx::drawingml::charts::parse_chart_space;
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

#[test]
fn parses_combo_chart_generated_by_rust_xlsxwriter() {
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

    let mut chart = Chart::new(XlsxChartType::Column);
    chart.title().set_name("Combo Bar + Line");
    chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    let mut line_chart = Chart::new(XlsxChartType::Line);
    line_chart
        .add_series()
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$C$2:$C$5");

    chart.combine(&line_chart);
    worksheet.insert_chart(1, 4, &chart).unwrap();

    let bytes = workbook.save_to_buffer().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let chart_xml = pkg
        .part("xl/charts/chart1.xml")
        .expect("workbook contains xl/charts/chart1.xml");

    let model = parse_chart_space(chart_xml, "xl/charts/chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 2);

    let PlotAreaModel::Combo(combo) = model.plot_area else {
        panic!("expected combo plot area, got {:?}", model.plot_area);
    };
    assert_eq!(combo.charts.len(), 2);

    match &combo.charts[0] {
        ComboChartEntry::Bar { series, .. } => {
            assert_eq!(series.start, 0);
            assert_eq!(series.end, 1);
        }
        other => panic!("expected first subplot to be bar, got {other:?}"),
    }
    match &combo.charts[1] {
        ComboChartEntry::Line { series, .. } => {
            assert_eq!(series.start, 1);
            assert_eq!(series.end, 2);
        }
        other => panic!("expected second subplot to be line, got {other:?}"),
    }

    assert_eq!(model.series[0].plot_index, Some(0));
    assert_eq!(model.series[1].plot_index, Some(1));
}

