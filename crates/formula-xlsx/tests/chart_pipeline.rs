use std::collections::HashSet;
use formula_model::charts::{ChartAnchor, ChartType};
use formula_xlsx::XlsxPackage;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};

fn build_workbook(chart_type: XlsxChartType) -> Vec<u8> {
    build_workbook_with_chart(chart_type, true)
}

fn build_workbook_with_chart(chart_type: XlsxChartType, include_chart: bool) -> Vec<u8> {
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

    if include_chart {
        worksheet.insert_chart(1, 3, &chart).unwrap();
    }

    workbook.save_to_buffer().unwrap()
}

fn assert_round_trip_preserves_all_parts(xlsx_bytes: &[u8]) {
    let package = XlsxPackage::from_bytes(xlsx_bytes).unwrap();
    let round_trip_bytes = package.write_to_bytes().unwrap();
    let round_trip = XlsxPackage::from_bytes(&round_trip_bytes).unwrap();

    let original_names: HashSet<_> = package.part_names().map(str::to_string).collect();
    let round_names: HashSet<_> = round_trip.part_names().map(str::to_string).collect();
    assert_eq!(original_names, round_names);

    for name in original_names {
        assert_eq!(package.part(&name), round_trip.part(&name), "mismatch for {name}");
    }
}

#[test]
fn bar_chart_round_trip_and_extract() {
    let bytes = build_workbook(XlsxChartType::Column);
    assert_round_trip_preserves_all_parts(&bytes);

    let package = XlsxPackage::from_bytes(&bytes).unwrap();
    let charts = package.extract_charts().unwrap();
    assert_eq!(charts.len(), 1);

    let chart = &charts[0];
    assert_eq!(chart.sheet_name.as_deref(), Some("Sheet1"));
    assert!(chart.sheet_part.as_deref().unwrap().starts_with("xl/worksheets/"));
    assert!(chart.chart_part.as_deref().unwrap().starts_with("xl/charts/"));
    assert_eq!(chart.chart_type, ChartType::Bar);
    assert_eq!(chart.title.as_deref(), Some("Example Chart"));
    assert_eq!(chart.series.len(), 1);
    assert!(chart.series[0]
        .categories
        .as_deref()
        .unwrap()
        .contains("$A$2:$A$5"));
    assert!(chart.series[0]
        .values
        .as_deref()
        .unwrap()
        .contains("$B$2:$B$5"));

    match chart.anchor {
        ChartAnchor::TwoCell {
            from_col,
            from_row,
            to_col,
            to_row,
            ..
        } => {
            assert!(to_col > from_col);
            assert!(to_row > from_row);
        }
        _ => panic!("expected two cell anchor"),
    };
}

#[test]
fn line_chart_round_trip_and_extract() {
    let bytes = build_workbook(XlsxChartType::Line);
    assert_round_trip_preserves_all_parts(&bytes);

    let package = XlsxPackage::from_bytes(&bytes).unwrap();
    let charts = package.extract_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].chart_type, ChartType::Line);
}

#[test]
fn pie_chart_round_trip_and_extract() {
    let bytes = build_workbook(XlsxChartType::Pie);
    assert_round_trip_preserves_all_parts(&bytes);

    let package = XlsxPackage::from_bytes(&bytes).unwrap();
    let charts = package.extract_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].chart_type, ChartType::Pie);
}

#[test]
fn scatter_chart_round_trip_and_extract() {
    let bytes = build_workbook(XlsxChartType::Scatter);
    assert_round_trip_preserves_all_parts(&bytes);

    let package = XlsxPackage::from_bytes(&bytes).unwrap();
    let charts = package.extract_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].chart_type, ChartType::Scatter);
}

#[test]
fn preserved_drawing_parts_can_be_reapplied_to_regenerated_workbook() {
    let bytes_with_chart = build_workbook_with_chart(XlsxChartType::Column, true);
    let pkg_with_chart = XlsxPackage::from_bytes(&bytes_with_chart).unwrap();
    let preserved = pkg_with_chart.preserve_drawing_parts().unwrap();
    assert!(!preserved.is_empty());

    let bytes_without_chart = build_workbook_with_chart(XlsxChartType::Column, false);
    let mut pkg_without_chart = XlsxPackage::from_bytes(&bytes_without_chart).unwrap();
    assert_eq!(pkg_without_chart.extract_charts().unwrap().len(), 0);

    pkg_without_chart
        .apply_preserved_drawing_parts(&preserved)
        .unwrap();
    let merged_bytes = pkg_without_chart.write_to_bytes().unwrap();

    let merged_pkg = XlsxPackage::from_bytes(&merged_bytes).unwrap();
    let charts = merged_pkg.extract_charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].chart_type, ChartType::Bar);
    assert_eq!(charts[0].title.as_deref(), Some("Example Chart"));
}
