use formula_xlsx::XlsxPackage;
use pretty_assertions::assert_eq;

#[test]
fn pivot_ux_graph_includes_connections_between_pivots_slicers_timelines_and_charts() {
    let fixture = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");
    let pkg = XlsxPackage::from_bytes(fixture).expect("read fixture");

    let graph = pkg.pivot_ux_graph().expect("build pivot ux graph");

    assert_eq!(graph.pivot_tables.len(), 1);
    assert_eq!(graph.slicers.len(), 1);
    assert_eq!(graph.timelines.len(), 1);
    assert_eq!(graph.pivot_charts.len(), 1);

    assert_eq!(graph.slicer_to_pivot_tables, vec![vec![0]]);
    assert_eq!(graph.timeline_to_pivot_tables, vec![vec![0]]);
    assert_eq!(graph.pivot_chart_to_pivot_table, vec![Some(0)]);

    assert_eq!(graph.pivot_table_to_slicers, vec![vec![0]]);
    assert_eq!(graph.pivot_table_to_timelines, vec![vec![0]]);
    assert_eq!(graph.pivot_table_to_pivot_charts, vec![vec![0]]);
}

