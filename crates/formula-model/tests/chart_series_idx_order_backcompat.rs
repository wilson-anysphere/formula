use formula_model::charts::ChartModel;

#[test]
fn chart_model_deserializes_missing_series_idx_order_as_position() {
    // Older serialized chart models predate `SeriesModel::{idx, order}`.
    // For backward compatibility, `ChartModel` fills missing idx/order with the
    // series' position, matching Excel's implied defaults.
    let json = r#"
        {
          "chartKind": { "kind": "bar" },
          "title": null,
          "legend": null,
          "plotArea": {
            "kind": "bar",
            "varyColors": null,
            "barDirection": null,
            "grouping": null,
            "gapWidth": null,
            "overlap": null,
            "axIds": []
          },
          "axes": [],
          "series": [{}, {}],
          "diagnostics": []
        }
    "#;

    let model: ChartModel = serde_json::from_str(json).expect("deserialize chart model");
    assert_eq!(model.series.len(), 2);

    assert_eq!(model.series[0].idx, Some(0));
    assert_eq!(model.series[0].order, Some(0));
    assert_eq!(model.series[1].idx, Some(1));
    assert_eq!(model.series[1].order, Some(1));
}
