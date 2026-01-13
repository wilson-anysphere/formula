use formula_model::charts::{ChartKind, PlotAreaModel, SeriesData};
use formula_xlsx::drawingml::charts::parse_chart_space;

fn wrap_plot_area(inner: &str) -> String {
    format!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      {inner}
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
    )
}

#[test]
fn parses_area_chart_plot_area_model() {
    // Use area3DChart to validate both 2D + 3D map to the same kind/model.
    let xml = wrap_plot_area(
        r#"<c:area3DChart>
  <c:grouping val="standard"/>
  <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
  <c:axId val="1"/>
  <c:axId val="2"/>
</c:area3DChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Area);
    match model.plot_area {
        PlotAreaModel::Area(area) => {
            assert_eq!(area.grouping.as_deref(), Some("standard"));
            assert_eq!(area.ax_ids, vec![1, 2]);
        }
        other => panic!("expected PlotAreaModel::Area, got {other:?}"),
    }
}

#[test]
fn parses_doughnut_chart_plot_area_model() {
    let xml = wrap_plot_area(
        r#"<c:doughnutChart>
  <c:varyColors val="1"/>
  <c:firstSliceAng val="90"/>
  <c:holeSize val="25"/>
  <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
</c:doughnutChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Doughnut);
    match model.plot_area {
        PlotAreaModel::Doughnut(doughnut) => {
            assert_eq!(doughnut.vary_colors, Some(true));
            assert_eq!(doughnut.first_slice_angle, Some(90));
            assert_eq!(doughnut.hole_size, Some(25));
        }
        other => panic!("expected PlotAreaModel::Doughnut, got {other:?}"),
    }
}

#[test]
fn parses_radar_chart_plot_area_model() {
    let xml = wrap_plot_area(
        r#"<c:radarChart>
  <c:radarStyle val="filled"/>
  <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
  <c:axId val="100"/>
  <c:axId val="200"/>
</c:radarChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Radar);
    match model.plot_area {
        PlotAreaModel::Radar(radar) => {
            assert_eq!(radar.radar_style.as_deref(), Some("filled"));
            assert_eq!(radar.ax_ids, vec![100, 200]);
        }
        other => panic!("expected PlotAreaModel::Radar, got {other:?}"),
    }
}

#[test]
fn parses_bubble_chart_plot_area_model_and_bubble_size_series_data() {
    let xml = wrap_plot_area(
        r#"<c:bubbleChart>
  <c:bubbleScale val="150"/>
  <c:showNegBubbles val="1"/>
  <c:sizeRepresents val="area"/>
  <c:ser>
    <c:idx val="0"/>
    <c:order val="0"/>
    <c:xVal>
      <c:numRef>
        <c:f>Sheet1!$A$2:$A$5</c:f>
      </c:numRef>
    </c:xVal>
    <c:yVal>
      <c:numRef>
        <c:f>Sheet1!$B$2:$B$5</c:f>
      </c:numRef>
    </c:yVal>
    <c:bubbleSize>
      <c:numRef>
        <c:f>Sheet1!$C$2:$C$5</c:f>
        <c:numCache>
          <c:formatCode>General</c:formatCode>
          <c:ptCount val="4"/>
          <c:pt idx="0"><c:v>10</c:v></c:pt>
          <c:pt idx="1"><c:v>20</c:v></c:pt>
          <c:pt idx="2"><c:v>30</c:v></c:pt>
          <c:pt idx="3"><c:v>40</c:v></c:pt>
        </c:numCache>
      </c:numRef>
    </c:bubbleSize>
  </c:ser>
  <c:axId val="111"/>
  <c:axId val="222"/>
</c:bubbleChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Bubble);
    match &model.plot_area {
        PlotAreaModel::Bubble(bubble) => {
            assert_eq!(bubble.bubble_scale, Some(150));
            assert_eq!(bubble.show_neg_bubbles, Some(true));
            assert_eq!(bubble.size_represents.as_deref(), Some("area"));
            assert_eq!(bubble.ax_ids, vec![111, 222]);
        }
        other => panic!("expected PlotAreaModel::Bubble, got {other:?}"),
    }

    assert_eq!(model.series.len(), 1);
    let ser = &model.series[0];

    // Verify x/y values parsed, since bubble charts depend on them.
    match ser.x_values.as_ref().expect("x_values") {
        SeriesData::Number(num) => {
            assert_eq!(num.formula.as_deref(), Some("Sheet1!$A$2:$A$5"));
        }
        other => panic!("expected x_values to be numeric, got {other:?}"),
    }
    match ser.y_values.as_ref().expect("y_values") {
        SeriesData::Number(num) => {
            assert_eq!(num.formula.as_deref(), Some("Sheet1!$B$2:$B$5"));
        }
        other => panic!("expected y_values to be numeric, got {other:?}"),
    }

    let bubble_size = ser.bubble_size.as_ref().expect("bubble_size");
    assert_eq!(bubble_size.formula.as_deref(), Some("Sheet1!$C$2:$C$5"));
    assert_eq!(bubble_size.format_code.as_deref(), Some("General"));
    assert_eq!(bubble_size.cache.as_deref(), Some(&[10.0, 20.0, 30.0, 40.0][..]));
}

#[test]
fn parses_stock_chart_plot_area_model() {
    let xml = wrap_plot_area(
        r#"<c:stockChart>
  <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
  <c:axId val="10"/>
  <c:axId val="20"/>
  <c:axId val="30"/>
</c:stockChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Stock);
    match model.plot_area {
        PlotAreaModel::Stock(stock) => {
            assert_eq!(stock.ax_ids, vec![10, 20, 30]);
        }
        other => panic!("expected PlotAreaModel::Stock, got {other:?}"),
    }
}

#[test]
fn parses_surface_chart_plot_area_model() {
    // Use surface3DChart to validate both 2D + 3D map to the same kind/model.
    let xml = wrap_plot_area(
        r#"<c:surface3DChart>
  <c:wireframe val="1"/>
  <c:axId val="1"/>
  <c:axId val="2"/>
  <c:axId val="3"/>
</c:surface3DChart>"#,
    );
    let model = parse_chart_space(xml.as_bytes(), "test.xml").expect("parse chartSpace");

    assert_eq!(model.chart_kind, ChartKind::Surface);
    match model.plot_area {
        PlotAreaModel::Surface(surface) => {
            assert_eq!(surface.wireframe, Some(true));
            assert_eq!(surface.ax_ids, vec![1, 2, 3]);
        }
        other => panic!("expected PlotAreaModel::Surface, got {other:?}"),
    }
}
