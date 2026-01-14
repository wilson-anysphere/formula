use formula_model::charts::{ChartKind, LegendPosition};
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_manual_layout_for_plot_area() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:layout>
        <c:manualLayout>
          <c:xMode val="edge"/>
          <c:yMode val="edge"/>
          <c:wMode val="factor"/>
          <c:hMode val="factor"/>
          <c:x val="0.25"/>
          <c:y val="0.5"/>
          <c:w val="0.75"/>
          <c:h val="0.25"/>
        </c:manualLayout>
      </c:layout>
      <c:barChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let model =
        parse_chart_space(xml.as_bytes(), "manual-layout-plot-area.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);

    let layout = model
        .plot_area_layout
        .as_ref()
        .expect("plot area manual layout present");
    assert_eq!(layout.x, Some(0.25));
    assert_eq!(layout.y, Some(0.5));
    assert_eq!(layout.w, Some(0.75));
    assert_eq!(layout.h, Some(0.25));
    assert_eq!(layout.x_mode.as_deref(), Some("edge"));
    assert_eq!(layout.y_mode.as_deref(), Some("edge"));
    assert_eq!(layout.w_mode.as_deref(), Some("factor"));
    assert_eq!(layout.h_mode.as_deref(), Some("factor"));
}

#[test]
fn parses_manual_layout_for_legend() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart/>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:layout>
        <c:manualLayout>
          <c:x val="0.125"/>
          <c:y val="0.25"/>
          <c:w val="0.5"/>
          <c:h val="0.75"/>
        </c:manualLayout>
      </c:layout>
    </c:legend>
  </c:chart>
</c:chartSpace>
"#;

    let model =
        parse_chart_space(xml.as_bytes(), "manual-layout-legend.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);

    let legend = model.legend.as_ref().expect("legend present");
    assert_eq!(legend.position, LegendPosition::Right);

    let layout = legend
        .layout
        .as_ref()
        .expect("legend manual layout present");
    assert_eq!(layout.x, Some(0.125));
    assert_eq!(layout.y, Some(0.25));
    assert_eq!(layout.w, Some(0.5));
    assert_eq!(layout.h, Some(0.75));
}

#[test]
fn parses_manual_layout_for_title() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:title>
      <c:layout>
        <c:manualLayout>
          <c:xMode val="edge"/>
          <c:yMode val="edge"/>
          <c:x val="0.1"/>
          <c:y val="0.2"/>
        </c:manualLayout>
      </c:layout>
      <c:tx>
        <c:v>My title</c:v>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let model =
        parse_chart_space(xml.as_bytes(), "manual-layout-title.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);

    let title = model.title.as_ref().expect("title present");
    assert_eq!(title.rich_text.plain_text(), "My title");

    let layout = title.layout.as_ref().expect("title manual layout present");
    assert_eq!(layout.x_mode.as_deref(), Some("edge"));
    assert_eq!(layout.y_mode.as_deref(), Some("edge"));
    assert_eq!(layout.x, Some(0.1));
    assert_eq!(layout.y, Some(0.2));
}
