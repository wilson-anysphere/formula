use formula_model::charts::AxisKind;
use formula_xlsx::drawingml::charts::parse_chart_space;

fn parse_axes(plot_area_children: &str) -> Vec<formula_model::charts::AxisModel> {
    let xml = format!(
        r#"<chartSpace>
  <chart>
    <plotArea>
      <lineChart />
      {plot_area_children}
    </plotArea>
  </chart>
</chartSpace>"#
    );
    parse_chart_space(xml.as_bytes(), "unit-test-chart.xml")
        .expect("parse chartSpace")
        .axes
}

#[test]
fn parses_date_axis_kind() {
    let axes = parse_axes(r#"<dateAx><axId val="10"/></dateAx>"#);
    assert_eq!(axes.len(), 1);
    assert_eq!(axes[0].id, 10);
    assert_eq!(axes[0].kind, AxisKind::Date);
}

#[test]
fn parses_series_axis_kind() {
    let axes = parse_axes(r#"<serAx><axId val="20"/></serAx>"#);
    assert_eq!(axes.len(), 1);
    assert_eq!(axes[0].id, 20);
    assert_eq!(axes[0].kind, AxisKind::Series);
}

#[test]
fn parses_axis_crossing_tick_marks_and_title() {
    let axes = parse_axes(
        r#"<valAx>
  <axId val="1"/>
  <crossAx val="2"/>
  <crosses val="max"/>
  <crossesAt val="5"/>
  <majorTickMark val="out"/>
  <minorTickMark val="in"/>
  <majorUnit val="10"/>
  <minorUnit val="2"/>
  <title>
    <tx>
      <strRef>
        <f>Sheet1!$A$1</f>
        <strCache>
          <ptCount val="1"/>
          <pt idx="0"><v>My Axis</v></pt>
        </strCache>
      </strRef>
    </tx>
    <txPr>
      <defRPr sz="1400" b="1">
        <latin typeface="Arial"/>
      </defRPr>
    </txPr>
  </title>
</valAx>"#,
    );

    assert_eq!(axes.len(), 1);
    let axis = &axes[0];

    assert_eq!(axis.cross_axis_id, Some(2));
    assert_eq!(axis.crosses.as_deref(), Some("max"));
    assert_eq!(axis.crosses_at, Some(5.0));
    assert_eq!(axis.major_tick_mark.as_deref(), Some("out"));
    assert_eq!(axis.minor_tick_mark.as_deref(), Some("in"));
    assert_eq!(axis.major_unit, Some(10.0));
    assert_eq!(axis.minor_unit, Some(2.0));

    let title = axis.title.as_ref().expect("axis title should be parsed");
    assert_eq!(title.formula.as_deref(), Some("Sheet1!$A$1"));
    assert_eq!(title.rich_text.text.as_str(), "My Axis");
    let style = title.style.as_ref().expect("axis title style should be parsed");
    assert_eq!(style.font_family.as_deref(), Some("Arial"));
    assert_eq!(style.size_100pt, Some(1400));
    assert_eq!(style.bold, Some(true));
}

