use formula_model::charts::ChartKind;
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_plot_area_chart_inside_mc_alternate_content() {
    // Minimal chartSpace where the chart type node is wrapped in mc:AlternateContent.
    // Real Excel files commonly use this pattern to wrap newer chart content.
    let xml = r#"<c:chartSpace
    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <c:chart>
    <c:plotArea>
      <mc:AlternateContent>
        <mc:Choice>
          <!-- Non-chart elements can appear in Choice while the chart type lives in Fallback. -->
          <c:spPr />
        </mc:Choice>
        <mc:Fallback>
          <c:barChart>
            <c:barDir val="col"/>
            <c:axId val="123"/>
            <c:axId val="456"/>
            <c:ser>
              <c:tx><c:v>Series 1</c:v></c:tx>
              <c:cat>
                <c:strRef>
                  <c:f>Sheet1!$A$2:$A$3</c:f>
                  <c:strCache>
                    <c:ptCount val="2"/>
                    <c:pt idx="0"><c:v>A</c:v></c:pt>
                    <c:pt idx="1"><c:v>B</c:v></c:pt>
                  </c:strCache>
                </c:strRef>
              </c:cat>
              <c:val>
                <c:numRef>
                  <c:f>Sheet1!$B$2:$B$3</c:f>
                  <c:numCache>
                    <c:formatCode>General</c:formatCode>
                    <c:ptCount val="2"/>
                    <c:pt idx="0"><c:v>1</c:v></c:pt>
                    <c:pt idx="1"><c:v>2</c:v></c:pt>
                  </c:numCache>
                </c:numRef>
              </c:val>
            </c:ser>
          </c:barChart>
          <c:catAx>
            <c:axId val="123" />
            <c:axPos val="b" />
            <c:crossAx val="456" />
            <c:crosses val="autoZero" />
          </c:catAx>
          <c:valAx>
            <c:axId val="456" />
            <c:axPos val="l" />
            <c:crossAx val="123" />
            <c:crosses val="autoZero" />
          </c:valAx>
        </mc:Fallback>
      </mc:AlternateContent>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let model =
        parse_chart_space(xml.as_bytes(), "xl/charts/chart1.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);
    assert!(
        model
            .diagnostics
            .iter()
            .any(|d| d.message.contains("AlternateContent")),
        "expected AlternateContent warning diagnostic"
    );

    assert_eq!(model.axes.len(), 2, "expected two axes to be parsed");

    assert_eq!(model.series.len(), 1);
    let ser = &model.series[0];

    let cats = ser.categories.as_ref().expect("categories parsed");
    assert_eq!(cats.formula.as_deref(), Some("Sheet1!$A$2:$A$3"));
    assert_eq!(
        cats.cache.as_ref().expect("category cache present"),
        &vec!["A".to_string(), "B".to_string()]
    );

    let vals = ser.values.as_ref().expect("values parsed");
    assert_eq!(vals.formula.as_deref(), Some("Sheet1!$B$2:$B$3"));
    assert_eq!(
        vals.cache.as_ref().expect("value cache present"),
        &vec![1.0, 2.0]
    );
}
