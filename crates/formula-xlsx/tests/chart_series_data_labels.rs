use formula_model::charts::{DataLabelsModel, NumberFormatModel};
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_series_data_label_settings() {
    // Minimal chartSpace with a single series containing `c:dLbls`.
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:dLbls>
                    <c:showVal/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="1"/>
                    <c:dLblPos val="outEnd"/>
                    <c:numFmt formatCode="0.00" sourceLinked="0"/>
                  </c:dLbls>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);

    let series = &model.series[0];
    assert_eq!(
        series.data_labels,
        Some(DataLabelsModel {
            show_val: Some(true),
            show_cat_name: Some(false),
            show_ser_name: Some(true),
            position: Some("outEnd".to_string()),
            num_fmt: Some(NumberFormatModel {
                format_code: "0.00".to_string(),
                source_linked: Some(false),
            }),
        })
    );
}
