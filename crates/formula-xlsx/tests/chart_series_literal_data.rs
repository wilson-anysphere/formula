use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_series_literal_categories_and_values() {
    // Minimal chartSpace containing a single series with literal category/value points.
    // Some Excel charts embed data directly in the chart XML (`c:strLit` / `c:numLit`)
    // instead of referencing worksheet cells via `c:strRef` / `c:numRef`.
    let xml = r#"
        <c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:cat>
                    <c:strLit>
                      <c:ptCount val="2"/>
                      <c:pt idx="0"><c:v>Alpha</c:v></c:pt>
                      <c:pt idx="1"><c:v>Beta</c:v></c:pt>
                    </c:strLit>
                  </c:cat>
                  <c:val>
                    <c:numLit>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="2"/>
                      <c:pt idx="0"><c:v>1</c:v></c:pt>
                      <c:pt idx="1"><c:v>2.5</c:v></c:pt>
                    </c:numLit>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "in-memory.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);

    let ser = &model.series[0];

    let categories = ser.categories.as_ref().expect("categories present");
    assert!(categories.formula.is_none());
    assert_eq!(
        categories.cache.as_ref().expect("category cache"),
        &vec!["Alpha".to_string(), "Beta".to_string()]
    );
    assert_eq!(categories.literal, categories.cache);

    let values = ser.values.as_ref().expect("values present");
    assert!(values.formula.is_none());
    assert_eq!(values.cache.as_ref().expect("value cache"), &vec![1.0, 2.5]);
    assert_eq!(values.format_code.as_deref(), Some("General"));
    assert_eq!(values.literal, values.cache);
}

