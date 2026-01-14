use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_numeric_categories_from_cat_numref() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:cat>
                    <c:numRef>
                      <c:f>Sheet1!$A$2:$A$5</c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="4"/>
                        <c:pt idx="0"><c:v>45123</c:v></c:pt>
                        <c:pt idx="1"><c:v>45124</c:v></c:pt>
                        <c:pt idx="2"><c:v>45125</c:v></c:pt>
                        <c:pt idx="3"><c:v>45126</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:cat>
                </c:ser>
              </c:lineChart>
              <c:catAx>
                <c:axId val="1"/>
                <c:axPos val="b"/>
              </c:catAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:axPos val="l"/>
              </c:valAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);

    let ser = &model.series[0];
    assert!(
        ser.categories.is_none(),
        "numeric categories should not be stringified into `categories`"
    );
    let cats = ser
        .categories_num
        .as_ref()
        .expect("expected numeric categories in categories_num");
    assert_eq!(cats.formula.as_deref(), Some("Sheet1!$A$2:$A$5"));
    assert_eq!(
        cats.cache.as_deref(),
        Some(&[45123.0, 45124.0, 45125.0, 45126.0][..])
    );
    assert_eq!(cats.format_code.as_deref(), Some("General"));

    // Numeric categories on a category (text) axis should emit a warning.
    assert_eq!(model.diagnostics.len(), 1);
    assert_eq!(
        model.diagnostics[0].message,
        "numeric series categories detected, but the category axis is not a date/value axis; rendering may interpret categories as text"
    );
}

#[test]
fn numeric_categories_with_date_axis_emits_no_warning() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:cat>
                    <c:numRef>
                      <c:f>Sheet1!$A$2:$A$5</c:f>
                      <c:numCache>
                        <c:formatCode>General</c:formatCode>
                        <c:ptCount val="2"/>
                        <c:pt idx="0"><c:v>45123</c:v></c:pt>
                        <c:pt idx="1"><c:v>45124</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:cat>
                </c:ser>
              </c:lineChart>
              <c:dateAx>
                <c:axId val="1"/>
                <c:axPos val="b"/>
              </c:dateAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:axPos val="l"/>
              </c:valAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);
    assert!(model.diagnostics.is_empty(), "unexpected diagnostics: {:?}", model.diagnostics);

    let ser = &model.series[0];
    assert!(ser.categories.is_none());
    assert!(ser.categories_num.is_some());
}

#[test]
fn parses_numeric_categories_from_cat_numlit() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:cat>
                    <c:numLit>
                      <c:formatCode>General</c:formatCode>
                      <c:ptCount val="2"/>
                      <c:pt idx="0"><c:v>45123</c:v></c:pt>
                      <c:pt idx="1"><c:v>45124</c:v></c:pt>
                    </c:numLit>
                  </c:cat>
                </c:ser>
              </c:lineChart>
              <c:dateAx>
                <c:axId val="1"/>
                <c:axPos val="b"/>
              </c:dateAx>
              <c:valAx>
                <c:axId val="2"/>
                <c:axPos val="l"/>
              </c:valAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert!(
        model.diagnostics.is_empty(),
        "unexpected diagnostics: {:?}",
        model.diagnostics
    );
    assert_eq!(model.series.len(), 1);

    let ser = &model.series[0];
    assert!(ser.categories.is_none());
    let cats = ser.categories_num.as_ref().expect("expected categories_num");
    assert_eq!(cats.formula, None);
    assert_eq!(cats.format_code.as_deref(), Some("General"));
    assert_eq!(cats.cache.as_deref(), Some(&[45123.0, 45124.0][..]));
    assert_eq!(
        cats.literal.as_deref(),
        Some(&[45123.0, 45124.0][..]),
        "numLit should populate literal values"
    );
}

