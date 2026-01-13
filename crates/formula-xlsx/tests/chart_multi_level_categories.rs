use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_multi_level_categories_from_multi_lvl_str_lit() {
    let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat>
            <c:multiLvlStrLit>
              <c:lvl>
                <c:strCache>
                  <c:ptCount val="2"/>
                  <c:pt idx="0"><c:v>Group A</c:v></c:pt>
                  <c:pt idx="1"><c:v>Group B</c:v></c:pt>
                </c:strCache>
              </c:lvl>
              <c:lvl>
                <c:strCache>
                  <c:ptCount val="2"/>
                  <c:pt idx="0"><c:v>Sub 1</c:v></c:pt>
                  <c:pt idx="1"><c:v>Sub 2</c:v></c:pt>
                </c:strCache>
              </c:lvl>
            </c:multiLvlStrLit>
          </c:cat>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;

    let model = parse_chart_space(xml.as_bytes(), "chart1.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 1);

    let cats = model.series[0]
        .categories
        .as_ref()
        .expect("series categories present");

    assert!(
        cats.cache.is_none(),
        "flat cache should be empty for multi-level labels"
    );
    assert!(
        cats.formula.is_none(),
        "multiLvlStrLit should not include a formula"
    );

    let multi = cats
        .multi_cache
        .as_ref()
        .expect("multi-level cache populated");
    assert_eq!(multi.len(), 2);
    assert_eq!(multi[0], vec!["Group A".to_string(), "Group B".to_string()]);
    assert_eq!(multi[1], vec!["Sub 1".to_string(), "Sub 2".to_string()]);
}
