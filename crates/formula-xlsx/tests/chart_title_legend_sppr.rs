use formula_model::charts::FillStyle;
use formula_model::Color;
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_title_and_legend_shape_properties() {
    let xml = r#"
        <c:chartSpace
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <c:chart>
            <c:title>
              <c:tx><c:v>My Title</c:v></c:tx>
              <c:spPr>
                <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                <a:ln w="12700">
                  <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                </a:ln>
              </c:spPr>
            </c:title>
            <c:legend>
              <c:legendPos val="r"/>
              <c:overlay val="0"/>
              <c:spPr>
                <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
                <a:ln w="38100">
                  <a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill>
                </a:ln>
              </c:spPr>
            </c:legend>
            <c:plotArea>
              <c:barChart/>
              <c:catAx>
                <c:axId val="123"/>
                <mc:AlternateContent>
                  <mc:Choice Requires="x16">
                    <c:title>
                      <mc:AlternateContent>
                        <mc:Choice Requires="x16">
                          <c:tx><c:v>X Axis</c:v></c:tx>
                        </mc:Choice>
                        <mc:Fallback>
                          <c:tx><c:v>Fallback Axis</c:v></c:tx>
                        </mc:Fallback>
                      </mc:AlternateContent>
                      <mc:AlternateContent>
                        <mc:Choice Requires="x16">
                          <c:spPr>
                            <a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>
                            <a:ln w="25400">
                              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                            </a:ln>
                          </c:spPr>
                        </mc:Choice>
                        <mc:Fallback>
                          <c:spPr>
                            <a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>
                          </c:spPr>
                        </mc:Fallback>
                      </mc:AlternateContent>
                    </c:title>
                  </mc:Choice>
                  <mc:Fallback>
                    <c:title>
                      <c:tx><c:v>Wrong Axis</c:v></c:tx>
                      <c:spPr>
                        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                      </c:spPr>
                    </c:title>
                  </mc:Fallback>
                </mc:AlternateContent>
              </c:catAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "in-memory-chart.xml").unwrap();

    let title = model.title.expect("title parsed");
    let title_style = title.box_style.expect("title spPr parsed");
    let title_fill = title_style.fill.expect("title fill parsed");
    let FillStyle::Solid(title_fill) = title_fill else {
        panic!("expected title fill to be solidFill, got {title_fill:?}");
    };
    assert_eq!(title_fill.color, Color::Argb(0xFF00FF00));
    let title_line = title_style.line.expect("title ln parsed");
    assert_eq!(title_line.color, Some(Color::Argb(0xFFFF0000)));
    assert_eq!(title_line.width_100pt, Some(100));

    let legend = model.legend.expect("legend parsed");
    let legend_style = legend.style.expect("legend spPr parsed");
    let legend_fill = legend_style.fill.expect("legend fill parsed");
    let FillStyle::Solid(legend_fill) = legend_fill else {
        panic!("expected legend fill to be solidFill, got {legend_fill:?}");
    };
    assert_eq!(legend_fill.color, Color::Argb(0xFF0000FF));
    let legend_line = legend_style.line.as_ref().expect("legend ln parsed");
    assert_eq!(legend_line.width_100pt, Some(300));
    assert_eq!(legend_line.color, Some(Color::Argb(0xFF00FFFF)));

    let axis = model
        .axes
        .iter()
        .find(|axis| axis.id == 123)
        .expect("axis parsed");
    let axis_title = axis.title.as_ref().expect("axis title parsed");
    let axis_box_style = axis_title.box_style.as_ref().expect("axis title spPr parsed");
    let axis_fill = axis_box_style.fill.as_ref().expect("axis title fill parsed");
    let FillStyle::Solid(axis_fill) = axis_fill else {
        panic!("expected axis title fill to be solidFill, got {axis_fill:?}");
    };
    assert_eq!(axis_fill.color, Color::Argb(0xFFFFFF00));
    let axis_line = axis_box_style.line.as_ref().expect("axis title line");
    assert_eq!(axis_line.width_100pt, Some(200));
    assert_eq!(axis_line.color, Some(Color::Argb(0xFF000000)));
}
