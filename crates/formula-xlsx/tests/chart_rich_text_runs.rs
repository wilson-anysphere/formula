use formula_model::Color;
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_chart_title_rich_text_runs_with_formatting() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400">
                <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                <a:latin typeface="Calibri"/>
              </a:rPr>
              <a:t>Hello</a:t>
            </a:r>
            <a:r>
              <a:rPr i="1" sz="1200">
                <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                <a:latin typeface="Arial"/>
              </a:rPr>
              <a:t>World</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    let model = parse_chart_space(xml.as_bytes(), "xl/charts/chart1.xml").expect("parse chartSpace");
    let title = model.title.expect("title parsed");

    assert_eq!(title.rich_text.text, "HelloWorld");
    assert_eq!(title.rich_text.runs.len(), 2);

    let run0 = &title.rich_text.runs[0];
    assert_eq!((run0.start, run0.end), (0, 5));
    assert_eq!(run0.style.bold, Some(true));
    assert_eq!(run0.style.italic, None);
    assert_eq!(run0.style.font.as_deref(), Some("Calibri"));
    assert_eq!(run0.style.size_100pt, Some(1400));
    assert_eq!(run0.style.color, Some(Color::new_argb(0xFFFF0000)));

    let run1 = &title.rich_text.runs[1];
    assert_eq!((run1.start, run1.end), (5, 10));
    assert_eq!(run1.style.bold, None);
    assert_eq!(run1.style.italic, Some(true));
    assert_eq!(run1.style.font.as_deref(), Some("Arial"));
    assert_eq!(run1.style.size_100pt, Some(1200));
    assert_eq!(run1.style.color, Some(Color::new_argb(0xFF00FF00)));
}

