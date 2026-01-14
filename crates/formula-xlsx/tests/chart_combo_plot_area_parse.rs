use formula_model::charts::{ChartKind, ComboChartEntry, PlotAreaModel};
use formula_xlsx::drawingml::charts::parse_chart_space;

#[test]
fn parses_combo_plot_area_with_multiple_chart_types() {
    // Minimal chartSpace containing a barChart + lineChart overlay (Excel combo chart style),
    // each with a single series sharing the same axes.
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:ser>
                  <c:tx><c:v>Bar Series</c:v></c:tx>
                  <c:cat>
                    <c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef>
                  </c:val>
                </c:ser>
                <c:axId val="1"/>
                <c:axId val="2"/>
              </c:barChart>

              <c:lineChart>
                <c:grouping val="standard"/>
                <c:ser>
                  <c:tx><c:v>Line Series</c:v></c:tx>
                  <c:cat>
                    <c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef>
                  </c:val>
                </c:ser>
                <c:axId val="1"/>
                <c:axId val="2"/>
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

    let model = parse_chart_space(xml.as_bytes(), "combo.xml").expect("parse chartSpace");
    assert_eq!(model.chart_kind, ChartKind::Bar);
    assert_eq!(model.series.len(), 2);

    let PlotAreaModel::Combo(combo) = model.plot_area else {
        panic!("expected combo plot area, got {:?}", model.plot_area);
    };
    assert_eq!(combo.charts.len(), 2);

    match &combo.charts[0] {
        ComboChartEntry::Bar { model, series } => {
            assert_eq!(series.start, 0);
            assert_eq!(series.end, 1);
            assert_eq!(model.ax_ids, vec![1, 2]);
        }
        other => panic!("expected first subplot to be bar, got {other:?}"),
    }

    match &combo.charts[1] {
        ComboChartEntry::Line { model, series } => {
            assert_eq!(series.start, 1);
            assert_eq!(series.end, 2);
            assert_eq!(model.ax_ids, vec![1, 2]);
        }
        other => panic!("expected second subplot to be line, got {other:?}"),
    }

    assert_eq!(model.series[0].plot_index, Some(0));
    assert_eq!(model.series[1].plot_index, Some(1));
}

#[test]
fn combo_plot_area_warns_only_for_unknown_chart_types() {
    let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:madeUpChart>
                <c:ser>
                  <c:tx><c:v>Area Series</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
                </c:ser>
              </c:madeUpChart>
              <c:barChart>
                <c:ser>
                  <c:tx><c:v>Bar Series</c:v></c:tx>
                  <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef></c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
    "#;

    let model = parse_chart_space(xml.as_bytes(), "combo-unknown.xml").expect("parse chartSpace");
    assert_eq!(model.series.len(), 2);
    assert_eq!(model.series[0].plot_index, Some(0));
    assert_eq!(model.series[1].plot_index, Some(1));

    let PlotAreaModel::Combo(combo) = model.plot_area else {
        panic!("expected combo plot area");
    };
    assert_eq!(combo.charts.len(), 2);

    // We should only warn about the unsupported chart type, not because multiple chart types exist.
    assert_eq!(model.diagnostics.len(), 1);
    assert!(
        model.diagnostics[0]
            .message
            .contains("unsupported chart type madeUpChart"),
        "diagnostic was {:?}",
        model.diagnostics[0]
    );
}
