use chrono::NaiveDate;
use formula_model::pivots::slicers::{SlicerSelection, TimelineSelection};
use formula_model::pivots::{DataTable, PivotManager, ScalarValue};
use std::collections::HashSet;

#[test]
fn slicer_filters_pivot_and_updates_chart() -> Result<(), String> {
    let table = DataTable::new(
        vec![
            "Region".to_string(),
            "Date".to_string(),
            "Sales".to_string(),
        ],
        vec![
            vec![
                ScalarValue::from("East"),
                ScalarValue::from(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap()),
                ScalarValue::from(10.0),
            ],
            vec![
                ScalarValue::from("West"),
                ScalarValue::from(NaiveDate::from_ymd_opt(2024, 1, 1).unwrap()),
                ScalarValue::from(7.0),
            ],
            vec![
                ScalarValue::from("East"),
                ScalarValue::from(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap()),
                ScalarValue::from(3.0),
            ],
        ],
    )?;

    let mut pivots = PivotManager::new();
    let pivot_id = pivots.create_pivot_table(
        "SalesByRegion",
        table.clone(),
        vec!["Region".to_string()],
        "Sales",
    )?;
    let chart_id = pivots.create_pivot_chart(pivot_id, "Sales chart")?;

    {
        let output = pivots
            .pivot_output(pivot_id)
            .ok_or_else(|| "missing pivot output".to_string())?;
        assert_eq!(output.headers, vec!["Region", "Sales"]);
        assert_eq!(output.rows.len(), 2);
    }

    {
        let chart = pivots
            .chart_data(chart_id)
            .ok_or_else(|| "missing chart data".to_string())?;
        assert_eq!(chart.categories.len(), 2);
        assert_eq!(chart.values.len(), 2);
    }

    let slicer_id = pivots.add_slicer_to_pivot(pivot_id, "Region slicer", "Region")?;
    let east_only = SlicerSelection::Items(HashSet::from([ScalarValue::from("East")]));
    pivots.set_slicer_selection(slicer_id, east_only)?;

    {
        let output = pivots
            .pivot_output(pivot_id)
            .ok_or_else(|| "missing pivot output".to_string())?;
        assert_eq!(output.rows.len(), 1);
        assert_eq!(output.rows[0][0], ScalarValue::from("East"));
        assert_eq!(output.rows[0][1].as_f64().unwrap(), 13.0);
    }

    {
        let chart = pivots
            .chart_data(chart_id)
            .ok_or_else(|| "missing chart data".to_string())?;
        assert_eq!(chart.categories, vec![vec![ScalarValue::from("East")]]);
        assert_eq!(chart.values, vec![13.0]);
    }

    let timeline_id = pivots.add_timeline_to_pivot(pivot_id, "Date timeline", "Date")?;
    pivots.set_timeline_selection(
        timeline_id,
        TimelineSelection {
            start: Some(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap()),
            end: Some(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap()),
        },
    )?;

    {
        let output = pivots
            .pivot_output(pivot_id)
            .ok_or_else(|| "missing pivot output".to_string())?;
        assert_eq!(output.rows.len(), 1);
        assert_eq!(output.rows[0][0], ScalarValue::from("East"));
        assert_eq!(output.rows[0][1].as_f64().unwrap(), 3.0);
    }

    Ok(())
}
