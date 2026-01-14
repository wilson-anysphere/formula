use formula_model::pivots::{PivotChartId, PivotTableId};
use formula_xlsx::load_from_bytes;

const FIXTURE: &[u8] = include_bytes!("fixtures/pivot_slicers_and_chart.xlsx");

// Keep in sync with `crates/formula-xlsx/src/read/mod.rs`.
const PIVOT_BINDING_NAMESPACE: PivotTableId = PivotTableId::from_u128(0xaa5f186245314193be90229689c7d364);

fn expected_pivot_table_id(part_name: &str) -> PivotTableId {
    let mut key = String::with_capacity("pivotTable:".len() + part_name.len());
    key.push_str("pivotTable:");
    key.push_str(part_name);
    PivotTableId::new_v5(&PIVOT_BINDING_NAMESPACE, key.as_bytes())
}

fn expected_pivot_chart_id(part_name: &str) -> PivotChartId {
    let mut key = String::with_capacity("pivotChart:".len() + part_name.len());
    key.push_str("pivotChart:");
    key.push_str(part_name);
    PivotChartId::new_v5(&PIVOT_BINDING_NAMESPACE, key.as_bytes())
}

#[test]
fn imports_pivot_chart_bindings_into_workbook_model() {
    let doc = load_from_bytes(FIXTURE).expect("load fixture");

    assert_eq!(doc.workbook.pivot_charts.len(), 1);

    let chart = &doc.workbook.pivot_charts[0];
    assert_eq!(chart.name, "PivotTable1");
    assert_eq!(chart.chart_part.as_deref(), Some("xl/charts/chart1.xml"));
    assert_eq!(
        chart.pivot_table_id,
        expected_pivot_table_id("xl/pivotTables/pivotTable1.xml")
    );
    assert_eq!(chart.id, expected_pivot_chart_id("xl/charts/chart1.xml"));
}

