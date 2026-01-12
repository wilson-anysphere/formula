use formula_engine::{Engine, Value};
use formula_model::CellValue;
use formula_xlsx::{read_workbook, write_workbook};
use tempfile::tempdir;

#[test]
fn loads_table_and_evaluates_structured_refs() {
    let wb = read_workbook("tests/fixtures/table.xlsx").expect("read fixture workbook");
    assert_eq!(wb.sheets.len(), 1);
    let sheet = &wb.sheets[0];
    assert_eq!(sheet.tables.len(), 1);
    let table = &sheet.tables[0];
    assert_eq!(table.name, "Table1");
    assert_eq!(table.range.to_string(), "A1:D4");
    assert_eq!(table.columns.len(), 4);
    assert_eq!(table.columns[1].name, "Qty");

    let mut engine = Engine::new();
    engine.set_sheet_tables(&sheet.name, sheet.tables.clone());

    let mut formulas: Vec<(String, String)> = Vec::new();
    for (cell_ref, cell) in sheet.iter_cells() {
        let a1 = cell_ref.to_a1();
        if let Some(formula) = &cell.formula {
            assert!(
                !formula.starts_with('='),
                "formula_model::Cell.formula should be stored without leading '='"
            );
            formulas.push((a1, formula.clone()));
            continue;
        }
        set_engine_value(&mut engine, &sheet.name, &a1, &cell.value);
    }
    for (a1, formula) in formulas {
        engine
            .set_cell_formula(&sheet.name, &a1, &formula)
            .expect("set formula");
    }
    engine.recalculate();

    // Table total column uses this-row structured refs: [@Qty]*[@Price].
    assert_eq!(engine.get_cell_value(&sheet.name, "D2"), Value::Number(6.0));

    // SUM over a column uses Table1[Column].
    assert_eq!(engine.get_cell_value(&sheet.name, "E1"), Value::Number(20.0));

    // Headers selection.
    assert_eq!(
        engine.get_cell_value(&sheet.name, "F1"),
        Value::Text("Qty".into())
    );
}

#[test]
fn round_trips_table_definitions() {
    let wb = read_workbook("tests/fixtures/table.xlsx").expect("read fixture workbook");
    let dir = tempdir().unwrap();
    let out_path = dir.path().join("roundtrip.xlsx");
    write_workbook(&wb, &out_path).expect("write workbook");

    let wb2 = read_workbook(&out_path).expect("read round-tripped workbook");
    assert_eq!(wb2.sheets.len(), wb.sheets.len());
    assert_eq!(wb2.sheets[0].tables, wb.sheets[0].tables);
}

fn set_engine_value(engine: &mut Engine, sheet: &str, addr: &str, value: &CellValue) {
    let v = match value {
        CellValue::Empty => Value::Blank,
        CellValue::Number(n) => Value::Number(*n),
        CellValue::String(s) => Value::Text(s.clone()),
        CellValue::Boolean(b) => Value::Bool(*b),
        CellValue::Error(_) => Value::Error(formula_engine::ErrorKind::Value),
        CellValue::RichText(r) => Value::Text(r.text.clone()),
        CellValue::Entity(e) => Value::Text(e.display_value.clone()),
        CellValue::Record(r) => Value::Text(r.to_string()),
        CellValue::Image(_) => Value::Blank,
        CellValue::Array(_) | CellValue::Spill(_) => Value::Blank,
    };
    engine
        .set_cell_value(sheet, addr, v)
        .expect("set cell value");
}
