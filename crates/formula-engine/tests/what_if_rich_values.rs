use formula_engine::eval::CellAddr;
use formula_engine::functions::{Reference, SheetId};
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::what_if::{CellRef, CellValue, EngineWhatIfModel, WhatIfModel};
use formula_engine::{Engine, Value};

#[test]
fn what_if_get_cell_value_degrades_rich_engine_values_to_text_display_string() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");

    let reference = Reference {
        sheet_id: SheetId::Local(0),
        start: CellAddr { row: 0, col: 0 },
        end: CellAddr { row: 0, col: 0 },
    };
    let rich_value = Value::Reference(reference);
    let expected = rich_value.to_string();

    engine.set_cell_value("Sheet1", "A1", rich_value).unwrap();

    let model = EngineWhatIfModel::new(&mut engine, "Sheet1");
    let value = model.get_cell_value(&CellRef::from("A1")).unwrap();

    assert_eq!(value, CellValue::Text(expected));
}

#[test]
fn what_if_get_cell_value_degrades_entity_values_to_display_text() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");

    let entity = Value::Entity(EntityValue::new("Entity display"));
    engine.set_cell_value("Sheet1", "A1", entity).unwrap();

    let model = EngineWhatIfModel::new(&mut engine, "Sheet1");
    let value = model.get_cell_value(&CellRef::from("A1")).unwrap();

    assert_eq!(value, CellValue::Text("Entity display".to_string()));
}

#[test]
fn what_if_get_cell_value_degrades_record_values_to_display_text() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");

    let record = Value::Record(RecordValue::new("Record display"));
    engine.set_cell_value("Sheet1", "A1", record).unwrap();

    let model = EngineWhatIfModel::new(&mut engine, "Sheet1");
    let value = model.get_cell_value(&CellRef::from("A1")).unwrap();

    assert_eq!(value, CellValue::Text("Record display".to_string()));
}
