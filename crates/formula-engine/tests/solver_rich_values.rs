use formula_engine::solver::EngineSolverModel;
use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::Engine;
use formula_engine::Value;

#[test]
fn solver_model_new_errors_on_entity_var_value_and_includes_display_string() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");

    // Objective/constraints can be empty for this test; the model should fail while validating
    // decision-variable cells.
    engine.set_cell_value("Sheet1", "B1", 0.0).unwrap();

    // Rich scalar value whose display string is non-numeric. Solver should surface the display
    // text in the error message.
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue::new("Entity display")),
        )
        .unwrap();

    let err = match EngineSolverModel::new(&mut engine, "Sheet1", "B1", vec!["A1"], Vec::new()) {
        Ok(_) => panic!("expected EngineSolverModel::new to fail for non-numeric var cell"),
        Err(err) => err,
    };

    assert!(
        err.message.contains("Sheet1!A1"),
        "error should mention the offending cell: {err}"
    );
    assert!(
        err.message.contains("Entity display"),
        "error should include the cell's display string: {err}"
    );
    assert!(
        !err.message.contains("#VALUE!"),
        "error should not replace the display string with a generic #VALUE!: {err}"
    );
}

#[test]
fn solver_model_new_errors_on_record_var_value_and_includes_display_string() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");

    engine.set_cell_value("Sheet1", "B1", 0.0).unwrap();
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::new("Record display")),
        )
        .unwrap();

    let err = match EngineSolverModel::new(&mut engine, "Sheet1", "B1", vec!["A1"], Vec::new()) {
        Ok(_) => panic!("expected EngineSolverModel::new to fail for non-numeric var cell"),
        Err(err) => err,
    };

    assert!(
        err.message.contains("Sheet1!A1"),
        "error should mention the offending cell: {err}"
    );
    assert!(
        err.message.contains("Record display"),
        "error should include the cell's display string: {err}"
    );
    assert!(
        !err.message.contains("#VALUE!"),
        "error should not replace the display string with a generic #VALUE!: {err}"
    );
}
