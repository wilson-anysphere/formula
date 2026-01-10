use std::time::Duration;

use formula_vba_runtime::{
    parse_program, InMemoryWorkbook, Spreadsheet, VbaError, VbaRuntime, VbaSandboxPolicy, VbaValue,
};
use pretty_assertions::assert_eq;

fn runtime_from_fixture() -> VbaRuntime {
    let code = include_str!("fixtures/simple.bas");
    let program = parse_program(code).expect("fixture VBA should parse");
    VbaRuntime::new(program)
}

#[test]
fn runs_simple_macro_writes_cells() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "WriteHello", &[]).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::from("Hello")
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::Double(42.0)
    );
}

#[test]
fn supports_for_loop_and_cells() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "FillColumn", &[]).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(1.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::Double(2.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A3").unwrap(),
        VbaValue::Double(3.0)
    );
}

#[test]
fn supports_select_and_activecell() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "SelectAndWrite", &[]).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "B2").unwrap(),
        VbaValue::from("X")
    );
    assert_eq!(wb.active_cell(), (2, 2));
}

#[test]
fn on_error_goto_label_jumps_to_handler() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "HandleError", &[]).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::from("handled")
    );
}

#[test]
fn workbook_open_event_is_callable() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.fire_workbook_open(&mut wb).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::from("opened")
    );
}

#[test]
fn worksheet_change_event_receives_target_range() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();
    let target = wb.range_ref(wb.active_sheet(), "C3").unwrap();

    runtime.fire_worksheet_change(&mut wb, target).unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "C3").unwrap(),
        VbaValue::from("changed")
    );
}

#[test]
fn worksheet_selection_change_event_receives_target_range() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();
    let target = wb.range_ref(wb.active_sheet(), "D4").unwrap();

    runtime
        .fire_worksheet_selection_change(&mut wb, target)
        .unwrap();

    assert_eq!(
        wb.get_value_a1("Sheet1", "D4").unwrap(),
        VbaValue::from("selected")
    );
}

#[test]
fn array_function_and_indexing() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "ArrayTest", &[]).unwrap();
    // VBA `Array(1,2,3)` is 0-based; index 1 is the second element.
    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(2.0)
    );
}

#[test]
fn functions_return_by_assigning_to_function_name() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "CallFunction", &[]).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(42.0)
    );
}

#[test]
fn collection_object_minimal_support() {
    let runtime = runtime_from_fixture();
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "CollectionTest", &[]).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(2.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::Double(2.0)
    );
}

#[test]
fn sandbox_enforces_step_limit() {
    let code = include_str!("fixtures/simple.bas");
    let program = parse_program(code).expect("fixture VBA should parse");
    let sandbox = VbaSandboxPolicy {
        max_execution_time: Duration::from_secs(5),
        max_steps: 500,
        ..VbaSandboxPolicy::default()
    };
    let runtime = VbaRuntime::new(program).with_sandbox_policy(sandbox);
    let mut wb = InMemoryWorkbook::new();

    let err = runtime.execute(&mut wb, "Infinite", &[]).unwrap_err();
    assert!(matches!(err, VbaError::StepLimit));
}
