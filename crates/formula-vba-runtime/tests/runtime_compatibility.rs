use std::time::Duration;

use formula_vba_runtime::{
    parse_program, InMemoryWorkbook, VbaError, VbaRuntime, VbaSandboxPolicy, VbaValue,
};
use pretty_assertions::assert_eq;

#[test]
fn default_member_value_semantics_match_vba() {
    let code = r#"
Sub Test()
    Range("A1") = 7
    Dim x
    x = Range("A1")
    Range("B1") = x
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Double(7.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(7.0));
}

#[test]
fn range_object_model_offset_resize_clear_contents_and_address() {
    let code = r#"
Sub Test()
    Range("A1") = 1
    Range("A1").Offset(, 1) = 2
    Range("A1").Offset(1, 0) = 3
    Range("A1").Offset(1, 1) = 4

    Range("C1") = Range("B1:D3").Rows.Count
    Range("C2") = Range("B1:D3").Columns.Count
    Range("C3") = Range("C5").Address
    Range("C4") = Range("C5:D6").Address

    Range("A1").Resize(2, 2).ClearContents
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    // Cleared range.
    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Empty);
    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Empty);
    assert_eq!(wb.get_value_a1("Sheet1", "A2").unwrap(), VbaValue::Empty);
    assert_eq!(wb.get_value_a1("Sheet1", "B2").unwrap(), VbaValue::Empty);

    // Rows/cols count.
    assert_eq!(wb.get_value_a1("Sheet1", "C1").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "C2").unwrap(), VbaValue::Double(3.0));

    // Address uses absolute markers.
    assert_eq!(
        wb.get_value_a1("Sheet1", "C3").unwrap(),
        VbaValue::from("$C$5")
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "C4").unwrap(),
        VbaValue::from("$C$5:$D$6")
    );
}

#[test]
fn range_copy_and_paste_special_values() {
    let code = r#"
Sub Test()
    Range("A1") = "X"
    Range("A1").Copy
    Range("B2").PasteSpecial
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "B2").unwrap(), VbaValue::from("X"));
}

#[test]
fn range_copy_to_single_cell_destination_expands_to_source_size() {
    let code = r#"
Sub Test()
    Range("A1") = 1
    Range("B1") = 2
    Range("A2") = 3
    Range("B2") = 4

    Range("A1:B2").Copy Destination:=Range("D1")
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "D1").unwrap(), VbaValue::Double(1.0));
    assert_eq!(wb.get_value_a1("Sheet1", "E1").unwrap(), VbaValue::Double(2.0));
    assert_eq!(wb.get_value_a1("Sheet1", "D2").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "E2").unwrap(), VbaValue::Double(4.0));
}

#[test]
fn paste_special_expands_multi_cell_clipboard_when_destination_is_single_cell() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1") = 1
    Range("B1") = 2
    Range("A2") = 3
    Range("B2") = 4

    Range("A1:B2").Copy
    Range("D1").PasteSpecial Paste:=xlPasteValues
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "D1").unwrap(), VbaValue::Double(1.0));
    assert_eq!(wb.get_value_a1("Sheet1", "E1").unwrap(), VbaValue::Double(2.0));
    assert_eq!(wb.get_value_a1("Sheet1", "D2").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "E2").unwrap(), VbaValue::Double(4.0));
}

#[test]
fn paste_special_respects_xlpastevalues_constant_and_option_explicit() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1").Copy
    Range("B1").PasteSpecial Paste:=xlPasteValues
    Range("A2").Copy
    Range("C1").PasteSpecial
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();
    wb.set_value_a1("Sheet1", "A1", VbaValue::Double(2.0)).unwrap();
    wb.set_formula_a1("Sheet1", "A2", "=1+1").unwrap();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(2.0));
    assert_eq!(wb.get_formula_a1("Sheet1", "B1").unwrap(), None);
    assert_eq!(
        wb.get_formula_a1("Sheet1", "C1").unwrap(),
        Some("=1+1".to_string())
    );
}

#[test]
fn range_end_uses_direction_constants() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1") = 1
    Range("A2") = 2
    Range("A3") = 3
    Range("B1") = Range("A1").End(xlDown).Row
    Range("B2") = Range("A3").End(xlUp).Row
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B2").unwrap(), VbaValue::Double(1.0));
}

#[test]
fn rows_and_columns_count_match_excel_limits() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1") = Rows.Count
    Range("A2") = Columns.Count
    Range("A3") = ActiveSheet.Rows.Count
    Range("A4") = ActiveSheet.Columns.Count
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "A1").unwrap(),
        VbaValue::Double(1_048_576.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::Double(16_384.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A3").unwrap(),
        VbaValue::Double(1_048_576.0)
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "A4").unwrap(),
        VbaValue::Double(16_384.0)
    );
}

#[test]
fn selection_variable_tracks_last_select() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1:B2").Select
    Range("C1") = Selection.Address
    Range("C2") = Application.Selection.Address
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "C1").unwrap(),
        VbaValue::from("$A$1:$B$2")
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "C2").unwrap(),
        VbaValue::from("$A$1:$B$2")
    );
}

#[test]
fn application_cut_copy_mode_false_clears_clipboard() {
    let code = r#"
Option Explicit

Sub Test()
    Range("A1") = 1
    Range("B1") = 5
    Range("A1").Copy
    Application.CutCopyMode = False
    Range("B1").PasteSpecial Paste:=xlPasteValues
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(5.0));
}

#[test]
fn sandbox_step_limit_applies_to_range_copy() {
    let code = r#"
Sub Test()
    Range("A1:J10").Copy Destination:=Range("K1")
End Sub
"#;
    let program = parse_program(code).unwrap();
    let sandbox = VbaSandboxPolicy {
        max_execution_time: Duration::from_secs(5),
        max_steps: 50,
        ..VbaSandboxPolicy::default()
    };
    let runtime = VbaRuntime::new(program).with_sandbox_policy(sandbox);
    let mut wb = InMemoryWorkbook::new();
    let err = runtime.execute(&mut wb, "Test", &[]).unwrap_err();
    assert!(matches!(err, VbaError::StepLimit));
}

#[test]
fn range_accepts_two_arguments_strings_and_cells() {
    let code = r#"
Option Explicit

Sub Test()
    Range("C1") = Range("A1", "B2").Address
    Range("C2") = Range(Cells(1, 1), Cells(2, 2)).Address
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(
        wb.get_value_a1("Sheet1", "C1").unwrap(),
        VbaValue::from("$A$1:$B$2")
    );
    assert_eq!(
        wb.get_value_a1("Sheet1", "C2").unwrap(),
        VbaValue::from("$A$1:$B$2")
    );
}

#[test]
fn for_each_over_array_and_collection() {
    let code = r#"
Sub Test()
    Dim arr
    arr = Array(1, 2, 3)
    Dim v
    Dim i
    i = 1
    For Each v In arr
        Cells(i, 1) = v
        i = i + 1
    Next v

    Dim c
    Set c = New Collection
    c.Add 10
    c.Add 20
    i = 1
    For Each v In c
        Cells(i, 2) = v
        i = i + 1
    Next v
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Double(1.0));
    assert_eq!(wb.get_value_a1("Sheet1", "A2").unwrap(), VbaValue::Double(2.0));
    assert_eq!(wb.get_value_a1("Sheet1", "A3").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(10.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B2").unwrap(), VbaValue::Double(20.0));
}

#[test]
fn builtins_string_numeric_date_and_operators() {
    let code = r#"
Sub Test()
    Range("A1") = CStr(123)
    Range("A2") = CLng(2.5)
    Range("A3") = CBool(0)
    Range("A4") = UCase("aB")
    Range("A5") = Replace("a-b-c", "-", "_")
    Range("A6") = Left("Hello", 2)
    Range("A7") = Right("Hello", 2)
    Range("A8") = Mid("Hello", 2, 2)
    Range("A9") = Len("Hello")

    Range("B1") = 7 \ 2
    Range("B2") = 7 Mod 4
    Range("B3") = 2 ^ 3
    Range("B4") = "a" & "b"

    Range("C1") = Format(1.2, "0.00")
    Range("C2") = Format(DateAdd("d", 1, CDate("2020-01-01")), "yyyy-mm-dd")
    Range("C3") = DateDiff("d", CDate("2020-01-01"), CDate("2020-01-03"))
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();

    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::from("123"));
    // Banker's rounding: 2.5 -> 2.
    assert_eq!(wb.get_value_a1("Sheet1", "A2").unwrap(), VbaValue::Double(2.0));
    assert_eq!(
        wb.get_value_a1("Sheet1", "A3").unwrap(),
        VbaValue::Boolean(false)
    );
    assert_eq!(wb.get_value_a1("Sheet1", "A4").unwrap(), VbaValue::from("AB"));
    assert_eq!(wb.get_value_a1("Sheet1", "A5").unwrap(), VbaValue::from("a_b_c"));
    assert_eq!(wb.get_value_a1("Sheet1", "A6").unwrap(), VbaValue::from("He"));
    assert_eq!(wb.get_value_a1("Sheet1", "A7").unwrap(), VbaValue::from("lo"));
    assert_eq!(wb.get_value_a1("Sheet1", "A8").unwrap(), VbaValue::from("el"));
    assert_eq!(wb.get_value_a1("Sheet1", "A9").unwrap(), VbaValue::Double(5.0));

    assert_eq!(wb.get_value_a1("Sheet1", "B1").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B2").unwrap(), VbaValue::Double(3.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B3").unwrap(), VbaValue::Double(8.0));
    assert_eq!(wb.get_value_a1("Sheet1", "B4").unwrap(), VbaValue::from("ab"));

    assert_eq!(wb.get_value_a1("Sheet1", "C1").unwrap(), VbaValue::from("1.20"));
    assert_eq!(
        wb.get_value_a1("Sheet1", "C2").unwrap(),
        VbaValue::from("2020-01-02")
    );
    assert_eq!(wb.get_value_a1("Sheet1", "C3").unwrap(), VbaValue::Double(2.0));
}

#[test]
fn createobject_dictionary_is_blocked_by_default_and_allowed_with_permission() {
    let code = r#"
Sub Test()
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    Range("A1") = d.Item("a")
    Range("A2") = d.Count
End Sub
"#;
    let program = parse_program(code).unwrap();

    let mut wb = InMemoryWorkbook::new();
    let runtime = VbaRuntime::new(program.clone());
    let err = runtime.execute(&mut wb, "Test", &[]).unwrap_err();
    assert!(matches!(err, VbaError::Sandbox(_)));

    let sandbox = VbaSandboxPolicy {
        allow_object_creation: true,
        max_execution_time: Duration::from_secs(5),
        ..VbaSandboxPolicy::default()
    };
    let runtime = VbaRuntime::new(program).with_sandbox_policy(sandbox);
    let mut wb = InMemoryWorkbook::new();
    runtime.execute(&mut wb, "Test", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Double(1.0));
    assert_eq!(wb.get_value_a1("Sheet1", "A2").unwrap(), VbaValue::Double(1.0));
}

#[test]
fn on_error_goto_sets_err_and_resume_next_continues() {
    let code = r#"
Sub Test()
    Range("A1") = "start"
    On Error GoTo ErrHandler
    Range().Value = 1
    Range("A2") = "after error"
    Exit Sub
ErrHandler:
    Range("A1") = Err.Description
    Resume Next
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Test", &[]).unwrap();
    let a1 = wb.get_value_a1("Sheet1", "A1").unwrap().to_string_lossy();
    assert!(a1.contains("Range()"), "unexpected Err.Description: {a1}");
    assert_eq!(
        wb.get_value_a1("Sheet1", "A2").unwrap(),
        VbaValue::from("after error")
    );
}

#[test]
fn module_level_globals_persist_across_calls_and_option_explicit_is_enforced() {
    let code = r#"
Option Explicit
Public counter As Long

Sub Inc()
    counter = counter + 1
    Range("A1") = counter
End Sub

Sub Bad()
    x = 1
End Sub
"#;
    let program = parse_program(code).unwrap();
    let runtime = VbaRuntime::new(program);
    let mut wb = InMemoryWorkbook::new();

    runtime.execute(&mut wb, "Inc", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Double(1.0));
    runtime.execute(&mut wb, "Inc", &[]).unwrap();
    assert_eq!(wb.get_value_a1("Sheet1", "A1").unwrap(), VbaValue::Double(2.0));

    let err = runtime.execute(&mut wb, "Bad", &[]).unwrap_err();
    assert!(matches!(err, VbaError::Runtime(_)));
}
