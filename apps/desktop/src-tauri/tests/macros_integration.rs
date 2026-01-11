use formula_desktop_tauri::file_io::read_xlsx_blocking;
use formula_desktop_tauri::file_io::Workbook;
use formula_desktop_tauri::macros::{MacroExecutionOptions, MacroPermission};
use formula_desktop_tauri::state::{AppState, CellScalar};
use std::io::Write;
use std::path::Path;

fn load_basic_xlsm_fixture() -> Workbook {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../../fixtures/xlsx/macros/basic.xlsm"
    );
    read_xlsx_blocking(Path::new(fixture_path)).expect("read xlsm fixture")
}

fn build_vba_project_bin(module_code: &str) -> Vec<u8> {
    use std::io::Cursor;

    let module_container = formula_vba::compress_container(module_code.as_bytes());

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTCODEPAGE (u16 LE) - 1252
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");

        // Single standard module named Module1.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes()); // reserved
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
        out
    };
    let dir_container = formula_vba::compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    ole.into_inner().into_inner()
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

#[test]
fn loads_fixture_xlsm_lists_and_runs_macro() {
    let workbook = load_basic_xlsm_fixture();
    assert!(
        workbook.vba_project_bin.is_some(),
        "fixture should include vbaProject.bin"
    );

    let mut state = AppState::new();
    let info = state.load_workbook(workbook);
    let sheet_id = info.sheets[0].id.clone();

    let macros = state.list_macros().expect("list macros");
    assert!(
        macros.iter().any(|m| m.id == "WriteCells"),
        "expected WriteCells macro in fixture, got: {macros:?}"
    );

    let outcome = state
        .run_macro(
            "WriteCells",
            MacroExecutionOptions {
                permissions: Vec::new(),
                timeout_ms: Some(2_000),
            },
        )
        .expect("run macro");
    assert!(outcome.ok, "macro should succeed: {outcome:?}");
    assert!(
        outcome
            .updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == 0 && u.col == 0),
        "expected updates to include A1"
    );

    let a1 = state.get_cell(&sheet_id, 0, 0).unwrap().value;
    assert_eq!(a1, CellScalar::Text("Written".to_string()));

    let b2 = state.get_cell(&sheet_id, 1, 1).unwrap().value;
    assert_eq!(b2, CellScalar::Number(42.0));
}

#[test]
fn macro_context_range_and_events_work_end_to_end() {
    let module_code = r#"
Sub FillRange()
    Range("A1:B2").Value = 5
End Sub

Sub SelectB2()
    Range("B2").Select
End Sub

Sub WriteActive()
    ActiveCell.Value = "X"
End Sub

Sub Worksheet_Change(Target)
    Target.Value = "changed"
End Sub

Sub Worksheet_SelectionChange(Target)
    Target.Value = "selected"
End Sub

Sub Workbook_BeforeClose(Cancel)
    Range("A1").Value = "closing"
End Sub

Sub TryCreateObject()
    CreateObject("WScript.Shell")
End Sub
"#;

    let vba_bin = build_vba_project_bin(module_code);
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.vba_project_bin = Some(vba_bin);

    let mut state = AppState::new();
    let info = state.load_workbook(workbook);
    let sheet_id = info.sheets[0].id.clone();

    // Seed a formula to ensure macro writes surface recalculation updates.
    state
        .set_cell(&sheet_id, 0, 2, None, Some("=A1+1".to_string()))
        .expect("seed formula");

    // Range.Value scalar assignment should fill the range.
    let outcome = state
        .run_macro("FillRange", MacroExecutionOptions::default())
        .expect("FillRange");
    assert!(outcome.ok);
    for row in 0..2 {
        for col in 0..2 {
            assert_eq!(
                state.get_cell(&sheet_id, row, col).unwrap().value,
                CellScalar::Number(5.0),
                "expected ({row},{col}) to be 5"
            );
        }
    }
    // Recalc update: C1 (=A1+1) should now be 6.
    assert_eq!(
        state.get_cell(&sheet_id, 0, 2).unwrap().value,
        CellScalar::Number(6.0)
    );

    // Persistent macro context across calls (ActiveCell).
    state
        .run_macro("SelectB2", MacroExecutionOptions::default())
        .expect("SelectB2");
    state
        .run_macro("WriteActive", MacroExecutionOptions::default())
        .expect("WriteActive");
    assert_eq!(
        state.get_cell(&sheet_id, 1, 1).unwrap().value,
        CellScalar::Text("X".to_string())
    );

    // Worksheet_Change should receive target range.
    let changed = state
        .fire_worksheet_change(&sheet_id, 2, 2, 2, 2, MacroExecutionOptions::default())
        .expect("Worksheet_Change");
    assert!(changed.ok);
    assert_eq!(
        state.get_cell(&sheet_id, 2, 2).unwrap().value,
        CellScalar::Text("changed".to_string())
    );

    // Worksheet_SelectionChange should receive target range.
    let selected = state
        .fire_selection_change(&sheet_id, 3, 3, 3, 3, MacroExecutionOptions::default())
        .expect("Worksheet_SelectionChange");
    assert!(selected.ok);
    assert_eq!(
        state.get_cell(&sheet_id, 3, 3).unwrap().value,
        CellScalar::Text("selected".to_string())
    );

    // Workbook_BeforeClose fires without blocking and can mutate the workbook.
    let closing = state
        .fire_workbook_before_close(MacroExecutionOptions::default())
        .expect("Workbook_BeforeClose");
    assert!(closing.ok);
    assert_eq!(
        state.get_cell(&sheet_id, 0, 0).unwrap().value,
        CellScalar::Text("closing".to_string())
    );

    // Sandbox should deny blocked actions and surface a permission request.
    let denied = state
        .run_macro("TryCreateObject", MacroExecutionOptions::default())
        .expect("TryCreateObject executes (but should be denied)");
    assert!(!denied.ok, "expected sandbox violation");
    let request = denied
        .permission_request
        .expect("expected permission_request payload");
    assert_eq!(request.macro_id, "TryCreateObject");
    assert_eq!(request.requested, vec![MacroPermission::ObjectCreation]);
}

#[test]
fn macro_ui_context_sets_active_sheet_cell_and_selection() {
    let module_code = r#"
Sub WriteActive()
    ActiveCell.Value = "OK"
End Sub

Sub FillSelection()
    Selection.Value = "S"
End Sub
"#;

    let vba_bin = build_vba_project_bin(module_code);
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());
    workbook.vba_project_bin = Some(vba_bin);

    let mut state = AppState::new();
    let info = state.load_workbook(workbook);
    let sheet1_id = info.sheets[0].id.clone();
    let sheet2_id = info.sheets[1].id.clone();

    // Set UI context to Sheet2!C3 before running a macro that writes ActiveCell.
    state
        .set_macro_ui_context(&sheet2_id, 2, 2, None)
        .expect("set macro ui context");
    let outcome = state
        .run_macro("WriteActive", MacroExecutionOptions::default())
        .expect("WriteActive runs");
    assert!(outcome.ok, "expected macro to succeed: {outcome:?}");
    assert_eq!(
        state.get_cell(&sheet2_id, 2, 2).unwrap().value,
        CellScalar::Text("OK".to_string())
    );
    assert_eq!(
        state.get_cell(&sheet1_id, 2, 2).unwrap().value,
        CellScalar::Empty
    );

    // Selection should also reflect the UI before macro execution.
    state
        .set_macro_ui_context(
            &sheet2_id,
            1,
            1,
            Some(formula_desktop_tauri::state::CellRect {
                start_row: 1,
                start_col: 1,
                end_row: 2,
                end_col: 2,
            }),
        )
        .expect("set selection context");
    let selected = state
        .run_macro("FillSelection", MacroExecutionOptions::default())
        .expect("FillSelection runs");
    assert!(selected.ok, "expected macro to succeed: {selected:?}");
    for row in 1..=2 {
        for col in 1..=2 {
            assert_eq!(
                state.get_cell(&sheet2_id, row, col).unwrap().value,
                CellScalar::Text("S".to_string()),
                "expected Sheet2!({}, {}) to be 'S'",
                row,
                col
            );
        }
    }
}
