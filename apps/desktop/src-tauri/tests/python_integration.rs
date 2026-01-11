use formula_desktop_tauri::commands::PythonRunContext;
use formula_desktop_tauri::file_io::Workbook;
use formula_desktop_tauri::python::run_python_script;
use formula_desktop_tauri::state::{AppState, CellScalar};
use std::process::Command;

fn python_executable() -> String {
    std::env::var("FORMULA_PYTHON_EXECUTABLE").unwrap_or_else(|_| "python3".to_string())
}

fn python_available() -> bool {
    Command::new(python_executable())
        .arg("--version")
        .output()
        .is_ok()
}

#[test]
fn runs_native_python_script_and_returns_cell_updates() {
    if !python_available() {
        eprintln!("python3 not available; skipping native python integration test");
        return;
    }

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());

    let mut state = AppState::new();
    let info = state.load_workbook(workbook);
    let sheet_id = info.sheets[0].id.clone();

    let code = r#"import formula

sheet = formula.active_sheet
sheet["A1"] = 1
sheet["A2"] = "=A1*2"
"#;

    let result = run_python_script(
        &mut state,
        code,
        None,
        Some(5_000),
        Some(256 * 1024 * 1024),
        Some(PythonRunContext {
            active_sheet_id: Some(sheet_id.clone()),
            selection: None,
        }),
    )
    .expect("run python");

    assert!(result.ok, "python run should succeed: {result:?}");
    assert!(
        result
            .updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == 0 && u.col == 0),
        "expected updates to include A1"
    );
    assert!(
        result
            .updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == 1 && u.col == 0),
        "expected updates to include A2"
    );

    assert_eq!(
        state.get_cell(&sheet_id, 0, 0).unwrap().value,
        CellScalar::Number(1.0)
    );
    assert_eq!(
        state.get_cell(&sheet_id, 1, 0).unwrap().formula,
        Some("=A1*2".to_string())
    );
    assert_eq!(
        state.get_cell(&sheet_id, 1, 0).unwrap().value,
        CellScalar::Number(2.0)
    );
}
