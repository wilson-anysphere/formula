use desktop::commands::{
    PythonFilesystemPermission, PythonNetworkPermission, PythonPermissions,
};
use desktop::python::run_python_script;
use desktop::state::AppState;

#[test]
fn rejects_python_permission_escalation_without_spawning_python() {
    // When developers explicitly opt-in to unsafe permissions locally, the guard is
    // bypassed (debug builds only). Skip this test in that configuration.
    if cfg!(debug_assertions)
        && matches!(
            std::env::var("FORMULA_UNSAFE_PYTHON_PERMISSIONS")
                .unwrap_or_default()
                .trim(),
            "1" | "true" | "TRUE" | "yes" | "YES"
        )
    {
        eprintln!("FORMULA_UNSAFE_PYTHON_PERMISSIONS enabled; skipping");
        return;
    }

    let mut state = AppState::new();
    let err = run_python_script(
        &mut state,
        "print('hello')",
        Some(PythonPermissions {
            filesystem: PythonFilesystemPermission::Read,
            network: PythonNetworkPermission::None,
            network_allowlist: None,
        }),
        None,
        None,
        None,
    )
    .expect_err("permission escalation should be rejected");

    assert!(
        err.contains("Python permission escalation is not supported yet"),
        "unexpected error: {err}"
    );
}
