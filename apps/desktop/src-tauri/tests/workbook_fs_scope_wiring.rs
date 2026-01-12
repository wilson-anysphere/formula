//! Guardrails for workbook open/save filesystem scope enforcement.
//!
//! These are source-level tests so they run in headless CI without the `desktop` feature
//! (which would pull in the system WebView toolchain on Linux).

#[test]
fn open_workbook_enforces_fs_scope_and_uses_canonical_path_for_autosave_db() {
    let commands_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/commands.rs"));

    let start = commands_rs
        .find("pub async fn open_workbook")
        .expect("commands.rs missing open_workbook()");
    let end = commands_rs[start..]
        .find("pub async fn new_workbook")
        .map(|idx| start + idx)
        .expect("commands.rs missing new_workbook() (used to bound open_workbook)");

    let body = &commands_rs[start..end];

    assert!(
        body.contains("canonicalize_in_allowed_roots"),
        "open_workbook must scope-check via fs_scope::canonicalize_in_allowed_roots"
    );
    assert!(
        body.contains("read_workbook(resolved)"),
        "open_workbook must pass the canonicalized path into read_workbook"
    );
    assert!(
        body.contains("autosave_db_path_for_workbook(&resolved_str)"),
        "open_workbook must compute autosave DB location using the canonicalized path"
    );
}

#[test]
fn save_workbook_validates_destination_with_fs_scope_and_marks_saved_with_canonical_path() {
    let commands_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/commands.rs"));

    let start = commands_rs
        .find("pub async fn save_workbook")
        .expect("commands.rs missing save_workbook()");
    let end = commands_rs[start..]
        .find("pub fn mark_saved")
        .map(|idx| start + idx)
        .expect("commands.rs missing mark_saved() (used to bound save_workbook)");

    let body = &commands_rs[start..end];

    assert!(
        body.contains("coerce_save_path_to_xlsx"),
        "save_workbook must coerce destination extension before validating scope"
    );
    assert!(
        body.contains("resolve_save_path_in_allowed_roots"),
        "save_workbook must validate destination via fs_scope::resolve_save_path_in_allowed_roots"
    );
    assert!(
        body.contains("Some(validated_save_path)"),
        "save_workbook must persist the validated (canonicalized) path via mark_saved"
    );
}

