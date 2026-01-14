//! Guardrails for `file-dropped` event emission.
//!
//! This is a source-level test so it runs in headless CI without the `desktop` feature (which
//! would pull in the system WebView toolchain on Linux).

#[test]
fn file_dropped_event_is_gated_by_trusted_origin_and_payload_limits() {
    let main_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/main.rs"));

    let start = main_rs
        .find("tauri::WindowEvent::DragDrop")
        .expect("src/main.rs missing DragDrop window event handler");
    let end = main_rs[start..]
        .find(".setup(")
        .map(|idx| start + idx)
        .expect("failed to bound DragDrop handler block (missing .setup(...))");

    let body = &main_rs[start..end];

    let emit_idx = body
        .find("\"file-dropped\"")
        .expect("DragDrop handler missing `file-dropped` emission");

    let origin_check_idx = body
        .find("ensure_stable_origin")
        .expect("`file-dropped` emission must be gated by ipc_origin::ensure_stable_origin");
    assert!(
        origin_check_idx < emit_idx,
        "expected trusted-origin check to occur before `file-dropped` emission"
    );

    let max_paths_idx = body
        .find("MAX_FILE_DROPPED_PATHS")
        .expect("`file-dropped` emission must enforce MAX_FILE_DROPPED_PATHS");
    assert!(
        max_paths_idx < emit_idx,
        "expected MAX_FILE_DROPPED_PATHS enforcement to occur before `file-dropped` emission"
    );

    let max_bytes_idx = body
        .find("MAX_FILE_DROPPED_PATH_BYTES")
        .expect("`file-dropped` emission must enforce MAX_FILE_DROPPED_PATH_BYTES");
    assert!(
        max_bytes_idx < emit_idx,
        "expected MAX_FILE_DROPPED_PATH_BYTES enforcement to occur before `file-dropped` emission"
    );
}
