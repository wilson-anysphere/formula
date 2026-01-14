use std::fs;
use std::path::PathBuf;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

#[test]
fn close_requested_flow_checks_trusted_origin_before_emitting_events() {
    // This is a *source-level* guardrail (headless; no `desktop` feature) that prevents
    // regressions where workbook-derived updates could be emitted to an untrusted origin.
    let main_rs_path = repo_path("src/main.rs");
    let main_rs_src = fs::read_to_string(&main_rs_path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", main_rs_path.display()));

    let start_marker = "tauri::WindowEvent::CloseRequested";
    let start = main_rs_src
        .find(start_marker)
        .unwrap_or_else(|| panic!("failed to find `{start_marker}` in src/main.rs"));

    // Only scan the CloseRequested match arm to avoid picking up unrelated uses.
    let rest = &main_rs_src[start..];
    let end_marker = "tauri::WindowEvent::DragDrop";
    let end = rest.find(end_marker).unwrap_or_else(|| {
        panic!("failed to find `{end_marker}` after `{start_marker}` in src/main.rs")
    });
    let block = &rest[..end];

    let trust_check_pos = block.find("ensure_stable_origin").unwrap_or_else(|| {
        panic!("expected CloseRequested handler to call `desktop::ipc_origin::ensure_stable_origin`")
    });

    let close_prep_emit_pos = block.find("emit(\"close-prep\"").unwrap_or_else(|| {
        panic!("expected CloseRequested handler to emit the `close-prep` event")
    });
    assert!(
        trust_check_pos < close_prep_emit_pos,
        "CloseRequested handler must check `ensure_stable_origin` before emitting `close-prep`"
    );

    let close_requested_emit_pos = block.find("emit(\"close-requested\"").unwrap_or_else(|| {
        panic!("expected CloseRequested handler to emit the `close-requested` event")
    });
    assert!(
        trust_check_pos < close_requested_emit_pos,
        "CloseRequested handler must check `ensure_stable_origin` before emitting `close-requested`"
    );
}
