use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};

fn basic_fixture() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/basic.xlsx")
}

#[test]
fn cli_text_output_does_not_panic_on_broken_pipe() {
    let path = basic_fixture();

    // Simulate a downstream consumer exiting early (e.g. `xlsx-diff ... | head`).
    let mut child = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&path)
        .arg(&path)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn xlsx-diff");

    // Closing the read end forces stdout writes to return EPIPE / BrokenPipe.
    drop(child.stdout.take());

    let output = child
        .wait_with_output()
        .expect("wait for xlsx-diff to finish");

    assert!(
        output.status.success(),
        "expected success even when stdout is closed\nstderr:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
}

#[test]
fn cli_json_output_does_not_panic_on_broken_pipe() {
    let path = basic_fixture();

    let mut child = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&path)
        .arg(&path)
        .arg("--format")
        .arg("json")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn xlsx-diff");

    drop(child.stdout.take());

    let output = child
        .wait_with_output()
        .expect("wait for xlsx-diff to finish");

    assert!(
        output.status.success(),
        "expected success even when stdout is closed\nstderr:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
}

