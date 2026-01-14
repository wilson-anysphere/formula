use std::process::{Command, Stdio};

#[test]
fn cli_does_not_panic_on_broken_pipe() {
    let fixture_path =
        concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/basic-encrypted.xlsm");

    let mut child = Command::new(assert_cmd::cargo::cargo_bin!("formula-vba-oracle-cli"))
        .args([
            "extract",
            "--input",
            fixture_path,
            "--format",
            "auto",
            "--password",
            "password",
        ])
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn formula-vba-oracle-cli");

    // Closing the read end forces stdout writes to return EPIPE / BrokenPipe.
    drop(child.stdout.take());

    let output = child
        .wait_with_output()
        .expect("wait for formula-vba-oracle-cli to finish");

    assert!(
        output.status.success(),
        "expected success even when stdout is closed\nstderr:\n{}",
        String::from_utf8_lossy(&output.stderr)
    );
}

