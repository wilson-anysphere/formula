use std::path::{Path, PathBuf};
use std::process::Command;

const PASSWORD: &str = "password";

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

#[test]
fn cli_errors_without_password() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .output()
        .expect("run xlsx-diff");

    assert!(
        !output.status.success(),
        "expected non-zero exit status when password is missing\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );

    let stderr = String::from_utf8_lossy(&output.stderr).to_ascii_lowercase();
    assert!(
        stderr.contains("password") || stderr.contains("encrypt"),
        "expected error message to mention password/encryption, got:\n{stderr}"
    );
}

#[test]
fn cli_succeeds_with_password() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--password")
        .arg(PASSWORD)
        .output()
        .expect("run xlsx-diff");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("No differences."),
        "expected output to indicate no differences, got:\n{stdout}"
    );
}

#[test]
fn cli_succeeds_with_password_file() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let pw_path = tmp.path().join("password.txt");
    std::fs::write(&pw_path, format!("{PASSWORD}\n")).expect("write password file");

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--password-file")
        .arg(&pw_path)
        .output()
        .expect("run xlsx-diff");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
}

#[test]
fn cli_original_password_overrides_shared_password() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    // Original is encrypted; modified is plain.
    // Provide an incorrect shared password, then override the original password to be correct.
    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&encrypted)
        .arg(&plain)
        .arg("--password")
        .arg("wrong-password")
        .arg("--original-password")
        .arg(PASSWORD)
        .output()
        .expect("run xlsx-diff");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
}

#[test]
fn cli_modified_password_overrides_shared_password() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    // Original is plain; modified is encrypted.
    // Provide an incorrect shared password, then override the modified password to be correct.
    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--password")
        .arg("wrong-password")
        .arg("--modified-password")
        .arg(PASSWORD)
        .output()
        .expect("run xlsx-diff");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
}
