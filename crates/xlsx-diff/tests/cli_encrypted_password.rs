use std::path::{Path, PathBuf};
use std::process::Command;
use std::process::Stdio;

use std::io::Write;

const PASSWORD: &str = "password";
const UNICODE_PASSWORD: &str = "pässwörd";
const UNICODE_PASSWORD_WITH_EMOJI: &str = "pässwörd🔒";

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
    // Ensure we tolerate Windows-style line endings in password files.
    std::fs::write(&pw_path, format!("{PASSWORD}\r\n")).expect("write password file");

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

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("No differences."),
        "expected output to indicate no differences, got:\n{stdout}"
    );
}

#[test]
fn cli_succeeds_with_empty_password_file() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile-empty-password.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let pw_path = tmp.path().join("password.txt");
    // Trailing newlines are trimmed; a file containing only `\n` represents an empty password.
    std::fs::write(&pw_path, "\r\n").expect("write password file");

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

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("No differences."),
        "expected output to indicate no differences, got:\n{stdout}"
    );
}

#[test]
fn cli_succeeds_with_unicode_password_file() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile-unicode.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let pw_path = tmp.path().join("password.txt");
    std::fs::write(&pw_path, format!("{UNICODE_PASSWORD}\r\n")).expect("write password file");

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

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("No differences."),
        "expected output to indicate no differences, got:\n{stdout}"
    );
}

#[test]
fn cli_succeeds_with_unicode_emoji_password_file() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let pw_path = tmp.path().join("password.txt");
    std::fs::write(&pw_path, format!("{UNICODE_PASSWORD_WITH_EMOJI}\r\n"))
        .expect("write password file");

    for (plain_name, encrypted_name) in [
        ("plaintext.xlsx", "standard-unicode.xlsx"),
        ("plaintext-excel.xlsx", "agile-unicode-excel.xlsx"),
    ] {
        let plain = fixture_path(plain_name);
        let encrypted = fixture_path(encrypted_name);

        let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
            .arg(&plain)
            .arg(&encrypted)
            .arg("--password-file")
            .arg(&pw_path)
            .output()
            .unwrap_or_else(|err| panic!("run xlsx-diff ({encrypted_name}): {err}"));

        assert!(
            output.status.success(),
            "{encrypted_name}: expected exit 0\nstdout:\n{}\nstderr:\n{}",
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr),
        );

        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(
            stdout.contains("No differences."),
            "{encrypted_name}: expected output to indicate no differences, got:\n{stdout}"
        );
    }
}

#[test]
fn cli_succeeds_with_encrypted_xlsm_fixtures() {
    let plain = fixture_path("plaintext-basic.xlsm");
    for encrypted_name in ["agile-basic.xlsm", "standard-basic.xlsm", "basic-password.xlsm"] {
        let encrypted = fixture_path(encrypted_name);

        let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
            .arg(&plain)
            .arg(&encrypted)
            .arg("--password")
            .arg(PASSWORD)
            .output()
            .unwrap_or_else(|err| panic!("run xlsx-diff ({encrypted_name}): {err}"));

        assert!(
            output.status.success(),
            "{encrypted_name}: expected exit 0\nstdout:\n{}\nstderr:\n{}",
            String::from_utf8_lossy(&output.stdout),
            String::from_utf8_lossy(&output.stderr),
        );

        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(
            stdout.contains("No differences."),
            "{encrypted_name}: expected output to indicate no differences, got:\n{stdout}"
        );
    }
}

#[test]
fn cli_succeeds_with_password_file_stdin() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let mut child = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--password-file")
        .arg("-")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("run xlsx-diff");

    child
        .stdin
        .as_mut()
        .expect("stdin should be piped")
        .write_all(format!("{PASSWORD}\r\n").as_bytes())
        .expect("write password to stdin");

    let output = child.wait_with_output().expect("wait for xlsx-diff");

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
fn cli_original_password_file_overrides_shared_password_file() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let wrong_path = tmp.path().join("wrong.txt");
    let correct_path = tmp.path().join("correct.txt");
    std::fs::write(&wrong_path, "wrong-password\n").expect("write wrong password file");
    std::fs::write(&correct_path, format!("{PASSWORD}\n")).expect("write correct password file");

    // Original is encrypted; modified is plain.
    // Provide an incorrect shared password file, then override the original password file to be correct.
    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&encrypted)
        .arg(&plain)
        .arg("--password-file")
        .arg(&wrong_path)
        .arg("--original-password-file")
        .arg(&correct_path)
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

#[test]
fn cli_modified_password_file_overrides_shared_password_file() {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");

    let tmp = tempfile::tempdir().expect("tempdir");
    let wrong_path = tmp.path().join("wrong.txt");
    let correct_path = tmp.path().join("correct.txt");
    std::fs::write(&wrong_path, "wrong-password\n").expect("write wrong password file");
    std::fs::write(&correct_path, format!("{PASSWORD}\n")).expect("write correct password file");

    // Original is plain; modified is encrypted.
    // Provide an incorrect shared password file, then override the modified password file to be correct.
    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&plain)
        .arg(&encrypted)
        .arg("--password-file")
        .arg(&wrong_path)
        .arg("--modified-password-file")
        .arg(&correct_path)
        .output()
        .expect("run xlsx-diff");

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );
}
