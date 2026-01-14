use std::io::Write;
use std::path::Path;
use std::process::{Command, Stdio};
use std::sync::OnceLock;

use anyhow::{Context, Result};
use formula_office_crypto::EncryptOptions;

const PASSWORD: &str = "correct-horse-battery-staple";
const FAST_TEST_SPIN_COUNT: u32 = 1_000;

fn fixture_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Vec<u8>> = OnceLock::new();
    BYTES
        .get_or_init(|| {
            let fixture_path = Path::new(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../formula-xlsb/tests/fixtures/simple.xlsb"
            ));
            std::fs::read(fixture_path).expect("read base xlsb fixture")
        })
        .as_slice()
}

fn encrypted_fixture_bytes() -> &'static [u8] {
    static BYTES: OnceLock<Vec<u8>> = OnceLock::new();
    BYTES
        .get_or_init(|| encrypt_ooxml_with_password(fixture_bytes(), PASSWORD).expect("encrypt fixture"))
        .as_slice()
}

fn encrypt_ooxml_with_password(plaintext_zip: &[u8], password: &str) -> Result<Vec<u8>> {
    formula_office_crypto::encrypt_package_to_ole(
        plaintext_zip,
        password,
        EncryptOptions {
            spin_count: FAST_TEST_SPIN_COUNT,
            ..Default::default()
        },
    )
    .context("encrypt OOXML package to EncryptedPackage OLE")
}

#[test]
fn library_can_diff_password_protected_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    // With the correct password, diffing the encrypted container against the underlying package
    // should produce no differences.
    let report = xlsb_diff::diff_workbooks_with_inputs(
        xlsb_diff::DiffInput {
            path: &encrypted_path,
            password: Some(PASSWORD),
        },
        xlsb_diff::DiffInput::new(&plain_path),
    )?;
    assert!(
        report.is_empty(),
        "expected no diffs between encrypted and plain fixture, got:\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<Vec<_>>()
            .join("\n")
    );

    Ok(())
}

#[test]
fn library_errors_on_wrong_password_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let err = xlsb_diff::diff_workbooks_with_inputs(
        xlsb_diff::DiffInput {
            path: &encrypted_path,
            password: Some("wrong-password"),
        },
        xlsb_diff::DiffInput::new(&plain_path),
    )
    .expect_err("expected wrong password to error");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("password") || msg.contains("decrypt"),
        "expected a decryption/password error, got: {msg}"
    );

    Ok(())
}

#[test]
fn cli_supports_password_file_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    let password_path = tmp.path().join("password.txt");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;
    std::fs::write(&password_path, format!("{PASSWORD}\n")).context("write password file")?;

    let output = Command::new(env!("CARGO_BIN_EXE_xlsb_diff"))
        .arg(&plain_path)
        .arg(&encrypted_path)
        .arg("--password-file")
        .arg(&password_path)
        .output()
        .context("run xlsb-diff")?;

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

    Ok(())
}

#[test]
fn cli_supports_password_file_stdin_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let mut child = Command::new(env!("CARGO_BIN_EXE_xlsb_diff"))
        .arg(&plain_path)
        .arg(&encrypted_path)
        .arg("--password-file")
        .arg("-")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("spawn xlsb-diff")?;

    child
        .stdin
        .as_mut()
        .expect("stdin should be piped")
        .write_all(format!("{PASSWORD}\n").as_bytes())
        .context("write password to stdin")?;

    let output = child.wait_with_output().context("wait for xlsb-diff")?;

    assert!(
        output.status.success(),
        "expected exit 0\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );

    Ok(())
}

#[test]
fn cli_succeeds_with_password_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let output = Command::new(env!("CARGO_BIN_EXE_xlsb_diff"))
        .arg(&plain_path)
        .arg(&encrypted_path)
        .arg("--password")
        .arg(PASSWORD)
        .output()
        .context("run xlsb-diff")?;

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

    Ok(())
}

#[test]
fn cli_errors_without_password_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let output = Command::new(env!("CARGO_BIN_EXE_xlsb_diff"))
        .arg(&plain_path)
        .arg(&encrypted_path)
        .output()
        .context("run xlsb-diff")?;

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

    Ok(())
}

#[test]
fn cli_errors_with_wrong_password_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let output = Command::new(env!("CARGO_BIN_EXE_xlsb_diff"))
        .arg(&plain_path)
        .arg(&encrypted_path)
        .arg("--password")
        .arg("wrong-password")
        .output()
        .context("run xlsb-diff")?;

    assert!(
        !output.status.success(),
        "expected non-zero exit status when password is incorrect\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr),
    );

    let stderr = String::from_utf8_lossy(&output.stderr).to_ascii_lowercase();
    assert!(
        stderr.contains("password") || stderr.contains("decrypt"),
        "expected error message to mention password/decryption, got:\n{stderr}"
    );

    Ok(())
}

#[test]
fn library_errors_without_password_for_encrypted_xlsb() -> Result<()> {
    let plaintext = fixture_bytes();
    let encrypted = encrypted_fixture_bytes();

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, encrypted).context("write encrypted xlsb")?;

    let err = xlsb_diff::diff_workbooks(&encrypted_path, &plain_path)
        .expect_err("expected encrypted input to require a password");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encrypted"),
        "expected error message to mention password/encryption, got: {msg}"
    );

    Ok(())
}
