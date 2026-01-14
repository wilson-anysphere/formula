use std::io::{Cursor, Write};
use std::path::Path;
use std::process::Command;

use anyhow::{Context, Result};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};

const PASSWORD: &str = "correct-horse-battery-staple";

fn encrypt_ooxml_with_password(plaintext_zip: &[u8], password: &str) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(Vec::<u8>::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, &mut cursor).context("create writer")?;
    writer
        .write_all(plaintext_zip)
        .context("write plaintext package")?;
    writer.finalize().context("finalize writer")?;
    Ok(cursor.into_inner())
}

#[test]
fn library_can_diff_password_protected_xlsb() -> Result<()> {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let plaintext = std::fs::read(fixture_path).context("read base xlsb fixture")?;
    let encrypted = encrypt_ooxml_with_password(&plaintext, PASSWORD)?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, &plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, &encrypted).context("write encrypted xlsb")?;

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
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let plaintext = std::fs::read(fixture_path).context("read base xlsb fixture")?;
    let encrypted = encrypt_ooxml_with_password(&plaintext, PASSWORD)?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&plain_path, &plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, &encrypted).context("write encrypted xlsb")?;

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
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    let plaintext = std::fs::read(fixture_path).context("read base xlsb fixture")?;
    let encrypted = encrypt_ooxml_with_password(&plaintext, PASSWORD)?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let plain_path = tmp.path().join("plain.xlsb");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    let password_path = tmp.path().join("password.txt");
    std::fs::write(&plain_path, &plaintext).context("write plain xlsb")?;
    std::fs::write(&encrypted_path, &encrypted).context("write encrypted xlsb")?;
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

