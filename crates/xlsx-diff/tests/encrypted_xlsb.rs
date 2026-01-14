use std::io::{Cursor, Write};
use std::path::Path;

use anyhow::{Context, Result};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};

const PASSWORD: &str = "correct-horse-battery-staple";

fn encrypt_ooxml_with_password(plaintext_zip: &[u8], password: &str) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(Vec::<u8>::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut writer = Ecma376AgileWriter::create(&mut rng, password, &mut cursor)
        .context("create Ecma376AgileWriter")?;
    writer
        .write_all(plaintext_zip)
        .context("write plaintext package")?;
    writer.finalize().context("finalize Ecma376AgileWriter")?;
    Ok(cursor.into_inner())
}

#[test]
fn diff_can_read_password_protected_xlsb() -> Result<()> {
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

    // Without a password, we should error before attempting to parse the decrypted ZIP.
    let err = xlsx_diff::diff_workbooks(&encrypted_path, &plain_path)
        .expect_err("expected encrypted input to require a password");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encrypted"),
        "expected a password/encryption error, got: {msg}"
    );

    // With the correct password, diffing the encrypted container against the underlying package
    // should produce no differences.
    let report = xlsx_diff::diff_workbooks_with_password(&encrypted_path, &plain_path, PASSWORD)?;
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

    // Wrong password should error.
    let err = xlsx_diff::diff_workbooks_with_inputs(
        xlsx_diff::DiffInput {
            path: &encrypted_path,
            password: Some("wrong-password"),
        },
        xlsx_diff::DiffInput::new(&plain_path),
    )
    .expect_err("expected wrong password to error");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("password") || msg.contains("decrypt"),
        "expected a decryption/password error, got: {msg}"
    );

    Ok(())
}
