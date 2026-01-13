use std::io::{Cursor, Write};
use std::path::Path;

use anyhow::{Context, Result};
use ms_offcrypto_writer::Ecma376AgileWriter;

#[test]
fn diff_encrypted_workbook_against_plain_no_differences() -> Result<()> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let plain_bytes =
        std::fs::read(&fixture).with_context(|| format!("read fixture {}", fixture.display()))?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let encrypted_path = tmp.path().join("encrypted.xlsx");
    let password = "xlsx-diff-test-password";

    // Encrypt the fixture bytes into an ECMA-376 agile-encrypted OLE container.
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut rng = rand::rng();
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).context("create encryptor")?;
    writer
        .write_all(&plain_bytes)
        .context("write plaintext workbook bytes")?;
    let cursor = writer.into_inner().context("finalize encryption")?;
    let encrypted_bytes = cursor.into_inner();
    std::fs::write(&encrypted_path, encrypted_bytes)
        .with_context(|| format!("write encrypted workbook {}", encrypted_path.display()))?;

    let report = xlsx_diff::diff_workbooks_with_inputs(
        xlsx_diff::DiffInput {
            path: &fixture,
            password: None,
        },
        xlsx_diff::DiffInput {
            path: &encrypted_path,
            password: Some(password),
        },
    )
        .context("diff workbooks")?;
    assert!(
        report.is_empty(),
        "expected no diffs between plaintext and encrypted-decrypted workbook, got: {:#?}",
        report.differences
    );

    Ok(())
}

#[test]
fn encrypted_workbook_requires_password() -> Result<()> {
    let fixture =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/basic.xlsx");
    let plain_bytes =
        std::fs::read(&fixture).with_context(|| format!("read fixture {}", fixture.display()))?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let encrypted_path = tmp.path().join("encrypted.xlsx");
    let password = "xlsx-diff-test-password";

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut rng = rand::rng();
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).context("create encryptor")?;
    writer
        .write_all(&plain_bytes)
        .context("write plaintext workbook bytes")?;
    let cursor = writer.into_inner().context("finalize encryption")?;
    let encrypted_bytes = cursor.into_inner();
    std::fs::write(&encrypted_path, encrypted_bytes)
        .with_context(|| format!("write encrypted workbook {}", encrypted_path.display()))?;

    let err = xlsx_diff::diff_workbooks_with_inputs(
        xlsx_diff::DiffInput {
            path: &encrypted_path,
            password: None,
        },
        xlsx_diff::DiffInput {
            path: &encrypted_path,
            password: None,
        },
    )
    .expect_err("expected diff to fail without password");
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encrypt"),
        "expected error message to mention password/encryption, got: {msg}"
    );

    Ok(())
}
