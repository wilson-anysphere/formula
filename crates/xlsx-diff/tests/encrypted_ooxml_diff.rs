use std::path::{Path, PathBuf};

use anyhow::{Context, Result};

const PASSWORD: &str = "password";
const UNICODE_PASSWORD: &str = "pässwörd";
const UNICODE_PASSWORD_WITH_EMOJI: &str = "pässwörd🔒";

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

fn xlsb_fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(name)
}

fn assert_diff_empty(
    expected: &Path,
    actual: &Path,
    expected_password: Option<&str>,
    actual_password: Option<&str>,
) -> Result<()> {
    let report = xlsx_diff::diff_workbooks_with_inputs(
        xlsx_diff::DiffInput {
            path: expected,
            password: expected_password,
        },
        xlsx_diff::DiffInput {
            path: actual,
            password: actual_password,
        },
    )
    .with_context(|| {
        format!(
            "diff {} (pw={:?}) vs {} (pw={:?})",
            expected.display(),
            expected_password,
            actual.display(),
            actual_password
        )
    })?;

    assert!(
        report.is_empty(),
        "expected no diffs, got:\n{}",
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
fn diff_agile_fixture_against_plain_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile.xlsx");
    assert_diff_empty(&plain, &encrypted, None, Some(PASSWORD))?;
    Ok(())
}

#[test]
fn diff_standard_fixture_against_plain_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext.xlsx");
    for encrypted in ["standard.xlsx", "standard-4.2.xlsx", "standard-rc4.xlsx"] {
        let encrypted = fixture_path(encrypted);
        assert_diff_empty(&plain, &encrypted, None, Some(PASSWORD))?;
    }
    Ok(())
}

#[test]
fn diff_unicode_password_fixtures_against_plain_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext.xlsx");

    let agile_unicode = fixture_path("agile-unicode.xlsx");
    assert_diff_empty(&plain, &agile_unicode, None, Some(UNICODE_PASSWORD))?;

    let standard_unicode = fixture_path("standard-unicode.xlsx");
    assert_diff_empty(&plain, &standard_unicode, None, Some(UNICODE_PASSWORD_WITH_EMOJI))?;

    let excel_plain = fixture_path("plaintext-excel.xlsx");
    let agile_unicode_excel = fixture_path("agile-unicode-excel.xlsx");
    assert_diff_empty(
        &excel_plain,
        &agile_unicode_excel,
        None,
        Some(UNICODE_PASSWORD_WITH_EMOJI),
    )?;

    Ok(())
}

#[test]
fn diff_xlsm_fixtures_against_plain_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext-basic.xlsm");
    for encrypted in ["agile-basic.xlsm", "standard-basic.xlsm"] {
        let encrypted = fixture_path(encrypted);
        assert_diff_empty(&plain, &encrypted, None, Some(PASSWORD))?;
    }
    Ok(())
}

#[test]
fn diff_encrypted_xlsb_against_plain_no_differences() -> Result<()> {
    let plain = xlsb_fixture_path("simple.xlsb");
    let plain_bytes = std::fs::read(&plain).context("read plaintext xlsb fixture")?;

    // Keep the KDF work factor low so the test stays fast while still exercising the full
    // decryption path inside xlsx-diff.
    let password = "secret";
    let mut opts = formula_office_crypto::EncryptOptions::default();
    opts.spin_count = 1_000;
    let encrypted = formula_office_crypto::encrypt_package_to_ole(&plain_bytes, password, opts)
        .context("encrypt xlsb fixture into OLE EncryptedPackage container")?;

    let tmp = tempfile::tempdir().context("tempdir")?;
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&encrypted_path, &encrypted).context("write encrypted xlsb fixture")?;

    assert_diff_empty(&plain, &encrypted_path, None, Some(password))?;
    Ok(())
}

#[test]
fn diff_large_fixtures_against_plain_large_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext-large.xlsx");

    // `agile-large.xlsx` exercises multi-segment (4096-byte) decryption.
    // Include the Standard fixture as well to cover larger-package CryptoAPI decryption through the
    // `xlsx-diff` tool's end-to-end path.
    for encrypted in ["agile-large.xlsx", "standard-large.xlsx"] {
        let encrypted = fixture_path(encrypted);
        assert_diff_empty(&plain, &encrypted, None, Some(PASSWORD))?;
    }

    Ok(())
}

#[test]
fn agile_empty_password_fixture_requires_explicit_empty_string() -> Result<()> {
    let plain = fixture_path("plaintext.xlsx");
    let encrypted = fixture_path("agile-empty-password.xlsx");

    // Missing password should error (even though the password is the empty string).
    let err = xlsx_diff::diff_workbooks_with_inputs(
        xlsx_diff::DiffInput {
            path: &plain,
            password: None,
        },
        xlsx_diff::DiffInput {
            path: &encrypted,
            password: None,
        },
    )
    .expect_err("expected encrypted input to require an explicit password");
    let msg = err.to_string().to_ascii_lowercase();
    assert!(
        msg.contains("password") || msg.contains("encrypt"),
        "expected error message to mention password/encryption, got: {msg}"
    );

    // An explicit empty password should successfully decrypt.
    assert_diff_empty(&plain, &encrypted, None, Some(""))?;
    Ok(())
}
