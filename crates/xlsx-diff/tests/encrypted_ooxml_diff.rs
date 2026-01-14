use std::path::{Path, PathBuf};

use anyhow::{Context, Result};

const PASSWORD: &str = "password";

fn fixture_path(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
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
    for encrypted in ["standard.xlsx", "standard-4.2.xlsx"] {
        let encrypted = fixture_path(encrypted);
        assert_diff_empty(&plain, &encrypted, None, Some(PASSWORD))?;
    }
    Ok(())
}

#[test]
fn diff_large_fixtures_against_plain_large_no_differences() -> Result<()> {
    let plain = fixture_path("plaintext-large.xlsx");

    // `agile-large.xlsx` exercises multi-segment (4096-byte) decryption.
    // Note: `standard-large.xlsx` is covered by `crates/formula-xlsx`’s encrypted fixture tests.
    // `xlsx-diff` decrypts encrypted inputs via `formula-office-crypto` and currently only uses the
    // Agile large fixture here.
    for encrypted in ["agile-large.xlsx"] {
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
