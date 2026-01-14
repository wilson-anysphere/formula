//! End-to-end decryption regression test for encrypted macro-enabled workbooks (`.xlsm`).
//!
//! This is gated behind the `encrypted-workbooks` feature because Office decryption support is
//! still landing.
#![cfg(feature = "encrypted-workbooks")]

use std::path::PathBuf;

use formula_io::{open_workbook, open_workbook_with_password, Error, Workbook};

#[cfg(feature = "vba")]
use formula_xlsx::vba::VBAProject;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

const PASSWORD: &str = "password";

#[test]
fn opens_encrypted_xlsm_with_password_and_preserves_vba_project() {
    let path = fixture_path("encrypted/ooxml/basic-password.xlsm");

    // Without a password, we should surface a "password required" error.
    let err =
        open_workbook(&path).expect_err("expected open_workbook to reject encrypted workbook");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );

    let wb =
        open_workbook_with_password(&path, Some(PASSWORD)).expect("open encrypted workbook");
    match wb {
        Workbook::Xlsx(pkg) => {
            let vba_bin = pkg
                .read_part("xl/vbaProject.bin")
                .expect("read xl/vbaProject.bin")
                .expect("expected decrypted .xlsm to preserve xl/vbaProject.bin");
            assert!(!vba_bin.is_empty(), "vbaProject.bin should be non-empty");

            // Optional deeper smoke test: ensure the VBA project is structurally parseable.
            //
            // This is feature-gated because VBA parsing pulls in `formula-vba` (and OpenSSL),
            // which we intentionally keep opt-in for `formula-io`.
            #[cfg(feature = "vba")]
            {
                VBAProject::parse(&vba_bin).expect("parse VBA project");
            }

            // Some producers store signatures in a separate `xl/vbaProjectSignature.bin` part. If
            // present, it must survive decryption.
            if let Some(sig) = pkg
                .read_part("xl/vbaProjectSignature.bin")
                .expect("read xl/vbaProjectSignature.bin")
            {
                assert!(
                    !sig.is_empty(),
                    "expected vbaProjectSignature.bin to be non-empty when present"
                );
            }
        }
        other => panic!("expected Workbook::Xlsx(pkg), got {other:?}"),
    }
}

#[test]
fn wrong_password_yields_invalid_password_error() {
    let path = fixture_path("encrypted/ooxml/basic-password.xlsm");
    let err = open_workbook_with_password(&path, Some("wrong-password"))
        .expect_err("expected invalid password to error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}
