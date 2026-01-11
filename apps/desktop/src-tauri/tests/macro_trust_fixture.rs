use std::path::Path;

use formula_desktop_tauri::commands::{evaluate_macro_trust, MacroBlockedReason, MacroSignatureStatus};
use formula_desktop_tauri::file_io::read_xlsx_blocking;
use formula_desktop_tauri::macro_trust::{MacroTrustDecision, MacroTrustStore};

#[test]
fn xlsm_fixture_is_blocked_by_default() {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../../fixtures/xlsx/macros/basic.xlsm"
    );
    let workbook = read_xlsx_blocking(Path::new(fixture_path)).expect("read fixture workbook");
    assert!(
        workbook.vba_project_bin.is_some(),
        "fixture should contain vbaProject.bin"
    );
    let fingerprint = workbook
        .macro_fingerprint
        .as_deref()
        .expect("fingerprint computed for macro-enabled workbook");

    let store = MacroTrustStore::new_ephemeral();
    assert_eq!(store.trust_state(fingerprint), MacroTrustDecision::Blocked);
}

#[test]
fn trusted_signed_only_requires_cryptographically_verified_signature() {
    assert!(
        evaluate_macro_trust(MacroTrustDecision::TrustedSignedOnly, MacroSignatureStatus::SignedVerified).is_ok()
    );

    for status in [
        MacroSignatureStatus::Unsigned,
        MacroSignatureStatus::SignedInvalid,
        MacroSignatureStatus::SignedParseError,
        MacroSignatureStatus::SignedUnverified,
    ] {
        assert_eq!(
            evaluate_macro_trust(MacroTrustDecision::TrustedSignedOnly, status),
            Err(MacroBlockedReason::SignatureRequired),
            "expected {status:?} to be blocked for TrustedSignedOnly"
        );
    }
}
