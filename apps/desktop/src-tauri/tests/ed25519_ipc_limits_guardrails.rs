use desktop::ed25519_verifier::LimitedByteVec;

#[test]
fn limited_byte_vec_deserialize_enforces_max_len() {
    for n in 0..=4usize {
        let json = format!(
            "[{}]",
            std::iter::repeat("0")
                .take(n)
                .collect::<Vec<_>>()
                .join(",")
        );
        let parsed: LimitedByteVec<4> =
            serde_json::from_str(&json).expect("expected payload to deserialize");
        assert_eq!(parsed.0.len(), n);
    }

    let err = serde_json::from_str::<LimitedByteVec<4>>("[0,0,0,0,0]")
        .expect_err("expected oversized payload to fail during deserialization");
    let msg = err.to_string();
    assert!(
        msg.contains("max 4 bytes"),
        "expected error to mention max length, got: {msg}"
    );
}

#[test]
fn verify_ed25519_signature_command_has_ipc_origin_checks() {
    let src = include_str!("../src/ed25519_verifier.rs");
    let start = src
        .find("fn verify_ed25519_signature")
        .expect("expected verify_ed25519_signature command to exist");
    let body = &src[start..];
    assert!(
        body.contains("ensure_main_window("),
        "expected verify_ed25519_signature to enforce main-window checks"
    );
    assert!(
        body.contains("ensure_trusted_origin("),
        "expected verify_ed25519_signature to enforce trusted-origin checks"
    );
}

