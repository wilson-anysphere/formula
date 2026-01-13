use base64::engine::general_purpose;
use base64::Engine as _;
use desktop::storage::encryption::{DesktopStorageEncryptionError, KeyRing, KeyRingJson};
use serde::Deserialize;
use std::collections::BTreeMap;
use std::path::Path;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct Fixture {
    algorithm: String,
    key: String,
    iv: String,
    plaintext: PlaintextFixture,
    aad_context: serde_json::Value,
    expected: ExpectedFixture,
}

#[derive(Debug, Deserialize)]
struct PlaintextFixture {
    encoding: String,
    value: String,
}

#[derive(Debug, Deserialize)]
struct ExpectedFixture {
    ciphertext: String,
    tag: String,
}

fn load_fixture() -> Fixture {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../../fixtures/crypto/desktop-storage-encryption-v1.json"
    );
    let raw = std::fs::read_to_string(Path::new(fixture_path)).expect("read desktop storage encryption fixture");
    serde_json::from_str(&raw).expect("parse desktop storage encryption fixture JSON")
}

fn decode_fixed<const N: usize>(value_b64: &str, field: &str) -> [u8; N] {
    let decoded = general_purpose::STANDARD
        .decode(value_b64)
        .unwrap_or_else(|err| panic!("failed to decode {field} as base64: {err}"));
    let decoded_len = decoded.len();
    decoded
        .try_into()
        .unwrap_or_else(|_| panic!("{field} must decode to {N} bytes (got {decoded_len})"))
}

fn plaintext_bytes(plain: &PlaintextFixture) -> Vec<u8> {
    match plain.encoding.as_str() {
        "utf8" => plain.value.as_bytes().to_vec(),
        "base64" => general_purpose::STANDARD
            .decode(&plain.value)
            .unwrap_or_else(|err| panic!("failed to decode plaintext as base64: {err}")),
        other => panic!("unsupported plaintext encoding: {other}"),
    }
}

#[test]
fn desktop_storage_encryption_vectors_v1_match_js() {
    let fixture = load_fixture();

    let key_version = 1;
    let _key = decode_fixed::<32>(&fixture.key, "key");
    let iv = decode_fixed::<12>(&fixture.iv, "iv");

    let mut keys = BTreeMap::new();
    keys.insert(key_version.to_string(), fixture.key.clone());
    let keyring = KeyRing::from_json(KeyRingJson {
        current_version: key_version,
        keys,
    })
    .expect("construct keyring from fixture key");

    let plaintext = plaintext_bytes(&fixture.plaintext);
    let encrypted = keyring
        .encrypt_with_iv(&plaintext, Some(&fixture.aad_context), Some(iv))
        .expect("encrypt fixture plaintext");

    assert_eq!(encrypted.key_version, key_version);
    assert_eq!(encrypted.algorithm, fixture.algorithm);
    assert_eq!(encrypted.iv, fixture.iv);
    assert_eq!(encrypted.ciphertext, fixture.expected.ciphertext);
    assert_eq!(encrypted.tag, fixture.expected.tag);

    let decrypted = keyring
        .decrypt(&encrypted, Some(&fixture.aad_context))
        .expect("decrypt fixture ciphertext");
    assert_eq!(decrypted, plaintext);

    // Ensure we didn't accidentally rely on base64 string comparisons that hide
    // byte-level differences.
    let ciphertext_bytes = general_purpose::STANDARD
        .decode(&encrypted.ciphertext)
        .expect("decode ciphertext");
    let tag_bytes = general_purpose::STANDARD.decode(&encrypted.tag).expect("decode tag");
    assert_eq!(
        general_purpose::STANDARD.encode(ciphertext_bytes),
        fixture.expected.ciphertext
    );
    assert_eq!(general_purpose::STANDARD.encode(tag_bytes), fixture.expected.tag);
}

#[test]
fn desktop_storage_fixture_decrypt_roundtrip_rejects_wrong_aad() {
    let fixture = load_fixture();
    let key_version = 1;
    let iv = decode_fixed::<12>(&fixture.iv, "iv");

    let mut keys = BTreeMap::new();
    keys.insert(key_version.to_string(), fixture.key.clone());
    let keyring = KeyRing::from_json(KeyRingJson {
        current_version: key_version,
        keys,
    })
    .expect("construct keyring from fixture key");

    let plaintext = plaintext_bytes(&fixture.plaintext);
    let encrypted = keyring
        .encrypt_with_iv(&plaintext, Some(&fixture.aad_context), Some(iv))
        .expect("encrypt fixture plaintext");

    let mut wrong_context = fixture.aad_context.clone();
    if let Some(obj) = wrong_context.as_object_mut() {
        obj.insert("scope".to_string(), serde_json::Value::String("wrong-scope".to_string()));
    }

    let err = keyring
        .decrypt(&encrypted, Some(&wrong_context))
        .expect_err("decrypt with wrong AAD should fail");
    assert!(matches!(err, DesktopStorageEncryptionError::Crypto));
}
