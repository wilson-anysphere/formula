use std::fs;

use desktop::storage::encryption::{
    DesktopStorageEncryption, InMemoryKeychainProvider, KeychainProvider,
};
use serde_json::json;

#[test]
fn enable_encryption_ciphertext_on_disk_then_disable_restores_plaintext() {
    let dir = tempfile::tempdir().unwrap();
    let file_path = dir.path().join("store.json");
    let keychain = InMemoryKeychainProvider::default();

    let store = DesktopStorageEncryption::new(file_path.clone(), keychain.clone());

    store.enable_encryption().unwrap();
    store
        .save_document("doc-1", json!({ "name": "Secret Document", "value": 42 }))
        .unwrap();

    let encrypted_raw = fs::read_to_string(&file_path).unwrap();
    assert!(
        !encrypted_raw.contains("Secret Document"),
        "expected ciphertext-only on disk"
    );
    assert!(
        encrypted_raw.contains("\"ciphertext\""),
        "expected encrypted payload fields on disk"
    );

    let store_reload = DesktopStorageEncryption::new(file_path.clone(), keychain.clone());
    let loaded = store_reload.load_document("doc-1").unwrap().unwrap();
    assert_eq!(loaded, json!({ "name": "Secret Document", "value": 42 }));

    store_reload.disable_encryption(true).unwrap();

    let secret = keychain
        .get_secret("formula.desktop", "storage-keyring")
        .unwrap();
    assert!(secret.is_none(), "expected keyring to be removed from keychain");

    let plaintext_raw = fs::read_to_string(&file_path).unwrap();
    assert!(
        plaintext_raw.contains("Secret Document"),
        "expected plaintext after disabling encryption"
    );

    let store_plain = DesktopStorageEncryption::new(file_path.clone(), keychain);
    let loaded_plain = store_plain.load_document("doc-1").unwrap().unwrap();
    assert_eq!(loaded_plain, json!({ "name": "Secret Document", "value": 42 }));
}

#[test]
fn rotate_key_reencrypts_and_old_keyring_cannot_decrypt_new_store() {
    let dir = tempfile::tempdir().unwrap();
    let file_path = dir.path().join("store.json");
    let keychain = InMemoryKeychainProvider::default();

    let store = DesktopStorageEncryption::new(file_path.clone(), keychain.clone());
    store.enable_encryption().unwrap();
    store.save_document("doc-1", json!({ "v": 1 })).unwrap();

    let old_keyring_bytes = keychain
        .get_secret("formula.desktop", "storage-keyring")
        .unwrap()
        .expect("keyring missing");
    let old_keyring =
        desktop::storage::encryption::KeyRing::from_bytes(&old_keyring_bytes).unwrap();

    let raw_before = fs::read_to_string(&file_path).unwrap();
    let before: serde_json::Value = serde_json::from_str(&raw_before).unwrap();
    let key_version_before = before.get("keyVersion").and_then(|v| v.as_u64()).unwrap();

    let next_version = store.rotate_key().unwrap();
    assert_eq!(next_version as u64, key_version_before + 1);

    let raw_after = fs::read_to_string(&file_path).unwrap();
    let after: serde_json::Value = serde_json::from_str(&raw_after).unwrap();
    let key_version_after = after.get("keyVersion").and_then(|v| v.as_u64()).unwrap();
    assert_eq!(key_version_after, key_version_before + 1);

    // Ensure the *old* keyring alone cannot decrypt the rotated store.
    let envelope: desktop::storage::encryption::EncryptedEnvelope =
        serde_json::from_value(after.clone()).unwrap();
    let aad = serde_json::json!({ "scope": "formula-desktop-store", "schemaVersion": 1 });
    let decrypt_err = old_keyring.decrypt(&envelope, Some(&aad)).unwrap_err();
    assert!(matches!(
        decrypt_err,
        desktop::storage::encryption::DesktopStorageEncryptionError::MissingKeyVersion(_)
    ));

    // The store (with the rotated keyring in the keychain) can still load docs.
    let loaded = store.load_document("doc-1").unwrap().unwrap();
    assert_eq!(loaded, json!({ "v": 1 }));
}
