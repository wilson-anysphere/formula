use std::path::PathBuf;

use base64::{engine::general_purpose::STANDARD, Engine as _};
use directories::ProjectDirs;
use serde::{Deserialize, Serialize};

use super::encryption::{
    DesktopStorageEncryption, DesktopStorageEncryptionError, KeychainProvider, OsKeychainProvider,
};

const COLLAB_ENCRYPTION_KEYS_AAD_SCOPE: &str = "formula-desktop-collab-encryption-keys";
const COLLAB_ENCRYPTION_KEYS_KEYCHAIN_SERVICE: &str = "formula.desktop";
const COLLAB_ENCRYPTION_KEYS_KEYCHAIN_ACCOUNT: &str = "collab-encryption-keys-keyring";

// We currently only support AES-256 keys for cell encryption.
const CELL_ENCRYPTION_KEY_BYTES: usize = 32;

#[derive(Debug, thiserror::Error)]
pub enum CollabEncryptionKeyStoreError {
    #[error("could not determine app data directory")]
    NoAppDataDir,
    #[error(transparent)]
    Encryption(#[from] DesktopStorageEncryptionError),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
    #[error(transparent)]
    Base64(#[from] base64::DecodeError),
    #[error("invalid key material")]
    InvalidKeyMaterial,
    #[error("invalid input: {0}")]
    InvalidInput(String),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CollabEncryptionKeyEntry {
    pub key_id: String,
    pub key_bytes_base64: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CollabEncryptionKeyListEntry {
    pub key_id: String,
}

fn default_collab_key_store_path() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(
        proj.data_local_dir()
            .join("collab")
            .join("encryption-keys.json"),
    )
}

fn scope_key_prefix(doc_id: &str) -> Result<String, CollabEncryptionKeyStoreError> {
    let doc_id = doc_id.trim();
    if doc_id.is_empty() {
        return Err(CollabEncryptionKeyStoreError::InvalidInput(
            "docId must be a non-empty string".to_string(),
        ));
    }
    Ok(format!("collab-enc:{doc_id}:"))
}

fn scope_key(doc_id: &str, key_id: &str) -> Result<String, CollabEncryptionKeyStoreError> {
    let key_id = key_id.trim();
    if key_id.is_empty() {
        return Err(CollabEncryptionKeyStoreError::InvalidInput(
            "keyId must be a non-empty string".to_string(),
        ));
    }
    Ok(format!("{}{}", scope_key_prefix(doc_id)?, key_id))
}

fn normalize_key_bytes_base64(value: &str) -> Result<String, CollabEncryptionKeyStoreError> {
    let trimmed = value.trim();
    let decoded = STANDARD.decode(trimmed)?;
    if decoded.len() != CELL_ENCRYPTION_KEY_BYTES {
        return Err(CollabEncryptionKeyStoreError::InvalidKeyMaterial);
    }
    Ok(STANDARD.encode(decoded))
}

/// OS-keychain-backed encrypted store for collaborative cell encryption keys.
///
/// Each entry is stored under a scope key: `collab-enc:${docId}:${keyId}` so we can list and
/// delete all keys for a given document.
#[derive(Debug, Clone)]
pub struct CollabEncryptionKeyStore<P: KeychainProvider> {
    storage: DesktopStorageEncryption<P>,
}

impl CollabEncryptionKeyStore<OsKeychainProvider> {
    pub fn open_default() -> Result<Self, CollabEncryptionKeyStoreError> {
        let path = default_collab_key_store_path().ok_or(CollabEncryptionKeyStoreError::NoAppDataDir)?;
        Ok(Self::new(path, OsKeychainProvider))
    }
}

impl<P: KeychainProvider> CollabEncryptionKeyStore<P> {
    pub fn new(file_path: PathBuf, keychain: P) -> Self {
        let storage = DesktopStorageEncryption::new(file_path, keychain)
            .with_keychain_namespace(
                COLLAB_ENCRYPTION_KEYS_KEYCHAIN_SERVICE,
                COLLAB_ENCRYPTION_KEYS_KEYCHAIN_ACCOUNT,
            )
            .with_aad_scope(COLLAB_ENCRYPTION_KEYS_AAD_SCOPE);
        Self { storage }
    }

    fn ensure_encrypted(&self) -> Result<(), CollabEncryptionKeyStoreError> {
        Ok(self.storage.ensure_encrypted()?)
    }

    pub fn get(
        &self,
        doc_id: &str,
        key_id: &str,
    ) -> Result<Option<CollabEncryptionKeyEntry>, CollabEncryptionKeyStoreError> {
        self.ensure_encrypted()?;
        let scope_key = scope_key(doc_id, key_id)?;
        let Some(value) = self.storage.load_document(&scope_key)? else {
            return Ok(None);
        };
        Ok(Some(serde_json::from_value(value)?))
    }

    pub fn set(
        &self,
        doc_id: &str,
        key_id: &str,
        key_bytes_base64: &str,
    ) -> Result<CollabEncryptionKeyListEntry, CollabEncryptionKeyStoreError> {
        self.ensure_encrypted()?;
        let scope_key = scope_key(doc_id, key_id)?;
        let normalized = normalize_key_bytes_base64(key_bytes_base64)?;

        let entry = CollabEncryptionKeyEntry {
            key_id: key_id.trim().to_string(),
            key_bytes_base64: normalized,
        };

        self.storage
            .save_document(&scope_key, serde_json::to_value(&entry)?)?;

        Ok(CollabEncryptionKeyListEntry {
            key_id: entry.key_id,
        })
    }

    pub fn delete(&self, doc_id: &str, key_id: &str) -> Result<(), CollabEncryptionKeyStoreError> {
        self.ensure_encrypted()?;
        let scope_key = scope_key(doc_id, key_id)?;
        Ok(self.storage.delete_document(&scope_key)?)
    }

    pub fn list(&self, doc_id: &str) -> Result<Vec<CollabEncryptionKeyListEntry>, CollabEncryptionKeyStoreError> {
        self.ensure_encrypted()?;
        let prefix = scope_key_prefix(doc_id)?;
        let mut out = Vec::new();
        for scope_key in self.storage.list_document_ids()? {
            let Some(key_id) = scope_key.strip_prefix(&prefix) else {
                continue;
            };
            if key_id.trim().is_empty() {
                continue;
            }
            out.push(CollabEncryptionKeyListEntry {
                key_id: key_id.to_string(),
            });
        }
        Ok(out)
    }

    pub fn file_path(&self) -> &std::path::Path {
        self.storage.file_path()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::fs;

    use crate::storage::encryption::InMemoryKeychainProvider;

    #[test]
    fn keys_are_encrypted_at_rest_and_can_be_deleted() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("collab_keys.json");
        let store = CollabEncryptionKeyStore::new(file_path.clone(), InMemoryKeychainProvider::default());

        let key_bytes = [42u8; CELL_ENCRYPTION_KEY_BYTES];
        let key_b64 = STANDARD.encode(key_bytes);

        let set = store
            .set("doc-1", "key-1", &key_b64)
            .expect("set");
        assert_eq!(set.key_id, "key-1");

        let on_disk = fs::read_to_string(&file_path).expect("read store file");
        assert!(on_disk.contains("\"encrypted\": true"));
        assert!(
            !on_disk.contains("\"keyBytesBase64\":"),
            "expected encrypted blob not to contain plaintext key material fields"
        );

        let loaded = store.get("doc-1", "key-1").expect("get").expect("present");
        assert_eq!(loaded.key_id, "key-1");
        assert_eq!(loaded.key_bytes_base64, key_b64);

        store.delete("doc-1", "key-1").expect("delete");
        assert!(store.get("doc-1", "key-1").expect("get after delete").is_none());
    }
}
