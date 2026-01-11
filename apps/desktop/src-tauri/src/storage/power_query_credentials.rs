use std::path::PathBuf;

use directories::ProjectDirs;
use rand_core::RngCore;
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

use super::encryption::{
    DesktopStorageEncryption, DesktopStorageEncryptionError, KeychainProvider, OsKeychainProvider,
};

const POWER_QUERY_CREDENTIAL_AAD_SCOPE: &str = "formula-desktop-power-query-credentials";
const POWER_QUERY_CREDENTIAL_KEYCHAIN_SERVICE: &str = "formula.desktop";
const POWER_QUERY_CREDENTIAL_KEYCHAIN_ACCOUNT: &str = "power-query-credentials-keyring";

#[derive(Debug, thiserror::Error)]
pub enum PowerQueryCredentialStoreError {
    #[error("could not determine app data directory")]
    NoAppDataDir,
    #[error(transparent)]
    Encryption(#[from] DesktopStorageEncryptionError),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PowerQueryCredentialEntry {
    pub id: String,
    pub secret: JsonValue,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PowerQueryCredentialListEntry {
    pub scope_key: String,
    pub id: String,
}

fn default_credential_store_path() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(
        proj.data_local_dir()
            .join("power-query")
            .join("credentials.json"),
    )
}

fn random_id() -> String {
    let mut bytes = [0u8; 16];
    rand_core::OsRng.fill_bytes(&mut bytes);
    hex::encode(bytes)
}

/// Encrypted credential store for Power Query on desktop.
///
/// The encrypted blob lives in the app's data directory. The encryption keyring
/// material is stored in the OS keychain (macOS Keychain, Windows Credential
/// Manager, etc) via the `keyring` crate.
#[derive(Debug, Clone)]
pub struct PowerQueryCredentialStore<P: KeychainProvider> {
    storage: DesktopStorageEncryption<P>,
}

impl PowerQueryCredentialStore<OsKeychainProvider> {
    pub fn open_default() -> Result<Self, PowerQueryCredentialStoreError> {
        let path = default_credential_store_path().ok_or(PowerQueryCredentialStoreError::NoAppDataDir)?;
        Ok(Self::new(path, OsKeychainProvider))
    }
}

impl<P: KeychainProvider> PowerQueryCredentialStore<P> {
    pub fn new(file_path: PathBuf, keychain: P) -> Self {
        let storage = DesktopStorageEncryption::new(file_path, keychain)
            .with_keychain_namespace(
                POWER_QUERY_CREDENTIAL_KEYCHAIN_SERVICE,
                POWER_QUERY_CREDENTIAL_KEYCHAIN_ACCOUNT,
            )
            .with_aad_scope(POWER_QUERY_CREDENTIAL_AAD_SCOPE);
        Self { storage }
    }

    fn ensure_encrypted(&self) -> Result<(), PowerQueryCredentialStoreError> {
        Ok(self.storage.ensure_encrypted()?)
    }

    pub fn get(&self, scope_key: &str) -> Result<Option<PowerQueryCredentialEntry>, PowerQueryCredentialStoreError> {
        self.ensure_encrypted()?;
        let Some(value) = self.storage.load_document(scope_key)? else {
            return Ok(None);
        };
        Ok(Some(serde_json::from_value(value)?))
    }

    pub fn set(&self, scope_key: &str, secret: JsonValue) -> Result<PowerQueryCredentialEntry, PowerQueryCredentialStoreError> {
        self.ensure_encrypted()?;
        let entry = PowerQueryCredentialEntry {
            id: random_id(),
            secret,
        };
        self.storage
            .save_document(scope_key, serde_json::to_value(&entry)?)?;
        Ok(entry)
    }

    pub fn delete(&self, scope_key: &str) -> Result<(), PowerQueryCredentialStoreError> {
        self.ensure_encrypted()?;
        Ok(self.storage.delete_document(scope_key)?)
    }

    pub fn list(&self) -> Result<Vec<PowerQueryCredentialListEntry>, PowerQueryCredentialStoreError> {
        self.ensure_encrypted()?;
        let mut out = Vec::new();
        for scope_key in self.storage.list_document_ids()? {
            if let Some(value) = self.storage.load_document(&scope_key)? {
                let parsed: PowerQueryCredentialEntry = serde_json::from_value(value)?;
                out.push(PowerQueryCredentialListEntry {
                    scope_key,
                    id: parsed.id,
                });
            }
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

    use serde_json::json;

    use crate::storage::encryption::InMemoryKeychainProvider;

    #[test]
    fn secrets_are_encrypted_at_rest_and_can_be_deleted() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("pq_creds.json");
        let store = PowerQueryCredentialStore::new(file_path.clone(), InMemoryKeychainProvider::default());

        let secret = json!({ "password": "supersecret" });
        let entry = store.set("scope-key", secret.clone()).expect("set");
        assert!(!entry.id.is_empty());
        assert_eq!(entry.secret, secret);

        let on_disk = fs::read_to_string(&file_path).expect("read store file");
        assert!(on_disk.contains("\"encrypted\": true"));
        assert!(
            !on_disk.contains("supersecret"),
            "expected encrypted blob not to contain plaintext secret"
        );

        let loaded = store.get("scope-key").expect("get").expect("present");
        assert_eq!(loaded.id, entry.id);
        assert_eq!(loaded.secret, secret);

        store.delete("scope-key").expect("delete");
        assert!(store.get("scope-key").expect("get after delete").is_none());
    }
}
