use std::path::PathBuf;

use directories::ProjectDirs;
use serde::{Deserialize, Serialize};
use std::time::{SystemTime, UNIX_EPOCH};

use super::encryption::{
    DesktopStorageEncryption, DesktopStorageEncryptionError, KeychainProvider, OsKeychainProvider,
};

const COLLAB_TOKEN_AAD_SCOPE: &str = "formula-desktop-collab-tokens";
const COLLAB_TOKEN_KEYCHAIN_SERVICE: &str = "formula.desktop";
const COLLAB_TOKEN_KEYCHAIN_ACCOUNT: &str = "collab-tokens-keyring";

#[derive(Debug, thiserror::Error)]
pub enum CollabTokenStoreError {
    #[error("could not determine app data directory")]
    NoAppDataDir,
    #[error(transparent)]
    Encryption(#[from] DesktopStorageEncryptionError),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CollabTokenEntry {
    pub token: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub expires_at_ms: Option<i64>,
}

fn default_token_store_path() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(proj.data_local_dir().join("collab").join("tokens.json"))
}

/// Encrypted token store for collaboration (sync server tokens) on desktop.
///
/// The encrypted blob lives in the app's data directory. The encryption keyring
/// material is stored in the OS keychain (macOS Keychain, Windows Credential
/// Manager, etc) via the `keyring` crate.
#[derive(Debug, Clone)]
pub struct CollabTokenStore<P: KeychainProvider> {
    storage: DesktopStorageEncryption<P>,
}

impl CollabTokenStore<OsKeychainProvider> {
    pub fn open_default() -> Result<Self, CollabTokenStoreError> {
        let path = default_token_store_path().ok_or(CollabTokenStoreError::NoAppDataDir)?;
        Ok(Self::new(path, OsKeychainProvider))
    }
}

impl<P: KeychainProvider> CollabTokenStore<P> {
    pub fn new(file_path: PathBuf, keychain: P) -> Self {
        let storage = DesktopStorageEncryption::new(file_path, keychain)
            .with_keychain_namespace(COLLAB_TOKEN_KEYCHAIN_SERVICE, COLLAB_TOKEN_KEYCHAIN_ACCOUNT)
            .with_aad_scope(COLLAB_TOKEN_AAD_SCOPE);
        Self { storage }
    }

    fn ensure_encrypted(&self) -> Result<(), CollabTokenStoreError> {
        Ok(self.storage.ensure_encrypted()?)
    }

    pub fn get(&self, token_key: &str) -> Result<Option<CollabTokenEntry>, CollabTokenStoreError> {
        self.ensure_encrypted()?;
        let Some(value) = self.storage.load_document(token_key)? else {
            return Ok(None);
        };
        let parsed: CollabTokenEntry = serde_json::from_value(value)?;

        if let Some(expires_at_ms) = parsed.expires_at_ms {
            let now_ms = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_millis()
                .min(i64::MAX as u128) as i64;
            if expires_at_ms <= now_ms {
                // Best-effort cleanup: remove expired tokens eagerly.
                let _ = self.storage.delete_document(token_key);
                return Ok(None);
            }
        }

        Ok(Some(parsed))
    }

    pub fn set(
        &self,
        token_key: &str,
        entry: CollabTokenEntry,
    ) -> Result<(), CollabTokenStoreError> {
        self.ensure_encrypted()?;
        self.storage
            .save_document(token_key, serde_json::to_value(&entry)?)?;
        Ok(())
    }

    pub fn delete(&self, token_key: &str) -> Result<(), CollabTokenStoreError> {
        self.ensure_encrypted()?;
        Ok(self.storage.delete_document(token_key)?)
    }

    pub fn file_path(&self) -> &std::path::Path {
        self.storage.file_path()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::fs;

    use crate::storage::encryption::DesktopStorageEncryption;
    use crate::storage::encryption::InMemoryKeychainProvider;

    #[test]
    fn tokens_are_encrypted_at_rest_and_can_be_deleted() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("collab_tokens.json");
        let keychain = InMemoryKeychainProvider::default();
        let store = CollabTokenStore::new(file_path.clone(), keychain.clone());

        let entry = CollabTokenEntry {
            token: "supersecret-token".to_string(),
            // Far-future expiry so the test remains stable over time.
            expires_at_ms: Some(4_000_000_000_000),
        };
        store
            .set("token-key", entry.clone())
            .expect("set token entry");

        let on_disk = fs::read_to_string(&file_path).expect("read store file");
        assert!(on_disk.contains("\"encrypted\": true"));
        assert!(
            !on_disk.contains("supersecret-token"),
            "expected encrypted blob not to contain plaintext token"
        );

        let loaded = store.get("token-key").expect("get").expect("present");
        assert_eq!(loaded.token, entry.token);
        assert_eq!(loaded.expires_at_ms, entry.expires_at_ms);

        store.delete("token-key").expect("delete");
        assert!(store.get("token-key").expect("get after delete").is_none());
    }

    #[test]
    fn get_deletes_expired_tokens() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("collab_tokens.json");
        let keychain = InMemoryKeychainProvider::default();
        let store = CollabTokenStore::new(file_path.clone(), keychain.clone());

        // Store an already-expired token.
        store
            .set(
                "expired-key",
                CollabTokenEntry {
                    token: "expired-token".to_string(),
                    expires_at_ms: Some(0),
                },
            )
            .expect("set expired token");

        // `get` should delete it and return None.
        assert!(store.get("expired-key").expect("get expired").is_none());

        // Verify the underlying encrypted store no longer contains the doc id.
        let storage = DesktopStorageEncryption::new(file_path.clone(), keychain)
            .with_keychain_namespace(COLLAB_TOKEN_KEYCHAIN_SERVICE, COLLAB_TOKEN_KEYCHAIN_ACCOUNT)
            .with_aad_scope(COLLAB_TOKEN_AAD_SCOPE);
        let ids = storage.list_document_ids().expect("list ids");
        assert!(
            !ids.iter().any(|id| id == "expired-key"),
            "expected expired entry to be removed from encrypted store"
        );
    }
}
