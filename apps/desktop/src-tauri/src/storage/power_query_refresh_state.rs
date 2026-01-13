use std::path::PathBuf;

use directories::ProjectDirs;
use serde_json::Value as JsonValue;

use super::encryption::{
    DesktopStorageEncryption, DesktopStorageEncryptionError, KeychainProvider, OsKeychainProvider,
};

const POWER_QUERY_REFRESH_STATE_AAD_SCOPE: &str = "formula-desktop-power-query-refresh-state";
const POWER_QUERY_REFRESH_STATE_KEYCHAIN_SERVICE: &str = "formula.desktop";
const POWER_QUERY_REFRESH_STATE_KEYCHAIN_ACCOUNT: &str = "power-query-refresh-state-keyring";

#[derive(Debug, thiserror::Error)]
pub enum PowerQueryRefreshStateStoreError {
    #[error("could not determine app data directory")]
    NoAppDataDir,
    #[error(transparent)]
    Encryption(#[from] DesktopStorageEncryptionError),
}

fn default_refresh_state_store_path() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(
        proj.data_local_dir()
            .join("power-query")
            .join("refresh_state.json"),
    )
}

/// Encrypted refresh state store for Power Query scheduling metadata on desktop.
///
/// This stores the `RefreshStateStore` JSON payload per workbook id so schedules
/// can survive app restarts without colliding across open documents.
#[derive(Debug, Clone)]
pub struct PowerQueryRefreshStateStore<P: KeychainProvider> {
    storage: DesktopStorageEncryption<P>,
}

impl PowerQueryRefreshStateStore<OsKeychainProvider> {
    pub fn open_default() -> Result<Self, PowerQueryRefreshStateStoreError> {
        let path = default_refresh_state_store_path().ok_or(PowerQueryRefreshStateStoreError::NoAppDataDir)?;
        Ok(Self::new(path, OsKeychainProvider))
    }
}

impl<P: KeychainProvider> PowerQueryRefreshStateStore<P> {
    pub fn new(file_path: PathBuf, keychain: P) -> Self {
        let storage = DesktopStorageEncryption::new(file_path, keychain)
            .with_keychain_namespace(
                POWER_QUERY_REFRESH_STATE_KEYCHAIN_SERVICE,
                POWER_QUERY_REFRESH_STATE_KEYCHAIN_ACCOUNT,
            )
            .with_aad_scope(POWER_QUERY_REFRESH_STATE_AAD_SCOPE);
        Self { storage }
    }

    fn ensure_encrypted(&self) -> Result<(), PowerQueryRefreshStateStoreError> {
        Ok(self.storage.ensure_encrypted()?)
    }

    pub fn load(&self, workbook_id: &str) -> Result<Option<JsonValue>, PowerQueryRefreshStateStoreError> {
        self.ensure_encrypted()?;
        let (value, _recovered) = self
            .storage
            .with_missing_keyring_recovery(|| self.storage.load_document(workbook_id))?;
        Ok(value)
    }

    pub fn save(&self, workbook_id: &str, state: JsonValue) -> Result<(), PowerQueryRefreshStateStoreError> {
        self.ensure_encrypted()?;
        self.storage.with_missing_keyring_recovery(|| {
            self.storage.save_document(workbook_id, state.clone())
        })?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::fs;

    use serde_json::json;

    use crate::storage::encryption::{InMemoryKeychainProvider, KeychainProvider};

    #[test]
    fn refresh_state_is_encrypted_at_rest_and_roundtrips() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("pq_refresh_state.json");
        let store = PowerQueryRefreshStateStore::new(file_path.clone(), InMemoryKeychainProvider::default());

        // Use a long query id to avoid false positives where the ciphertext's base64 encoding
        // happens to contain a short plaintext substring (e.g. "q1").
        let query_id = "power-query-test-id-1234567890";
        let state = json!({
            query_id: { "policy": { "type": "interval", "intervalMs": 123 }, "lastRunAtMs": 456 }
        });
        store.save("workbook-1", state.clone()).expect("save");

        let on_disk = fs::read_to_string(&file_path).expect("read store file");
        assert!(on_disk.contains("\"encrypted\": true"));
        assert!(
            !on_disk.contains("\"intervalMs\""),
            "expected encrypted blob not to contain plaintext schedule"
        );
        assert!(
            !on_disk.contains(query_id),
            "expected encrypted blob not to contain plaintext query ids"
        );

        let loaded = store.load("workbook-1").expect("load").expect("present");
        assert_eq!(loaded, state);
    }

    #[test]
    fn missing_keyring_resets_store_and_allows_resaving() {
        let dir = tempfile::tempdir().expect("tempdir");
        let file_path = dir.path().join("pq_refresh_state.json");

        // Create an encrypted store with an existing keyring.
        let store = PowerQueryRefreshStateStore::new(file_path.clone(), InMemoryKeychainProvider::default());
        store
            .save(
                "workbook-1",
                json!({ "some": { "state": "before-migration-1234567890" } }),
            )
            .expect("initial save");
        assert!(file_path.is_file(), "expected store file to exist");

        // Simulate a missing keychain entry on a new machine.
        let new_keychain = InMemoryKeychainProvider::default();
        let migrated = PowerQueryRefreshStateStore::new(file_path.clone(), new_keychain.clone());

        // Loading should not error; it should reset the store and return None.
        assert!(
            migrated.load("workbook-1").expect("load after migration").is_none(),
            "expected store to be reset and return None"
        );

        // Subsequent saves should succeed and remain encrypted.
        migrated
            .save(
                "workbook-1",
                json!({ "some": { "state": "after-recovery-1234567890" } }),
            )
            .expect("save after recovery");

        let on_disk = fs::read_to_string(&file_path).expect("read store file after save");
        assert!(on_disk.contains("\"encrypted\": true"));
        assert!(
            !on_disk.contains("after-recovery"),
            "expected encrypted blob not to contain plaintext refresh state"
        );

        let loaded = migrated.load("workbook-1").expect("load").expect("present");
        assert_eq!(loaded, json!({ "some": { "state": "after-recovery-1234567890" } }));

        // Ensure the keyring was recreated.
        let stored_keyring = new_keychain
            .get_secret(
                POWER_QUERY_REFRESH_STATE_KEYCHAIN_SERVICE,
                POWER_QUERY_REFRESH_STATE_KEYCHAIN_ACCOUNT,
            )
            .expect("keychain get");
        assert!(stored_keyring.is_some(), "expected keyring to be created during recovery");
    }
}
