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
        Ok(self.storage.load_document(workbook_id)?)
    }

    pub fn save(&self, workbook_id: &str, state: JsonValue) -> Result<(), PowerQueryRefreshStateStoreError> {
        self.ensure_encrypted()?;
        Ok(self.storage.save_document(workbook_id, state)?)
    }
}

