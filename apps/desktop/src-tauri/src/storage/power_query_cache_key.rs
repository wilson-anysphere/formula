use base64::{engine::general_purpose::STANDARD, Engine as _};
use rand_core::RngCore;
use serde::{Deserialize, Serialize};

use super::encryption::{DesktopStorageEncryptionError, KeychainProvider, OsKeychainProvider};

const POWER_QUERY_CACHE_KEY_KEYCHAIN_SERVICE: &str = "formula.desktop";
const POWER_QUERY_CACHE_KEY_KEYCHAIN_ACCOUNT: &str = "power-query-cache-key";
const POWER_QUERY_CACHE_KEY_BYTES: usize = 32;
const POWER_QUERY_CACHE_KEY_VERSION: u32 = 1;

#[derive(Debug, thiserror::Error)]
pub enum PowerQueryCacheKeyStoreError {
    #[error(transparent)]
    Encryption(#[from] DesktopStorageEncryptionError),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
    #[error(transparent)]
    Base64(#[from] base64::DecodeError),
    #[error("invalid key material: {0}")]
    Invalid(String),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PowerQueryCacheKey {
    pub key_version: u32,
    pub key_base64: String,
}

impl PowerQueryCacheKey {
    fn generate() -> Self {
        let mut bytes = [0u8; POWER_QUERY_CACHE_KEY_BYTES];
        rand_core::OsRng.fill_bytes(&mut bytes);
        Self {
            key_version: POWER_QUERY_CACHE_KEY_VERSION,
            key_base64: STANDARD.encode(bytes),
        }
    }

    fn validate(&self) -> Result<(), PowerQueryCacheKeyStoreError> {
        if self.key_version < 1 {
            return Err(PowerQueryCacheKeyStoreError::Invalid(
                "keyVersion must be >= 1".to_string(),
            ));
        }
        let decoded = STANDARD.decode(&self.key_base64)?;
        if decoded.len() != POWER_QUERY_CACHE_KEY_BYTES {
            return Err(PowerQueryCacheKeyStoreError::Invalid(format!(
                "expected {POWER_QUERY_CACHE_KEY_BYTES} bytes, got {}",
                decoded.len()
            )));
        }
        Ok(())
    }
}

/// Keychain-backed AES-256-GCM key for encrypting Power Query cache entries.
///
/// The JavaScript Power Query runtime uses WebCrypto AES-256-GCM for cache
/// encryption; this store persists the raw 32-byte key in the OS keychain so cache
/// entries remain decryptable across app restarts.
#[derive(Debug, Clone)]
pub struct PowerQueryCacheKeyStore<P: KeychainProvider> {
    keychain: P,
}

impl PowerQueryCacheKeyStore<OsKeychainProvider> {
    pub fn open_default() -> Self {
        Self::new(OsKeychainProvider)
    }
}

impl<P: KeychainProvider> PowerQueryCacheKeyStore<P> {
    pub fn new(keychain: P) -> Self {
        Self { keychain }
    }

    pub fn get_or_create(&self) -> Result<PowerQueryCacheKey, PowerQueryCacheKeyStoreError> {
        let secret = self
            .keychain
            .get_secret(
                POWER_QUERY_CACHE_KEY_KEYCHAIN_SERVICE,
                POWER_QUERY_CACHE_KEY_KEYCHAIN_ACCOUNT,
            )?
            .and_then(|bytes| serde_json::from_slice::<PowerQueryCacheKey>(&bytes).ok());

        if let Some(existing) = secret {
            if existing.validate().is_ok() {
                return Ok(existing);
            }
        }

        let created = PowerQueryCacheKey::generate();
        created.validate()?;
        let bytes = serde_json::to_vec(&created)?;
        self.keychain.set_secret(
            POWER_QUERY_CACHE_KEY_KEYCHAIN_SERVICE,
            POWER_QUERY_CACHE_KEY_KEYCHAIN_ACCOUNT,
            &bytes,
        )?;
        Ok(created)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use crate::storage::encryption::InMemoryKeychainProvider;

    #[test]
    fn key_is_stored_in_keychain_and_is_stable() {
        let store = PowerQueryCacheKeyStore::new(InMemoryKeychainProvider::default());
        let first = store.get_or_create().expect("first key");
        assert_eq!(first.key_version, 1);
        assert_eq!(STANDARD.decode(&first.key_base64).unwrap().len(), POWER_QUERY_CACHE_KEY_BYTES);

        let second = store.get_or_create().expect("second key");
        assert_eq!(second.key_version, 1);
        assert_eq!(second.key_base64, first.key_base64);
    }
}

