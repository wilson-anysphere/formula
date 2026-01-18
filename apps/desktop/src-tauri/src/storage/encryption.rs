//! Desktop storage encryption-at-rest (Tauri/Rust).
//!
//! This module provides a production-grade, testable implementation of the
//! desktop storage encryption flow used by the JavaScript reference store in
//! `encryptedDocumentStore.js`.
//!
//! ## Encryption format (stable)
//!
//! Encrypted payloads are stored as a JSON "envelope" with the following fields
//! (camelCase), designed to match `packages/security/crypto`:
//!
//! ```json
//! {
//!   "keyVersion": 1,
//!   "algorithm": "aes-256-gcm",
//!   "iv": "<base64(12 bytes)>",
//!   "ciphertext": "<base64>",
//!   "tag": "<base64(16 bytes)>",
//!   "createdAt": "optional RFC3339 timestamp"
//! }
//! ```
//!
//! The store file on disk is either plaintext:
//!
//! ```json
//! { "schemaVersion": 1, "encrypted": false, "documents": { "doc-1": { ... } } }
//! ```
//!
//! …or encrypted:
//!
//! ```json
//! { "schemaVersion": 1, "encrypted": true,  ...<envelope fields> }
//! ```
//!
//! ## AAD / encryption context
//!
//! We bind ciphertexts to a deterministic AAD (additional authenticated data)
//! context that includes:
//! - a stable scope string (default `"formula-desktop-store"`)
//! - the store schema version
//! - an optional document id
//!
//! The AAD bytes are a deterministic JSON encoding matching the JS helper
//! `aadFromContext()` (`packages/security/crypto/utils.js`): recursively sort
//! object keys, then `JSON.stringify`, then UTF-8 bytes.

use std::collections::{BTreeMap, HashMap};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use crate::atomic_write::write_file_atomic_io;
use aes_gcm::aead::{AeadInPlace, KeyInit};
use aes_gcm::{Aes256Gcm, Nonce};
use base64::engine::general_purpose;
use base64::Engine as _;
use rand_core::RngCore;
use serde::{Deserialize, Serialize};

const AES_256_GCM: &str = "aes-256-gcm";
const AES_GCM_IV_BYTES: usize = 12;
const AES_GCM_TAG_BYTES: usize = 16;
const AES_256_KEY_BYTES: usize = 32;

pub const DEFAULT_AAD_SCOPE: &str = "formula-desktop-store";
pub const DEFAULT_KEYCHAIN_SERVICE: &str = "formula.desktop";
pub const DEFAULT_KEYCHAIN_ACCOUNT: &str = "storage-keyring";
pub const DEFAULT_STORE_SCHEMA_VERSION: u32 = 1;

#[derive(Debug, thiserror::Error)]
pub enum DesktopStorageEncryptionError {
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("base64 error: {0}")]
    Base64(#[from] base64::DecodeError),
    #[error("invalid encrypted payload: {0}")]
    InvalidPayload(String),
    #[error("unsupported algorithm: {0}")]
    UnsupportedAlgorithm(String),
    #[error("crypto operation failed")]
    Crypto,
    #[error("encrypted store present but no keyring in keychain")]
    MissingKeyRing,
    #[error("cannot rotate key: store is not encrypted")]
    StoreNotEncrypted,
    #[error("missing key material for version {0}")]
    MissingKeyVersion(u32),
    #[error("keychain error: {0}")]
    Keychain(String),
}

pub trait KeychainProvider: Send + Sync {
    fn get_secret(&self, service: &str, account: &str) -> Result<Option<Vec<u8>>, DesktopStorageEncryptionError>;

    fn set_secret(
        &self,
        service: &str,
        account: &str,
        secret: &[u8],
    ) -> Result<(), DesktopStorageEncryptionError>;

    fn delete_secret(&self, service: &str, account: &str) -> Result<(), DesktopStorageEncryptionError>;
}

/// OS-backed keychain provider (macOS Keychain, Windows Credential Manager,
/// Linux Secret Service) via the `keyring` crate.
///
/// This is intentionally thin; we store the serialized keyring JSON as UTF-8.
#[derive(Debug, Default, Clone, Copy)]
pub struct OsKeychainProvider;

impl KeychainProvider for OsKeychainProvider {
    fn get_secret(
        &self,
        service: &str,
        account: &str,
    ) -> Result<Option<Vec<u8>>, DesktopStorageEncryptionError> {
        let entry = keyring::Entry::new(service, account)
            .map_err(|err| DesktopStorageEncryptionError::Keychain(err.to_string()))?;

        match entry.get_secret() {
            Ok(secret) => Ok(Some(secret)),
            Err(keyring::Error::NoEntry) => Ok(None),
            Err(err) => Err(DesktopStorageEncryptionError::Keychain(err.to_string())),
        }
    }

    fn set_secret(
        &self,
        service: &str,
        account: &str,
        secret: &[u8],
    ) -> Result<(), DesktopStorageEncryptionError> {
        let entry = keyring::Entry::new(service, account)
            .map_err(|err| DesktopStorageEncryptionError::Keychain(err.to_string()))?;

        entry
            .set_secret(secret)
            .map_err(|err| DesktopStorageEncryptionError::Keychain(err.to_string()))?;
        Ok(())
    }

    fn delete_secret(&self, service: &str, account: &str) -> Result<(), DesktopStorageEncryptionError> {
        let entry = keyring::Entry::new(service, account)
            .map_err(|err| DesktopStorageEncryptionError::Keychain(err.to_string()))?;

        match entry.delete_credential() {
            Ok(()) => Ok(()),
            Err(keyring::Error::NoEntry) => Ok(()),
            Err(err) => Err(DesktopStorageEncryptionError::Keychain(err.to_string())),
        }
    }
}

/// In-memory keychain provider intended for tests / CI.
#[derive(Debug, Default, Clone)]
pub struct InMemoryKeychainProvider {
    inner: Arc<Mutex<HashMap<(String, String), Vec<u8>>>>,
}

impl KeychainProvider for InMemoryKeychainProvider {
    fn get_secret(
        &self,
        service: &str,
        account: &str,
    ) -> Result<Option<Vec<u8>>, DesktopStorageEncryptionError> {
        let guard = self.inner.lock().map_err(|_| {
            DesktopStorageEncryptionError::Keychain("keychain mutex is poisoned".to_string())
        })?;
        Ok(guard.get(&(service.to_string(), account.to_string())).cloned())
    }

    fn set_secret(
        &self,
        service: &str,
        account: &str,
        secret: &[u8],
    ) -> Result<(), DesktopStorageEncryptionError> {
        let mut guard = self.inner.lock().map_err(|_| {
            DesktopStorageEncryptionError::Keychain("keychain mutex is poisoned".to_string())
        })?;
        guard.insert((service.to_string(), account.to_string()), secret.to_vec());
        Ok(())
    }

    fn delete_secret(&self, service: &str, account: &str) -> Result<(), DesktopStorageEncryptionError> {
        let mut guard = self.inner.lock().map_err(|_| {
            DesktopStorageEncryptionError::Keychain("keychain mutex is poisoned".to_string())
        })?;
        guard.remove(&(service.to_string(), account.to_string()));
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct KeyRing {
    current_version: u32,
    keys_by_version: BTreeMap<u32, [u8; AES_256_KEY_BYTES]>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct KeyRingJson {
    pub current_version: u32,
    pub keys: BTreeMap<String, String>,
}

impl KeyRing {
    pub fn create() -> Self {
        let mut key = [0u8; AES_256_KEY_BYTES];
        rand_core::OsRng.fill_bytes(&mut key);
        let mut keys_by_version = BTreeMap::new();
        keys_by_version.insert(1, key);
        Self {
            current_version: 1,
            keys_by_version,
        }
    }

    pub fn rotate(&mut self) -> u32 {
        let next_version = self.current_version.saturating_add(1);
        let mut key = [0u8; AES_256_KEY_BYTES];
        rand_core::OsRng.fill_bytes(&mut key);
        self.keys_by_version.insert(next_version, key);
        self.current_version = next_version;
        next_version
    }

    fn get_key(&self, version: u32) -> Result<&[u8; AES_256_KEY_BYTES], DesktopStorageEncryptionError> {
        self.keys_by_version
            .get(&version)
            .ok_or(DesktopStorageEncryptionError::MissingKeyVersion(version))
    }

    pub fn encrypt(
        &self,
        plaintext: &[u8],
        aad_context: Option<&serde_json::Value>,
    ) -> Result<EncryptedEnvelope, DesktopStorageEncryptionError> {
        self.encrypt_with_iv(plaintext, aad_context, None)
    }

    pub fn encrypt_with_iv(
        &self,
        plaintext: &[u8],
        aad_context: Option<&serde_json::Value>,
        iv: Option<[u8; AES_GCM_IV_BYTES]>,
    ) -> Result<EncryptedEnvelope, DesktopStorageEncryptionError> {
        let aad = aad_from_context(aad_context)?;
        let key_version = self.current_version;
        let key = self.get_key(key_version)?;
        let encrypted = encrypt_aes256_gcm(plaintext, key, aad.as_deref(), iv)?;

        Ok(EncryptedEnvelope {
            key_version,
            algorithm: AES_256_GCM.to_string(),
            iv: general_purpose::STANDARD.encode(encrypted.iv),
            ciphertext: general_purpose::STANDARD.encode(encrypted.ciphertext),
            tag: general_purpose::STANDARD.encode(encrypted.tag),
            created_at: None,
        })
    }

    pub fn decrypt(
        &self,
        encrypted: &EncryptedEnvelope,
        aad_context: Option<&serde_json::Value>,
    ) -> Result<Vec<u8>, DesktopStorageEncryptionError> {
        if encrypted.algorithm != AES_256_GCM {
            return Err(DesktopStorageEncryptionError::UnsupportedAlgorithm(
                encrypted.algorithm.clone(),
            ));
        }
        let aad = aad_from_context(aad_context)?;
        let iv = decode_fixed::<AES_GCM_IV_BYTES>(&encrypted.iv, "iv")?;
        let tag = decode_fixed::<AES_GCM_TAG_BYTES>(&encrypted.tag, "tag")?;
        let ciphertext = general_purpose::STANDARD.decode(&encrypted.ciphertext)?;
        let key = self.get_key(encrypted.key_version)?;

        decrypt_aes256_gcm(&ciphertext, key, &iv, &tag, aad.as_deref())
    }

    pub fn to_json(&self) -> KeyRingJson {
        let mut keys = BTreeMap::new();
        for (version, key) in &self.keys_by_version {
            keys.insert(version.to_string(), general_purpose::STANDARD.encode(key));
        }
        KeyRingJson {
            current_version: self.current_version,
            keys,
        }
    }

    pub fn from_json(value: KeyRingJson) -> Result<Self, DesktopStorageEncryptionError> {
        if value.current_version < 1 {
            return Err(DesktopStorageEncryptionError::InvalidPayload(
                "currentVersion must be >= 1".to_string(),
            ));
        }

        let mut keys_by_version = BTreeMap::new();
        for (version_str, key_b64) in value.keys {
            let version: u32 = version_str.parse().map_err(|_| {
                DesktopStorageEncryptionError::InvalidPayload(format!("invalid key version: {version_str}"))
            })?;
            if version < 1 {
                return Err(DesktopStorageEncryptionError::InvalidPayload(format!(
                    "invalid key version: {version}"
                )));
            }

            let key = decode_fixed::<AES_256_KEY_BYTES>(&key_b64, &format!("keys[{version}]"))?;
            keys_by_version.insert(version, key);
        }

        if !keys_by_version.contains_key(&value.current_version) {
            return Err(DesktopStorageEncryptionError::InvalidPayload(
                "keys must include currentVersion".to_string(),
            ));
        }

        Ok(Self {
            current_version: value.current_version,
            keys_by_version,
        })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>, DesktopStorageEncryptionError> {
        Ok(serde_json::to_vec(&self.to_json())?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self, DesktopStorageEncryptionError> {
        let parsed: KeyRingJson = serde_json::from_slice(bytes)?;
        Self::from_json(parsed)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EncryptedEnvelope {
    pub key_version: u32,
    pub algorithm: String,
    pub iv: String,
    pub ciphertext: String,
    pub tag: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub created_at: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct PlaintextStoreFile {
    schema_version: u32,
    encrypted: bool,
    #[serde(default)]
    documents: BTreeMap<String, serde_json::Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct EncryptedStoreFile {
    schema_version: u32,
    encrypted: bool,
    #[serde(flatten)]
    envelope: EncryptedEnvelope,
}

#[derive(Debug, Clone)]
enum StoreFile {
    Plaintext(PlaintextStoreFile),
    Encrypted(EncryptedStoreFile),
}

fn aad_from_context(
    context: Option<&serde_json::Value>,
) -> Result<Option<Vec<u8>>, DesktopStorageEncryptionError> {
    let Some(context) = context else {
        return Ok(None);
    };
    let canonicalized = canonical_json(context);
    Ok(Some(serde_json::to_string(&canonicalized)?.into_bytes()))
}

fn canonical_json(value: &serde_json::Value) -> serde_json::Value {
    match value {
        serde_json::Value::Array(values) => {
            serde_json::Value::Array(values.iter().map(canonical_json).collect())
        }
        serde_json::Value::Object(map) => {
            let mut keys: Vec<&String> = map.keys().collect();
            keys.sort();
            let mut out = serde_json::Map::with_capacity(map.len());
            for key in keys {
                out.insert(key.clone(), canonical_json(&map[key]));
            }
            serde_json::Value::Object(out)
        }
        other => other.clone(),
    }
}

fn decode_fixed<const N: usize>(
    value_b64: &str,
    field: &str,
) -> Result<[u8; N], DesktopStorageEncryptionError> {
    let decoded = general_purpose::STANDARD.decode(value_b64)?;
    let decoded_len = decoded.len();
    let bytes: [u8; N] = decoded.try_into().map_err(|_| {
        DesktopStorageEncryptionError::InvalidPayload(format!(
            "{field} must decode to {N} bytes (got {})",
            decoded_len
        ))
    })?;
    Ok(bytes)
}

struct RawEncryptedPayload {
    iv: [u8; AES_GCM_IV_BYTES],
    ciphertext: Vec<u8>,
    tag: [u8; AES_GCM_TAG_BYTES],
}

fn encrypt_aes256_gcm(
    plaintext: &[u8],
    key: &[u8; AES_256_KEY_BYTES],
    aad: Option<&[u8]>,
    iv: Option<[u8; AES_GCM_IV_BYTES]>,
) -> Result<RawEncryptedPayload, DesktopStorageEncryptionError> {
    let mut nonce_bytes = iv.unwrap_or([0u8; AES_GCM_IV_BYTES]);
    if iv.is_none() {
        rand_core::OsRng.fill_bytes(&mut nonce_bytes);
    }

    let cipher = Aes256Gcm::new_from_slice(key).map_err(|_| DesktopStorageEncryptionError::Crypto)?;
    let nonce = Nonce::from_slice(&nonce_bytes);
    let mut buffer = plaintext.to_vec();
    let tag = cipher
        .encrypt_in_place_detached(nonce, aad.unwrap_or(&[]), &mut buffer)
        .map_err(|_| DesktopStorageEncryptionError::Crypto)?;
    let tag_bytes: [u8; AES_GCM_TAG_BYTES] = tag.into();

    Ok(RawEncryptedPayload {
        iv: nonce_bytes,
        ciphertext: buffer,
        tag: tag_bytes,
    })
}

fn decrypt_aes256_gcm(
    ciphertext: &[u8],
    key: &[u8; AES_256_KEY_BYTES],
    iv: &[u8; AES_GCM_IV_BYTES],
    tag: &[u8; AES_GCM_TAG_BYTES],
    aad: Option<&[u8]>,
) -> Result<Vec<u8>, DesktopStorageEncryptionError> {
    let cipher = Aes256Gcm::new_from_slice(key).map_err(|_| DesktopStorageEncryptionError::Crypto)?;
    let nonce = Nonce::from_slice(iv);
    let mut buffer = ciphertext.to_vec();
    cipher
        .decrypt_in_place_detached(nonce, aad.unwrap_or(&[]), &mut buffer, tag.into())
        .map_err(|_| DesktopStorageEncryptionError::Crypto)?;
    Ok(buffer)
}

fn read_store_file(path: &Path) -> Result<StoreFile, DesktopStorageEncryptionError> {
    let raw = match fs::read_to_string(path) {
        Ok(raw) => raw,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            return Ok(StoreFile::Plaintext(PlaintextStoreFile {
                schema_version: DEFAULT_STORE_SCHEMA_VERSION,
                encrypted: false,
                documents: BTreeMap::new(),
            }));
        }
        Err(err) => return Err(err.into()),
    };

    let value: serde_json::Value = serde_json::from_str(&raw)?;
    let encrypted = value
        .get("encrypted")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);

    if encrypted {
        Ok(StoreFile::Encrypted(serde_json::from_value(value)?))
    } else {
        Ok(StoreFile::Plaintext(serde_json::from_value(value)?))
    }
}

fn write_json_file(path: &Path, value: &impl Serialize) -> Result<(), DesktopStorageEncryptionError> {
    let json = serde_json::to_string_pretty(value)?;
    write_file_atomic_io(path, json.as_bytes())?;
    Ok(())
}

fn store_aad(schema_version: u32, scope: &str) -> serde_json::Value {
    serde_json::json!({
        "scope": scope,
        "schemaVersion": schema_version
    })
}

/// Desktop store wrapper that can transparently migrate between plaintext and
/// encrypted on-disk formats.
#[derive(Debug, Clone)]
pub struct DesktopStorageEncryption<P: KeychainProvider> {
    file_path: PathBuf,
    keychain: P,
    keychain_service: String,
    keychain_account: String,
    aad_scope: String,
}

impl<P: KeychainProvider> DesktopStorageEncryption<P> {
    pub fn new(file_path: impl Into<PathBuf>, keychain: P) -> Self {
        Self {
            file_path: file_path.into(),
            keychain,
            keychain_service: DEFAULT_KEYCHAIN_SERVICE.to_string(),
            keychain_account: DEFAULT_KEYCHAIN_ACCOUNT.to_string(),
            aad_scope: DEFAULT_AAD_SCOPE.to_string(),
        }
    }

    pub fn with_keychain_namespace(
        mut self,
        service: impl Into<String>,
        account: impl Into<String>,
    ) -> Self {
        self.keychain_service = service.into();
        self.keychain_account = account.into();
        self
    }

    pub fn with_aad_scope(mut self, scope: impl Into<String>) -> Self {
        self.aad_scope = scope.into();
        self
    }

    /// Runs `operation`. If it fails because an encrypted store exists on disk but the
    /// keyring is missing from the keychain (`DesktopStorageEncryptionError::MissingKeyRing`),
    /// the store is reset to an empty encrypted store with a fresh keyring and the operation
    /// is retried once.
    ///
    /// Returns `(result, recovered)` where `recovered` indicates whether the store was reset.
    pub fn with_missing_keyring_recovery<T, F>(
        &self,
        mut operation: F,
    ) -> Result<(T, bool), DesktopStorageEncryptionError>
    where
        F: FnMut() -> Result<T, DesktopStorageEncryptionError>,
    {
        match operation() {
            Ok(value) => Ok((value, false)),
            Err(DesktopStorageEncryptionError::MissingKeyRing) => {
                self.reset_encrypted_store_for_missing_keyring()?;
                Ok((operation()?, true))
            }
            Err(err) => Err(err),
        }
    }

    fn reset_encrypted_store_for_missing_keyring(&self) -> Result<(), DesktopStorageEncryptionError> {
        // Preserve the schema version from disk if we can parse it (encrypted stores can be
        // parsed without needing the keyring).
        let schema_version = match read_store_file(&self.file_path)? {
            StoreFile::Encrypted(encrypted) => encrypted.schema_version,
            StoreFile::Plaintext(plain) => plain.schema_version,
        };

        crate::stdio::stderrln(format_args!(
            "[storage] WARNING: resetting encrypted store at {:?} because keyring was missing in the OS keychain (service={}, account={}). Existing data was discarded.",
            self.file_path,
            self.keychain_service,
            self.keychain_account
        ));

        // Create a fresh keyring and re-initialize the store as empty ciphertext.
        //
        // We write the on-disk store first (using the new keyring) to avoid leaving a state
        // where the keychain contains a new keyring but the store still contains ciphertext
        // encrypted with the old (now-missing) keyring. If we crash mid-recovery, the worst
        // case is an encrypted empty store without a keyring, which will trigger recovery
        // again on the next load.
        let keyring = KeyRing::create();

        let plaintext = serde_json::to_vec(&PlaintextStoreFile {
            schema_version,
            encrypted: false,
            documents: BTreeMap::new(),
        })?;
        let aad = store_aad(schema_version, &self.aad_scope);
        let envelope = keyring.encrypt(&plaintext, Some(&aad))?;

        write_json_file(
            &self.file_path,
            &EncryptedStoreFile {
                schema_version,
                encrypted: true,
                envelope,
            },
        )?;

        self.store_keyring(&keyring)?;

        Ok(())
    }

    fn load_keyring(&self) -> Result<Option<KeyRing>, DesktopStorageEncryptionError> {
        let secret = self
            .keychain
            .get_secret(&self.keychain_service, &self.keychain_account)?;
        let Some(secret) = secret else {
            return Ok(None);
        };
        Ok(Some(KeyRing::from_bytes(&secret)?))
    }

    fn store_keyring(&self, keyring: &KeyRing) -> Result<(), DesktopStorageEncryptionError> {
        let bytes = keyring.to_bytes()?;
        self.keychain
            .set_secret(&self.keychain_service, &self.keychain_account, &bytes)
    }

    fn delete_keyring(&self) -> Result<(), DesktopStorageEncryptionError> {
        self.keychain
            .delete_secret(&self.keychain_service, &self.keychain_account)
    }

    fn load_plaintext_documents(&self) -> Result<(u32, BTreeMap<String, serde_json::Value>), DesktopStorageEncryptionError> {
        match read_store_file(&self.file_path)? {
            StoreFile::Plaintext(plain) => Ok((plain.schema_version, plain.documents)),
            StoreFile::Encrypted(encrypted) => {
                let keyring = self.load_keyring()?.ok_or(DesktopStorageEncryptionError::MissingKeyRing)?;
                let aad = store_aad(encrypted.schema_version, &self.aad_scope);
                let plaintext_bytes = keyring.decrypt(&encrypted.envelope, Some(&aad))?;
                let parsed: PlaintextStoreFile = serde_json::from_slice(&plaintext_bytes)?;
                Ok((parsed.schema_version, parsed.documents))
            }
        }
    }

    fn write_documents(
        &self,
        schema_version: u32,
        documents: BTreeMap<String, serde_json::Value>,
        encrypted: bool,
    ) -> Result<(), DesktopStorageEncryptionError> {
        if !encrypted {
            return write_json_file(
                &self.file_path,
                &PlaintextStoreFile {
                    schema_version,
                    encrypted: false,
                    documents,
                },
            );
        }

        let keyring = match self.load_keyring()? {
            Some(existing) => existing,
            None => {
                let keyring = KeyRing::create();
                self.store_keyring(&keyring)?;
                keyring
            }
        };

        // Ensure keyring is stored even if we created it above (and allow
        // future key metadata changes).
        self.store_keyring(&keyring)?;

        let plaintext = serde_json::to_vec(&PlaintextStoreFile {
            schema_version,
            encrypted: false,
            documents,
        })?;
        let aad = store_aad(schema_version, &self.aad_scope);
        let envelope = keyring.encrypt(&plaintext, Some(&aad))?;

        write_json_file(
            &self.file_path,
            &EncryptedStoreFile {
                schema_version,
                encrypted: true,
                envelope,
            },
        )
    }

    /// Enable encryption-at-rest for the on-disk store and migrate existing
    /// plaintext data to ciphertext.
    pub fn enable_encryption(&self) -> Result<(), DesktopStorageEncryptionError> {
        let (schema_version, documents) = self.load_plaintext_documents()?;
        self.write_documents(schema_version, documents, true)
    }

    pub fn ensure_encrypted(&self) -> Result<(), DesktopStorageEncryptionError> {
        match read_store_file(&self.file_path)? {
            StoreFile::Encrypted(_) => Ok(()),
            StoreFile::Plaintext(_) => self.enable_encryption(),
        }
    }

    /// Disable encryption-at-rest and migrate ciphertext back to plaintext.
    ///
    /// If `delete_key` is true, the stored keyring is removed from the keychain.
    pub fn disable_encryption(&self, delete_key: bool) -> Result<(), DesktopStorageEncryptionError> {
        let (schema_version, documents) = self.load_plaintext_documents()?;
        self.write_documents(schema_version, documents, false)?;
        if delete_key {
            self.delete_keyring()?;
        }
        Ok(())
    }

    /// Rotate the active data key and re-encrypt the store on disk.
    ///
    /// Old key versions are preserved in the keyring so older ciphertexts can
    /// still be decrypted (e.g. backups), while the current store is re-written
    /// with the new key version.
    pub fn rotate_key(&self) -> Result<u32, DesktopStorageEncryptionError> {
        let on_disk = read_store_file(&self.file_path)?;
        let StoreFile::Encrypted(encrypted) = on_disk else {
            return Err(DesktopStorageEncryptionError::StoreNotEncrypted);
        };

        let (schema_version, documents) = self.load_plaintext_documents()?;
        let mut keyring = self.load_keyring()?.ok_or(DesktopStorageEncryptionError::MissingKeyRing)?;
        let next_version = keyring.rotate();
        self.store_keyring(&keyring)?;

        self.write_documents(schema_version, documents, true)?;

        // Ensure we actually bumped the key version on disk.
        let StoreFile::Encrypted(updated) = read_store_file(&self.file_path)? else {
            return Err(DesktopStorageEncryptionError::StoreNotEncrypted);
        };
        if updated.envelope.key_version != next_version || updated.envelope.key_version <= encrypted.envelope.key_version {
            return Err(DesktopStorageEncryptionError::InvalidPayload(
                "key rotation did not update keyVersion on disk".to_string(),
            ));
        }

        Ok(next_version)
    }

    pub fn save_document(
        &self,
        doc_id: &str,
        document: serde_json::Value,
    ) -> Result<(), DesktopStorageEncryptionError> {
        if doc_id.is_empty() {
            return Err(DesktopStorageEncryptionError::InvalidPayload(
                "docId must be a non-empty string".to_string(),
            ));
        }

        let on_disk = read_store_file(&self.file_path)?;
        let (schema_version, mut documents) = self.load_plaintext_documents()?;
        documents.insert(doc_id.to_string(), document);
        let should_encrypt = matches!(on_disk, StoreFile::Encrypted(_));
        self.write_documents(schema_version, documents, should_encrypt)
    }

    pub fn load_document(
        &self,
        doc_id: &str,
    ) -> Result<Option<serde_json::Value>, DesktopStorageEncryptionError> {
        let (_, documents) = self.load_plaintext_documents()?;
        Ok(documents.get(doc_id).cloned())
    }

    pub fn delete_document(&self, doc_id: &str) -> Result<(), DesktopStorageEncryptionError> {
        if doc_id.is_empty() {
            return Err(DesktopStorageEncryptionError::InvalidPayload(
                "docId must be a non-empty string".to_string(),
            ));
        }

        let on_disk = read_store_file(&self.file_path)?;
        let (schema_version, mut documents) = self.load_plaintext_documents()?;
        documents.remove(doc_id);
        let should_encrypt = matches!(on_disk, StoreFile::Encrypted(_));
        self.write_documents(schema_version, documents, should_encrypt)
    }

    pub fn list_document_ids(&self) -> Result<Vec<String>, DesktopStorageEncryptionError> {
        let (_, documents) = self.load_plaintext_documents()?;
        Ok(documents.keys().cloned().collect())
    }

    pub fn file_path(&self) -> &Path {
        &self.file_path
    }

    pub fn keychain_provider(&self) -> &P {
        &self.keychain
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use serde_json::json;

    #[test]
    fn aad_canonicalization_matches_js() {
        // JS canonicalJson sorts object keys, so "schemaVersion" comes before "scope".
        let aad = store_aad(1, DEFAULT_AAD_SCOPE);
        let bytes = aad_from_context(Some(&aad)).unwrap().unwrap();
        let as_str = std::str::from_utf8(&bytes).unwrap();
        assert_eq!(
            as_str,
            "{\"schemaVersion\":1,\"scope\":\"formula-desktop-store\"}"
        );
    }

    #[test]
    fn aes256gcm_matches_node_crypto_for_fixed_vector() {
        let key = decode_fixed::<AES_256_KEY_BYTES>(
            "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8=",
            "key",
        )
        .unwrap();
        let iv = decode_fixed::<AES_GCM_IV_BYTES>("GvOMLcK5b/3YZpQJ", "iv").unwrap();
        let aad = b"{\"schemaVersion\":1,\"scope\":\"formula-desktop-store\"}";
        let plaintext = b"example plaintext payload";

        let encrypted = encrypt_aes256_gcm(plaintext, &key, Some(aad), Some(iv)).unwrap();
        assert_eq!(
            general_purpose::STANDARD.encode(&encrypted.ciphertext),
            "xwj/oEeNg2EMTZ9hOsFvzfEmqwXB9hoXkg=="
        );
        assert_eq!(
            general_purpose::STANDARD.encode(encrypted.tag),
            "+5Xz+PWjxxL4+mo2nRKXzQ=="
        );

        let decrypted =
            decrypt_aes256_gcm(&encrypted.ciphertext, &key, &encrypted.iv, &encrypted.tag, Some(aad))
                .unwrap();
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn keyring_roundtrip_matches_js_shape() {
        let mut keyring = KeyRing::create();
        keyring.rotate();
        let json = keyring.to_json();
        let serialized = serde_json::to_string(&json).unwrap();
        assert!(serialized.contains("\"currentVersion\""));
        assert!(serialized.contains("\"keys\""));
        let parsed: KeyRingJson = serde_json::from_str(&serialized).unwrap();
        let rebuilt = KeyRing::from_json(parsed).unwrap();
        assert_eq!(rebuilt.current_version, keyring.current_version);
        assert_eq!(rebuilt.keys_by_version.len(), keyring.keys_by_version.len());
    }

    #[test]
    fn wrong_aad_fails_to_decrypt() {
        let keyring = KeyRing::create();
        let plaintext = b"secret";
        let good_aad = json!({ "scope": DEFAULT_AAD_SCOPE, "schemaVersion": 1 });
        let bad_aad = json!({ "scope": "other-scope", "schemaVersion": 1 });
        let encrypted = keyring.encrypt(plaintext, Some(&good_aad)).unwrap();
        let err = keyring.decrypt(&encrypted, Some(&bad_aad)).unwrap_err();
        assert!(matches!(err, DesktopStorageEncryptionError::Crypto));
    }
}
