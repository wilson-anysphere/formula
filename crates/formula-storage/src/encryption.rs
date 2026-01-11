use aes_gcm::aead::{AeadInPlace, KeyInit};
use aes_gcm::{Aes256Gcm, Key, Nonce, Tag};
use base64::engine::general_purpose::STANDARD_NO_PAD;
use base64::Engine;
use rand::rngs::OsRng;
use rand::RngCore;
use serde::{de::Error as _, Deserialize, Deserializer, Serialize, Serializer};
use std::collections::BTreeMap;
use std::sync::Mutex;
use thiserror::Error;

const CONTAINER_MAGIC: &[u8; 8] = b"FSTORAGE";
const CONTAINER_VERSION: u8 = 1;

const KEY_LEN: usize = 32;
const NONCE_LEN: usize = 12;
const TAG_LEN: usize = 16;

const HEADER_LEN: usize = 8 /* magic */
    + 1 /* container version */
    + 4 /* key version */
    + NONCE_LEN
    + TAG_LEN;

#[derive(Debug, Error)]
pub enum EncryptionError {
    #[error("encrypted container is truncated")]
    TruncatedContainer,
    #[error("encrypted container magic header mismatch")]
    InvalidMagic,
    #[error("unsupported encrypted container version: {0}")]
    UnsupportedContainerVersion(u8),
    #[error("missing key for version {0}")]
    MissingKey(u32),
    #[error("invalid key length: expected 32 bytes, got {0}")]
    InvalidKeyLength(usize),
    #[error("base64 error: {0}")]
    Base64(#[from] base64::DecodeError),
    #[error("json error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("key provider error: {0}")]
    KeyProvider(#[from] KeyProviderError),
    #[error("aes-gcm error")]
    Aead,
}

impl From<aes_gcm::aead::Error> for EncryptionError {
    fn from(_: aes_gcm::aead::Error) -> Self {
        EncryptionError::Aead
    }
}

#[derive(Debug, Error, Clone)]
#[error("{0}")]
pub struct KeyProviderError(pub String);

impl KeyProviderError {
    pub fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

/// Consumer-provided key management hook.
///
/// `formula-storage` stays self-contained; platform-specific keychain integration should live
/// in the consumer (e.g. desktop) by implementing this trait.
pub trait KeyProvider: Send + Sync + 'static {
    fn load_keyring(&self) -> std::result::Result<Option<KeyRing>, KeyProviderError>;
    fn store_keyring(&self, keyring: &KeyRing) -> std::result::Result<(), KeyProviderError>;
}

#[derive(Default)]
pub struct InMemoryKeyProvider {
    keyring: Mutex<Option<KeyRing>>,
}

impl InMemoryKeyProvider {
    pub fn new(keyring: Option<KeyRing>) -> Self {
        Self {
            keyring: Mutex::new(keyring),
        }
    }

    pub fn keyring(&self) -> Option<KeyRing> {
        self.keyring.lock().expect("key provider mutex poisoned").clone()
    }
}

impl std::fmt::Debug for InMemoryKeyProvider {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("InMemoryKeyProvider").finish()
    }
}

impl KeyProvider for InMemoryKeyProvider {
    fn load_keyring(&self) -> std::result::Result<Option<KeyRing>, KeyProviderError> {
        Ok(self.keyring.lock().expect("key provider mutex poisoned").clone())
    }

    fn store_keyring(&self, keyring: &KeyRing) -> std::result::Result<(), KeyProviderError> {
        *self.keyring.lock().expect("key provider mutex poisoned") = Some(keyring.clone());
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct KeyBytes([u8; KEY_LEN]);

impl KeyBytes {
    pub fn new(bytes: [u8; KEY_LEN]) -> Self {
        Self(bytes)
    }

    pub fn as_bytes(&self) -> &[u8; KEY_LEN] {
        &self.0
    }
}

impl Serialize for KeyBytes {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.serialize_str(&STANDARD_NO_PAD.encode(self.0))
    }
}

impl<'de> Deserialize<'de> for KeyBytes {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let encoded = String::deserialize(deserializer)?;
        let decoded = STANDARD_NO_PAD.decode(encoded).map_err(D::Error::custom)?;
        if decoded.len() != KEY_LEN {
            return Err(D::Error::custom(format!(
                "expected {KEY_LEN} bytes, got {}",
                decoded.len()
            )));
        }
        let mut bytes = [0u8; KEY_LEN];
        bytes.copy_from_slice(&decoded);
        Ok(Self(bytes))
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct KeyRing {
    pub current_version: u32,
    pub keys: BTreeMap<u32, KeyBytes>,
}

impl KeyRing {
    pub fn new_random() -> Self {
        let mut key = [0u8; KEY_LEN];
        OsRng.fill_bytes(&mut key);
        let mut keys = BTreeMap::new();
        keys.insert(1, KeyBytes::new(key));
        Self {
            current_version: 1,
            keys,
        }
    }

    pub fn from_key(version: u32, key: [u8; KEY_LEN]) -> Self {
        let mut keys = BTreeMap::new();
        keys.insert(version, KeyBytes::new(key));
        Self {
            current_version: version,
            keys,
        }
    }

    pub fn insert_key(&mut self, version: u32, key: [u8; KEY_LEN]) {
        self.keys.insert(version, KeyBytes::new(key));
        self.current_version = self.current_version.max(version);
    }

    pub fn rotate(&mut self) {
        let next = self.current_version + 1;
        let mut key = [0u8; KEY_LEN];
        OsRng.fill_bytes(&mut key);
        self.keys.insert(next, KeyBytes::new(key));
        self.current_version = next;
    }

    pub fn current_key(&self) -> Result<(u32, [u8; KEY_LEN]), EncryptionError> {
        let key = self
            .keys
            .get(&self.current_version)
            .ok_or(EncryptionError::MissingKey(self.current_version))?;
        Ok((self.current_version, *key.as_bytes()))
    }

    pub fn key(&self, version: u32) -> Option<[u8; KEY_LEN]> {
        self.keys.get(&version).map(|k| *k.as_bytes())
    }
}

pub fn is_encrypted_container(bytes: &[u8]) -> bool {
    bytes.len() >= CONTAINER_MAGIC.len() && &bytes[..CONTAINER_MAGIC.len()] == CONTAINER_MAGIC
}

pub fn encrypt_sqlite_bytes(plaintext: &[u8], keyring: &KeyRing) -> Result<Vec<u8>, EncryptionError> {
    let (key_version, key_bytes) = keyring.current_key()?;
    encrypt_sqlite_bytes_with_key(plaintext, key_version, &key_bytes)
}

fn encrypt_sqlite_bytes_with_key(
    plaintext: &[u8],
    key_version: u32,
    key_bytes: &[u8; KEY_LEN],
) -> Result<Vec<u8>, EncryptionError> {
    let mut nonce_bytes = [0u8; NONCE_LEN];
    OsRng.fill_bytes(&mut nonce_bytes);
    let nonce = Nonce::from_slice(&nonce_bytes);

    let cipher = Aes256Gcm::new(Key::<Aes256Gcm>::from_slice(key_bytes));
    let mut buffer = plaintext.to_vec();
    let tag = cipher.encrypt_in_place_detached(nonce, &aad_for_key_version(key_version), &mut buffer)?;

    let mut out = Vec::with_capacity(HEADER_LEN + buffer.len());
    out.extend_from_slice(CONTAINER_MAGIC);
    out.push(CONTAINER_VERSION);
    out.extend_from_slice(&key_version.to_be_bytes());
    out.extend_from_slice(&nonce_bytes);
    out.extend_from_slice(tag.as_slice());
    out.extend_from_slice(&buffer);
    Ok(out)
}

pub fn decrypt_sqlite_bytes(container: &[u8], keyring: &KeyRing) -> Result<Vec<u8>, EncryptionError> {
    let parsed = parse_container(container)?;
    let key = keyring
        .key(parsed.key_version)
        .ok_or(EncryptionError::MissingKey(parsed.key_version))?;

    let cipher = Aes256Gcm::new(Key::<Aes256Gcm>::from_slice(&key));
    let mut buffer = parsed.ciphertext.to_vec();
    let nonce = Nonce::from_slice(&parsed.nonce);
    cipher.decrypt_in_place_detached(
        nonce,
        &aad_for_key_version(parsed.key_version),
        &mut buffer,
        Tag::from_slice(&parsed.tag),
    )?;
    Ok(buffer)
}

#[derive(Debug)]
struct ParsedContainer<'a> {
    key_version: u32,
    nonce: [u8; NONCE_LEN],
    tag: [u8; TAG_LEN],
    ciphertext: &'a [u8],
}

fn parse_container(bytes: &[u8]) -> Result<ParsedContainer<'_>, EncryptionError> {
    if bytes.len() < HEADER_LEN {
        return Err(EncryptionError::TruncatedContainer);
    }
    if &bytes[..CONTAINER_MAGIC.len()] != CONTAINER_MAGIC {
        return Err(EncryptionError::InvalidMagic);
    }
    let version = bytes[CONTAINER_MAGIC.len()];
    if version != CONTAINER_VERSION {
        return Err(EncryptionError::UnsupportedContainerVersion(version));
    }
    let mut key_version_bytes = [0u8; 4];
    key_version_bytes.copy_from_slice(&bytes[9..13]);
    let key_version = u32::from_be_bytes(key_version_bytes);

    let mut nonce = [0u8; NONCE_LEN];
    nonce.copy_from_slice(&bytes[13..25]);
    let mut tag = [0u8; TAG_LEN];
    tag.copy_from_slice(&bytes[25..41]);

    Ok(ParsedContainer {
        key_version,
        nonce,
        tag,
        ciphertext: &bytes[HEADER_LEN..],
    })
}

fn aad_for_key_version(key_version: u32) -> [u8; 8 + 1 + 4] {
    let mut aad = [0u8; 13];
    aad[..8].copy_from_slice(CONTAINER_MAGIC);
    aad[8] = CONTAINER_VERSION;
    aad[9..13].copy_from_slice(&key_version.to_be_bytes());
    aad
}

pub fn load_or_create_keyring(
    provider: &dyn KeyProvider,
    create_if_missing: bool,
) -> Result<KeyRing, EncryptionError> {
    match provider.load_keyring()? {
        Some(keyring) => Ok(keyring),
        None if create_if_missing => {
            let keyring = KeyRing::new_random();
            provider.store_keyring(&keyring)?;
            Ok(keyring)
        }
        None => Err(EncryptionError::KeyProvider(KeyProviderError::new(
            "missing keyring",
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn encrypt_decrypt_round_trip() {
        let keyring = KeyRing::from_key(1, [7u8; KEY_LEN]);
        let plaintext = b"sqlite bytes go here";
        let encrypted = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt");
        assert!(is_encrypted_container(&encrypted));
        let decrypted = decrypt_sqlite_bytes(&encrypted, &keyring).expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn key_rotation_retains_old_versions() {
        let mut keyring = KeyRing::from_key(1, [1u8; KEY_LEN]);
        let plaintext = b"workbook";
        let encrypted_v1 = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt v1");

        keyring.rotate();

        let decrypted = decrypt_sqlite_bytes(&encrypted_v1, &keyring).expect("decrypt v1 after rotate");
        assert_eq!(decrypted, plaintext);

        let encrypted_v2 = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt v2");
        let parsed_v2 = parse_container(&encrypted_v2).expect("parse v2");
        assert_eq!(parsed_v2.key_version, 2);
    }

    #[test]
    fn tamper_detection_fails() {
        let keyring = KeyRing::from_key(1, [2u8; KEY_LEN]);
        let plaintext = b"some bytes";
        let mut encrypted = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt");

        // Flip a bit in the tag.
        encrypted[30] ^= 0b0000_0001;
        let err = decrypt_sqlite_bytes(&encrypted, &keyring).expect_err("decrypt should fail");
        match err {
            EncryptionError::Aead => {}
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
