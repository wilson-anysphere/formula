//! Encryption-at-rest primitives for `formula-storage`.
//!
//! # Container format
//! The persisted workbook bytes are stored in a small versioned container so we can distinguish
//! encrypted data from plaintext SQLite files.
//!
//! This crate writes **version 1** using the same header as the JS implementation
//! (`packages/security/crypto/encryptedFile.js`):
//!
//! ```text
//! 8B   magic:      "FMLENC01"
//! 4B   keyVersion: uint32 big-endian
//! 12B  iv:         AES-GCM nonce
//! 16B  tag:        AES-GCM authentication tag
//! ...  ciphertext
//! ```
//!
//! We also support reading a legacy Rust-only format (magic `"FSTORAGE"`) for backwards
//! compatibility with earlier iterations of this crate.
//!
//! # Key management
//! Consumers provide key storage via [`KeyProvider`]. The crate includes an in-memory provider for
//! tests; production consumers should back this with an OS keychain or other secure secret store.

use aes_gcm::aead::{AeadInPlace, KeyInit};
use aes_gcm::{Aes256Gcm, Key, Nonce, Tag};
use base64::engine::general_purpose::{STANDARD, STANDARD_NO_PAD};
use base64::Engine;
use rand::rngs::OsRng;
use rand::RngCore;
use serde::{de::Error as _, Deserialize, Deserializer, Serialize, Serializer};
use std::collections::BTreeMap;
use std::sync::Mutex;
use thiserror::Error;

use crate::lock_unpoisoned;

/// Matches the JS `packages/security/crypto/encryptedFile.js` magic header for encrypted blobs.
/// The trailing two digits encode the container version.
const MAGIC_FMLENC_V1: &[u8; 8] = b"FMLENC01";
const MAGIC_FMLENC_PREFIX: &[u8; 6] = b"FMLENC";

/// Legacy container magic used by earlier `formula-storage` versions.
const LEGACY_MAGIC: &[u8; 8] = b"FSTORAGE";
const LEGACY_CONTAINER_VERSION: u8 = 1;

const KEY_LEN: usize = 32;
const NONCE_LEN: usize = 12;
const TAG_LEN: usize = 16;

const HEADER_LEN_FMLENC_V1: usize = 8 /* magic */ + 4 /* key version */ + NONCE_LEN + TAG_LEN;
const HEADER_LEN_LEGACY_V1: usize = 8 /* magic */
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
        lock_unpoisoned(&self.keyring).clone()
    }
}

impl std::fmt::Debug for InMemoryKeyProvider {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("InMemoryKeyProvider").finish()
    }
}

impl KeyProvider for InMemoryKeyProvider {
    fn load_keyring(&self) -> std::result::Result<Option<KeyRing>, KeyProviderError> {
        Ok(lock_unpoisoned(&self.keyring).clone())
    }

    fn store_keyring(&self, keyring: &KeyRing) -> std::result::Result<(), KeyProviderError> {
        *lock_unpoisoned(&self.keyring) = Some(keyring.clone());
        Ok(())
    }
}

#[derive(Clone, PartialEq, Eq)]
pub struct KeyBytes([u8; KEY_LEN]);

impl KeyBytes {
    pub fn new(bytes: [u8; KEY_LEN]) -> Self {
        Self(bytes)
    }

    pub fn as_bytes(&self) -> &[u8; KEY_LEN] {
        &self.0
    }
}

impl std::fmt::Debug for KeyBytes {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("KeyBytes")
            .field(&format_args!("<redacted; {KEY_LEN} bytes>"))
            .finish()
    }
}

impl Serialize for KeyBytes {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        serializer.serialize_str(&STANDARD.encode(self.0))
    }
}

impl<'de> Deserialize<'de> for KeyBytes {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let encoded = String::deserialize(deserializer)?;
        // Accept both padded (Node/JS) and unpadded (older Rust) base64 encodings.
        let decoded = STANDARD
            .decode(encoded.as_bytes())
            .or_else(|_| STANDARD_NO_PAD.decode(encoded.as_bytes()))
            .map_err(D::Error::custom)?;
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

#[derive(Clone, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct KeyRing {
    pub current_version: u32,
    pub keys: BTreeMap<u32, KeyBytes>,
}

impl std::fmt::Debug for KeyRing {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let versions: Vec<u32> = self.keys.keys().copied().collect();
        f.debug_struct("KeyRing")
            .field("current_version", &self.current_version)
            .field("key_versions", &versions)
            .finish()
    }
}

impl<'de> Deserialize<'de> for KeyRing {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        #[derive(Deserialize)]
        #[serde(rename_all = "camelCase")]
        struct RawKeyRing {
            current_version: u32,
            keys: BTreeMap<u32, KeyBytes>,
        }

        let raw = RawKeyRing::deserialize(deserializer)?;
        if raw.current_version < 1 {
            return Err(D::Error::custom("currentVersion must be >= 1"));
        }
        if raw.keys.is_empty() {
            return Err(D::Error::custom("keys must be non-empty"));
        }
        if raw.keys.keys().any(|v| *v < 1) {
            return Err(D::Error::custom("key versions must be >= 1"));
        }
        if !raw.keys.contains_key(&raw.current_version) {
            return Err(D::Error::custom("keys must include currentVersion"));
        }
        Ok(KeyRing {
            current_version: raw.current_version,
            keys: raw.keys,
        })
    }
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
    // Treat any `FMLENC??` prefix as encrypted so corrupted/truncated containers
    // are surfaced as encryption errors instead of being misinterpreted as a
    // plaintext SQLite file.
    if bytes.len() >= MAGIC_FMLENC_PREFIX.len() && bytes[..MAGIC_FMLENC_PREFIX.len()] == *MAGIC_FMLENC_PREFIX
    {
        return true;
    }
    bytes.len() >= LEGACY_MAGIC.len() && bytes[..LEGACY_MAGIC.len()] == *LEGACY_MAGIC
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
    let aad = aad_for_magic(MAGIC_FMLENC_V1, None, key_version);

    // Layout output as `[header][ciphertext]` so we can encrypt in-place without
    // allocating an intermediate ciphertext buffer (important for large workbooks).
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(HEADER_LEN_FMLENC_V1 + plaintext.len());
    out.resize(HEADER_LEN_FMLENC_V1, 0);
    out.extend_from_slice(plaintext);

    let tag = cipher.encrypt_in_place_detached(nonce, &aad, &mut out[HEADER_LEN_FMLENC_V1..])?;

    out[..8].copy_from_slice(MAGIC_FMLENC_V1);
    out[8..12].copy_from_slice(&key_version.to_be_bytes());
    out[12..24].copy_from_slice(&nonce_bytes);
    out[24..40].copy_from_slice(tag.as_slice());
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
    let aad = match parsed.format {
        ContainerFormat::Fmlenc { .. } => aad_for_magic(&parsed.magic, None, parsed.key_version),
        ContainerFormat::Legacy { version } => aad_for_magic(&parsed.magic, Some(version), parsed.key_version),
    };
    cipher.decrypt_in_place_detached(
        nonce,
        &aad,
        &mut buffer,
        Tag::from_slice(&parsed.tag),
    )?;
    Ok(buffer)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ContainerFormat {
    Fmlenc { version: u8 },
    Legacy { version: u8 },
}

#[derive(Debug)]
struct ParsedContainer<'a> {
    magic: [u8; 8],
    format: ContainerFormat,
    key_version: u32,
    nonce: [u8; NONCE_LEN],
    tag: [u8; TAG_LEN],
    ciphertext: &'a [u8],
}

fn parse_container(bytes: &[u8]) -> Result<ParsedContainer<'_>, EncryptionError> {
    let magic_slice = bytes.get(..8).ok_or(EncryptionError::TruncatedContainer)?;
    let magic: [u8; 8] = magic_slice
        .try_into()
        .map_err(|_| EncryptionError::TruncatedContainer)?;

    if magic == *MAGIC_FMLENC_V1 {
        if bytes.len() < HEADER_LEN_FMLENC_V1 {
            return Err(EncryptionError::TruncatedContainer);
        }
        let key_version_bytes = bytes.get(8..12).ok_or(EncryptionError::TruncatedContainer)?;
        let key_version_bytes: [u8; 4] = key_version_bytes
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;
        let key_version = u32::from_be_bytes(key_version_bytes);

        let nonce_slice = bytes.get(12..24).ok_or(EncryptionError::TruncatedContainer)?;
        let nonce: [u8; NONCE_LEN] = nonce_slice
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;

        let tag_slice = bytes.get(24..40).ok_or(EncryptionError::TruncatedContainer)?;
        let tag: [u8; TAG_LEN] = tag_slice
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;
        return Ok(ParsedContainer {
            magic,
            format: ContainerFormat::Fmlenc { version: 1 },
            key_version,
            nonce,
            tag,
            ciphertext: &bytes[HEADER_LEN_FMLENC_V1..],
        });
    }

    if magic == *LEGACY_MAGIC {
        if bytes.len() < HEADER_LEN_LEGACY_V1 {
            return Err(EncryptionError::TruncatedContainer);
        }
        let version = bytes[8];
        if version != LEGACY_CONTAINER_VERSION {
            return Err(EncryptionError::UnsupportedContainerVersion(version));
        }
        let key_version_bytes = bytes.get(9..13).ok_or(EncryptionError::TruncatedContainer)?;
        let key_version_bytes: [u8; 4] = key_version_bytes
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;
        let key_version = u32::from_be_bytes(key_version_bytes);

        let nonce_slice = bytes.get(13..25).ok_or(EncryptionError::TruncatedContainer)?;
        let nonce: [u8; NONCE_LEN] = nonce_slice
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;

        let tag_slice = bytes.get(25..41).ok_or(EncryptionError::TruncatedContainer)?;
        let tag: [u8; TAG_LEN] = tag_slice
            .try_into()
            .map_err(|_| EncryptionError::TruncatedContainer)?;
        return Ok(ParsedContainer {
            magic,
            format: ContainerFormat::Legacy { version },
            key_version,
            nonce,
            tag,
            ciphertext: &bytes[HEADER_LEN_LEGACY_V1..],
        });
    }

    if let Some(version) = parse_fmlenc_version(&magic) {
        return Err(EncryptionError::UnsupportedContainerVersion(version));
    }

    Err(EncryptionError::InvalidMagic)
}

fn parse_fmlenc_version(magic: &[u8; 8]) -> Option<u8> {
    if magic[..MAGIC_FMLENC_PREFIX.len()] != *MAGIC_FMLENC_PREFIX {
        return None;
    }
    let tens = magic[6];
    let ones = magic[7];
    if !tens.is_ascii_digit() || !ones.is_ascii_digit() {
        return None;
    }
    Some(((tens - b'0') * 10) + (ones - b'0'))
}

fn aad_for_magic(magic: &[u8; 8], version_byte: Option<u8>, key_version: u32) -> Vec<u8> {
    let mut aad = Vec::new();
    let _ = aad.try_reserve_exact(8 + version_byte.map_or(0, |_| 1) + 4);
    aad.extend_from_slice(magic);
    if let Some(version) = version_byte {
        aad.push(version);
    }
    aad.extend_from_slice(&key_version.to_be_bytes());
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
        assert_eq!(&encrypted[..8], MAGIC_FMLENC_V1);
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

    #[test]
    fn iv_tamper_detection_fails() {
        let keyring = KeyRing::from_key(1, [9u8; KEY_LEN]);
        let plaintext = b"more bytes";
        let mut encrypted = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt");

        // Flip a bit in the IV/nonce.
        encrypted[12] ^= 0b0000_0001;
        let err = decrypt_sqlite_bytes(&encrypted, &keyring).expect_err("decrypt should fail");
        match err {
            EncryptionError::Aead => {}
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn legacy_container_still_decrypts() {
        let key = [5u8; KEY_LEN];
        let keyring = KeyRing::from_key(1, key);
        let plaintext = b"legacy container bytes";

        let mut nonce_bytes = [0u8; NONCE_LEN];
        nonce_bytes[0] = 1;
        let nonce = Nonce::from_slice(&nonce_bytes);

        let cipher = Aes256Gcm::new(Key::<Aes256Gcm>::from_slice(&key));
        let mut buffer = plaintext.to_vec();
        let aad = aad_for_magic(LEGACY_MAGIC, Some(LEGACY_CONTAINER_VERSION), 1);
        let tag = cipher
            .encrypt_in_place_detached(nonce, &aad, &mut buffer)
            .expect("encrypt legacy");

        let mut container = Vec::new();
        container.extend_from_slice(LEGACY_MAGIC);
        container.push(LEGACY_CONTAINER_VERSION);
        container.extend_from_slice(&1u32.to_be_bytes());
        container.extend_from_slice(&nonce_bytes);
        container.extend_from_slice(tag.as_slice());
        container.extend_from_slice(&buffer);

        let decrypted = decrypt_sqlite_bytes(&container, &keyring).expect("decrypt legacy");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn unsupported_container_version_is_reported() {
        let bytes = b"FMLENC02".to_vec();
        let keyring = KeyRing::from_key(1, [0u8; KEY_LEN]);
        let err = decrypt_sqlite_bytes(&bytes, &keyring).expect_err("should fail");
        match err {
            EncryptionError::UnsupportedContainerVersion(2) => {}
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn tampered_key_version_fails_even_if_key_material_is_same() {
        let key = [7u8; KEY_LEN];
        let mut keyring = KeyRing::from_key(1, key);
        // Add a second version with identical key bytes.
        keyring.keys.insert(2, KeyBytes::new(key));

        let plaintext = b"header auth";
        let mut encrypted = encrypt_sqlite_bytes(plaintext, &keyring).expect("encrypt");
        assert_eq!(
            u32::from_be_bytes(encrypted[8..12].try_into().expect("key version bytes")),
            1
        );

        // Flip keyVersion from 1 -> 2 while keeping ciphertext/tag intact.
        encrypted[8..12].copy_from_slice(&2u32.to_be_bytes());

        let err = decrypt_sqlite_bytes(&encrypted, &keyring).expect_err("should fail");
        match err {
            EncryptionError::Aead => {}
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn keyring_json_is_compatible_with_padded_and_unpadded_base64() {
        let keyring = KeyRing::from_key(1, [0u8; KEY_LEN]);
        let mut json: serde_json::Value =
            serde_json::to_value(&keyring).expect("serialize keyring json");

        let key_str = json["keys"]["1"].as_str().expect("key string");
        assert!(
            key_str.ends_with('='),
            "expected padded base64 to match JS keyring encoding"
        );

        // Ensure we also accept unpadded base64 for backwards compatibility.
        let unpadded = key_str.trim_end_matches('=').to_string();
        json["keys"]["1"] = serde_json::Value::String(unpadded);
        let decoded: KeyRing = serde_json::from_value(json).expect("deserialize keyring json");
        assert_eq!(decoded.current_version, 1);
        assert_eq!(decoded.key(1).expect("key bytes"), [0u8; KEY_LEN]);
    }
}
