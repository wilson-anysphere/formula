//! Desktop storage encryption hooks (Tauri/Rust).
//!
//! This repository is a reference implementation focused on policy enforcement
//! and crypto/key-management APIs. The production desktop app should implement
//! these hooks using a vetted Rust crypto library (e.g. `aes-gcm`) and a native
//! keychain integration (macOS Keychain, Windows DPAPI, Linux Secret Service).
//!
//! The JavaScript implementation in `encryptedDocumentStore.js` is used for
//! cross-platform tests in this repo.

/// Marker type for future integration.
pub struct DesktopStorageEncryption;

impl DesktopStorageEncryption {
    /// Enable encryption-at-rest for the on-disk store and migrate existing
    /// plaintext data to ciphertext.
    pub fn enable_encryption(&self) {
        unimplemented!("Implemented in the desktop app (Rust) in the real repo");
    }

    /// Disable encryption-at-rest and migrate ciphertext back to plaintext.
    pub fn disable_encryption(&self) {
        unimplemented!("Implemented in the desktop app (Rust) in the real repo");
    }
}

