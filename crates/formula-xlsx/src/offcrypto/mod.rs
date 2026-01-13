//! MS-OFFCRYPTO helpers.
//!
//! This module contains small, reusable primitives used by Office document encryption
//! implementations (e.g. Agile encryption in OOXML).

mod crypto;

pub use crypto::{
    derive_iv, derive_key, hash_password, CryptoError, HashAlgorithm, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};

