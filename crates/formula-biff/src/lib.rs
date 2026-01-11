//! BIFF12 formula token helpers.
//!
//! This crate provides a small subset of Excel's BIFF12 `rgce` formula token
//! stream:
//! - `decode_rgce`: best-effort decoding of `rgce` into Excel formula text
//! - `encode_rgce` (feature `encode`): encoding of formula text into `rgce`
//!
//! The encoder is intentionally scoped to the initial editing workflows:
//! constants, A1-style refs, basic operators, and a curated set of built-in
//! functions.

mod ftab;
mod function_ids;
mod rgce;

pub use ftab::{function_id_from_name, function_name_from_id, FTAB_USER_DEFINED};
pub use function_ids::{function_id_to_name, function_name_to_id, function_spec_from_id};
pub use rgce::{decode_rgce, DecodeRgceError};

#[cfg(feature = "encode")]
pub use rgce::{encode_rgce, EncodeRgceError};
