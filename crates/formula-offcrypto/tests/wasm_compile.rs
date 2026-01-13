#![cfg(target_arch = "wasm32")]

use formula_offcrypto::{parse_encrypted_package_header, parse_encryption_info};

/// Compile-only smoke test for `wasm32-unknown-unknown`.
///
/// CI compiles this via:
/// `cargo check -p formula-offcrypto --target wasm32-unknown-unknown --all-targets --locked`
///
/// Note: `cargo test -p formula-offcrypto --target wasm32-unknown-unknown` requires the
/// `wasm-bindgen-test-runner` binary; use `--no-run` for compile-only verification.
#[test]
fn wasm_compile_smoke() {
    // `EncryptedPackage` always begins with an 8-byte (u64 LE) original size.
    let encrypted_package_header = [0u8; 8];
    let _ = parse_encrypted_package_header(&encrypted_package_header);

    // `EncryptionInfo` starts with the 8-byte `EncryptionVersionInfo` header.
    // Use a version that is unsupported so we don't need to construct the rest of the structure.
    let encryption_info_header = [0u8; 8];
    let _ = parse_encryption_info(&encryption_info_header);
}
