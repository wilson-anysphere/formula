#![cfg(target_arch = "wasm32")]

use formula_offcrypto::standard_decrypt_package;

/// Compile-only smoke test for `wasm32-unknown-unknown`.
///
/// CI can run:
/// `cargo test -p formula-offcrypto --target wasm32-unknown-unknown --no-run`
#[test]
fn wasm_compile_smoke() {
    // Deterministic, minimal "EncryptedPackage" buffer:
    // - total_size (u64 LE) = 0 so the decrypted output truncates to empty
    // - 16 bytes of ciphertext (required multiple-of-16)
    let key = [0u8; 16];
    let encrypted_package = [0u8; 8 + 16];

    let res = standard_decrypt_package(&key, &encrypted_package);
    assert!(res.is_ok());
}
