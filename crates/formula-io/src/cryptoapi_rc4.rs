use sha1::{Digest as _, Sha1};

/// Derive an RC4 key for "RC4 CryptoAPI" encryption (MS-OFFCRYPTO §2.3.5.2).
///
/// Notes on key length semantics:
/// - `key_size_bits` is read from `EncryptionHeader.KeySize` (bits).
/// - Per MS-OFFCRYPTO, `key_size_bits == 0` MUST be interpreted as 40-bit RC4.
/// - The RC4 key bytes passed into the cipher are exactly the first `key_size_bits/8` bytes of
///   `Hfinal` (no additional zero padding for 40-bit keys).
pub(crate) fn derive_rc4_cryptoapi_key(
    password: &str,
    salt: &[u8; 16],
    key_size_bits: u32,
    block: u32,
) -> Vec<u8> {
    // MS-OFFCRYPTO: if KeySize is 0, interpret as 40 bits.
    let key_size_bits = if key_size_bits == 0 {
        40
    } else {
        key_size_bits
    };
    assert!(
        key_size_bits >= 40 && key_size_bits <= 128 && key_size_bits % 8 == 0,
        "invalid RC4 CryptoAPI keySize={key_size_bits} (expected 40..=128, multiple of 8)"
    );

    // H0 = SHA1(salt + password)
    let mut h0 = Sha1::new();
    h0.update(salt);
    // Password is an array of Unicode characters; Office uses UTF-16LE bytes without a BOM.
    for unit in password.encode_utf16() {
        h0.update(unit.to_le_bytes());
    }
    let h0 = h0.finalize();

    // Hfinal = SHA1(H0 + block)
    let mut hfinal = Sha1::new();
    hfinal.update(&h0);
    hfinal.update(&block.to_le_bytes());
    let hfinal = hfinal.finalize();

    hfinal[..(key_size_bits as usize / 8)].to_vec()
}

#[cfg(test)]
mod tests {
    use super::derive_rc4_cryptoapi_key;

    #[test]
    fn rc4_cryptoapi_key_derivation_keysize_56_is_7_bytes_not_zero_padded() {
        // Test vector chosen to be deterministic and small.
        //
        // Computation (MS-OFFCRYPTO §2.3.5.2):
        //   H0     = SHA1(salt || UTF-16LE(password))
        //   Hfinal = SHA1(H0 || block_le32)
        //   key_bytes = Hfinal[0..keySize/8] (with keySize==0 interpreted as 40-bit => 5 bytes)
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let password = "VelvetSweatshop";
        let block = 0u32;

        // 40-bit: exactly 5 bytes (no padding to 16)
        let key40 = derive_rc4_cryptoapi_key(password, &salt, 40, block);
        assert_eq!(key40, vec![0xAC, 0x2A, 0x7B, 0x17, 0x24]);
        assert_eq!(key40.len(), 5);

        // 56-bit: exactly 7 bytes
        let key56 = derive_rc4_cryptoapi_key(password, &salt, 56, block);
        assert_eq!(key56, vec![0xAC, 0x2A, 0x7B, 0x17, 0x24, 0x74, 0xC7]);
        assert_eq!(key56.len(), 7);

        // 128-bit: 16 bytes
        let key128 = derive_rc4_cryptoapi_key(password, &salt, 128, block);
        assert_eq!(
            key128,
            vec![
                0xAC, 0x2A, 0x7B, 0x17, 0x24, 0x74, 0xC7, 0x9C, 0x0B, 0x12, 0x92, 0xE5,
                0x58, 0xDF, 0xD9, 0xB1
            ]
        );
    }

    #[test]
    fn rc4_cryptoapi_keysize_zero_means_40_bits() {
        let salt = [0u8; 16];
        let password = "p";
        let block = 0u32;
        let key0 = derive_rc4_cryptoapi_key(password, &salt, 0, block);
        let key40 = derive_rc4_cryptoapi_key(password, &salt, 40, block);
        assert_eq!(key0, key40);
    }
}
