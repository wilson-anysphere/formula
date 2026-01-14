use formula_offcrypto::cryptoapi;
use formula_offcrypto::HashAlgorithm;

fn hex_bytes(s: &str) -> Vec<u8> {
    hex::decode(s).expect("valid hex")
}

#[test]
fn cryptoapi_iterated_hash_and_block_hash_vectors() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    // SHA1 vectors (spinCount = 50,000; password="password"; salt=00..0f)
    let expected_h_sha1 = hex_bytes("1b5972284eab6481eb6565a0985b334b3e65e041"); // 20 bytes
    let expected_block_sha1 = [
        "6ad7dedf2da3514b1d85eabee069d47dd058967f",
        "2ed4e8825cd48aa4a47994cda7415b4a9687377d",
        "9ce57d0699be3938951f47fa949361dbe64fdbbc",
        "e65b2643eaba3815a37a61159f13784085a577e3",
    ];

    let h_sha1 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("derive SHA1 iterated hash");
    assert_eq!(h_sha1.as_slice(), expected_h_sha1.as_slice());

    for (block, expected) in expected_block_sha1.iter().enumerate() {
        let hfinal = cryptoapi::block_hash(h_sha1.as_slice(), block as u32, HashAlgorithm::Sha1)
            .expect("SHA1 block hash");
        assert_eq!(
            hfinal.as_slice(),
            hex_bytes(expected).as_slice(),
            "SHA1 block hash mismatch for block {block}"
        );
    }

    // MD5 vectors (same parameters, but Hash=MD5)
    let expected_h_md5 = hex_bytes("2079476089fda784c3a3cfeb98102c7e"); // 16 bytes
    let expected_block_md5 = [
        "69badcae244868e209d4e053ccd2a3bc",
        "6f4d502ab37700ffdab5704160455b47",
        "ac69022e396c7750872133f37e2c7afc",
        "1b056e7118ab8d35e9d67adee8b11104",
    ];

    let h_md5 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Md5,
    )
    .expect("derive MD5 iterated hash");
    assert_eq!(h_md5.as_slice(), expected_h_md5.as_slice());

    for (block, expected) in expected_block_md5.iter().enumerate() {
        let hfinal = cryptoapi::block_hash(h_md5.as_slice(), block as u32, HashAlgorithm::Md5)
            .expect("MD5 block hash");
        assert_eq!(
            hfinal.as_slice(),
            hex_bytes(expected).as_slice(),
            "MD5 block hash mismatch for block {block}"
        );
    }
}

#[test]
fn cryptoapi_rc4_key_is_raw_truncation_not_cryptderivekey() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    let h_sha1 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("derive SHA1 iterated hash");

    // For keySize=128, raw-truncation keys are the first 16 bytes of Hfinal.
    let expected_hfinal_prefix_sha1 = [
        "6ad7dedf2da3514b1d85eabee069d47d",
        "2ed4e8825cd48aa4a47994cda7415b4a",
        "9ce57d0699be3938951f47fa949361db",
        "e65b2643eaba3815a37a61159f137840",
    ];
    // If we (incorrectly) applied the CryptoAPI `CryptDeriveKey` ipad/opad transform to Hfinal, the
    // first 16 bytes differ.
    let expected_cryptderivekey_sha1 = [
        "de5451b9dc3fcb383792cbeec80b6bc3",
        "f270061ed91886704b71823ba68dd311",
        "e047f75f017c50d27ac7dec8223b3749",
        "943a5534e626685ba225840572576676",
    ];

    for block in 0u32..4 {
        let rc4_key =
            cryptoapi::rc4_key_for_block(h_sha1.as_slice(), block, 128, HashAlgorithm::Sha1)
                .expect("rc4 key");
        assert_eq!(
            rc4_key.as_slice(),
            hex_bytes(expected_hfinal_prefix_sha1[block as usize]).as_slice(),
            "RC4 key mismatch for block {block}"
        );

        let hfinal =
            cryptoapi::block_hash(h_sha1.as_slice(), block, HashAlgorithm::Sha1).expect("hfinal");
        let derived = cryptoapi::crypt_derive_key(hfinal.as_slice(), 128, HashAlgorithm::Sha1)
            .expect("cryptderivekey");
        assert_eq!(
            derived.as_slice(),
            hex_bytes(expected_cryptderivekey_sha1[block as usize]).as_slice(),
            "CryptDeriveKey mismatch for block {block}"
        );
        assert_ne!(
            rc4_key.as_slice(),
            derived.as_slice(),
            "RC4 key must not use CryptDeriveKey transform"
        );
    }

    let h_md5 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Md5,
    )
    .expect("derive MD5 iterated hash");

    let expected_hfinal_md5 = [
        "69badcae244868e209d4e053ccd2a3bc",
        "6f4d502ab37700ffdab5704160455b47",
        "ac69022e396c7750872133f37e2c7afc",
        "1b056e7118ab8d35e9d67adee8b11104",
    ];
    let expected_cryptderivekey_md5 = [
        "8d666ec55103fdbdc3281cc271f6cb7c",
        "892b60ddd451139fed758f20fe5d1be0",
        "d9034198455f9bd171ad16d04cea4c42",
        "06f5756e6e23c795cd6786f5dd565830",
    ];

    for block in 0u32..4 {
        let rc4_key =
            cryptoapi::rc4_key_for_block(h_md5.as_slice(), block, 128, HashAlgorithm::Md5)
                .expect("rc4 key (md5)");
        assert_eq!(
            rc4_key.as_slice(),
            hex_bytes(expected_hfinal_md5[block as usize]).as_slice(),
            "MD5 RC4 key mismatch for block {block}"
        );

        let hfinal =
            cryptoapi::block_hash(h_md5.as_slice(), block, HashAlgorithm::Md5).expect("hfinal");
        let derived = cryptoapi::crypt_derive_key(hfinal.as_slice(), 128, HashAlgorithm::Md5)
            .expect("cryptderivekey (md5)");
        assert_eq!(
            derived.as_slice(),
            hex_bytes(expected_cryptderivekey_md5[block as usize]).as_slice(),
            "MD5 CryptDeriveKey mismatch for block {block}"
        );
        assert_ne!(
            rc4_key.as_slice(),
            derived.as_slice(),
            "RC4 key must not use CryptDeriveKey transform (MD5)"
        );
    }
}

#[test]
fn cryptoapi_rc4_40bit_keys_truncate_to_5_bytes() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    let h = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("iterated hash");

    let key =
        cryptoapi::rc4_key_for_block(h.as_slice(), 0, 40, HashAlgorithm::Sha1).expect("40-bit key");
    assert_eq!(key.len(), 5);
    assert_eq!(key.as_slice(), hex_bytes("6ad7dedf2d").as_slice());
}

#[test]
fn cryptoapi_rc4_56bit_keys_truncate_to_7_bytes() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    let h = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("iterated hash");

    let key =
        cryptoapi::rc4_key_for_block(h.as_slice(), 0, 56, HashAlgorithm::Sha1).expect("56-bit key");
    assert_eq!(key.len(), 7);
    assert_eq!(key.as_slice(), hex_bytes("6ad7dedf2da351").as_slice());
}

#[test]
fn cryptoapi_rc4_keysize_zero_is_interpreted_as_40bit() {
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();

    let h = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("iterated hash");

    let key0 = cryptoapi::rc4_key_for_block(h.as_slice(), 0, 0, HashAlgorithm::Sha1)
        .expect("keySize=0 must be accepted");
    let key40 = cryptoapi::rc4_key_for_block(h.as_slice(), 0, 40, HashAlgorithm::Sha1)
        .expect("40-bit key");
    assert_eq!(key0.as_slice(), key40.as_slice());
    assert_eq!(key0.len(), 5);
    assert_eq!(key0.as_slice(), hex_bytes("6ad7dedf2d").as_slice());
}

#[test]
fn cryptoapi_rc4_40bit_padding_affects_ciphertext_vectors() {
    // Deterministic vectors from `docs/offcrypto-standard-cryptoapi-rc4.md`:
    // password="password", salt=00..0f, spin=50,000, block=0, plaintext="Hello, RC4 CryptoAPI!"
    let password = "password";
    let salt: Vec<u8> = (0u8..16).collect();
    let plaintext = b"Hello, RC4 CryptoAPI!";

    // --- SHA1 ----------------------------------------------------------------
    let h_sha1 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .expect("derive SHA1 iterated hash");

    // Expected ciphertext when using the **unpadded** 5-byte key (`keyLen = keySize/8`).
    let ciphertext_sha1_unpadded = hex_bytes("d1fa444913b4839b06eb4851750a07761005f025bf");
    let decrypted_sha1 = cryptoapi::rc4_decrypt_stream(
        &ciphertext_sha1_unpadded,
        h_sha1.as_slice(),
        40,
        HashAlgorithm::Sha1,
    )
    .expect("decrypt SHA1 unpadded");
    assert_eq!(decrypted_sha1.as_slice(), plaintext);

    // keySize=0 must be interpreted as 40-bit.
    let decrypted_sha1_keysize0 = cryptoapi::rc4_decrypt_stream(
        &ciphertext_sha1_unpadded,
        h_sha1.as_slice(),
        0,
        HashAlgorithm::Sha1,
    )
    .expect("decrypt SHA1 unpadded (keySize=0)");
    assert_eq!(decrypted_sha1_keysize0.as_slice(), plaintext);

    // Regression guard: if we (incorrectly) zero-padded the 5-byte key material to 16 bytes, we'd
    // produce/decrypt a different ciphertext.
    let ciphertext_sha1_padded = hex_bytes("7a8bd000713a6e30ba9916476d27b01d36707a6ef8");
    let decrypted_sha1_wrong = cryptoapi::rc4_decrypt_stream(
        &ciphertext_sha1_padded,
        h_sha1.as_slice(),
        40,
        HashAlgorithm::Sha1,
    )
    .expect("decrypt SHA1 padded ciphertext");
    assert_ne!(decrypted_sha1_wrong.as_slice(), plaintext);

    // --- MD5 -----------------------------------------------------------------
    let h_md5 = cryptoapi::iterated_hash_from_password(
        password,
        &salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        HashAlgorithm::Md5,
    )
    .expect("derive MD5 iterated hash");

    let ciphertext_md5_unpadded = hex_bytes("db037cd60d38c882019b5f5d8c43382373f476da28");
    let decrypted_md5 = cryptoapi::rc4_decrypt_stream(
        &ciphertext_md5_unpadded,
        h_md5.as_slice(),
        40,
        HashAlgorithm::Md5,
    )
    .expect("decrypt MD5 unpadded");
    assert_eq!(decrypted_md5.as_slice(), plaintext);

    let decrypted_md5_keysize0 = cryptoapi::rc4_decrypt_stream(
        &ciphertext_md5_unpadded,
        h_md5.as_slice(),
        0,
        HashAlgorithm::Md5,
    )
    .expect("decrypt MD5 unpadded (keySize=0)");
    assert_eq!(decrypted_md5_keysize0.as_slice(), plaintext);

    let ciphertext_md5_padded = hex_bytes("565016a3d8158632bb36ce1d76996628512061bfa3");
    let decrypted_md5_wrong = cryptoapi::rc4_decrypt_stream(
        &ciphertext_md5_padded,
        h_md5.as_slice(),
        40,
        HashAlgorithm::Md5,
    )
    .expect("decrypt MD5 padded ciphertext");
    assert_ne!(decrypted_md5_wrong.as_slice(), plaintext);
}
