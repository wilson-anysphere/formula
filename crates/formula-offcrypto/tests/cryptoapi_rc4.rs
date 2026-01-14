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
fn cryptoapi_rc4_40bit_keys_are_padded_to_16_bytes() {
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
    assert_eq!(key.len(), 16);
    assert!(key[5..].iter().all(|b| *b == 0));
    assert_eq!(&key[..5], &hex_bytes("6ad7dedf2d")[..]);
}
