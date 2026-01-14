//! End-to-end test for MS-OFFCRYPTO Standard / CryptoAPI / RC4 encrypted OOXML using **MD5**
//! (`EncryptionHeader.algIdHash = CALG_MD5`).
//!
//! This covers the high-level password open APIs (not just the standalone RC4 reader), ensuring we
//! dispatch `algIdHash` correctly when decrypting the Standard RC4 `EncryptedPackage` stream and
//! validating the verifier hash.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};

use formula_io::offcrypto::cryptoapi::{hash_password_fixed_spin, password_to_utf16le, HashAlg};
use formula_io::offcrypto::{CALG_MD5, CALG_RC4};
use formula_io::{open_workbook_with_password_and_preserved_ole, Error, Workbook};

const F_CRYPTOAPI: u32 = 0x0000_0004;

struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");
        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }
        let mut j: u8 = 0;
        for i in 0..256usize {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

fn simple_xlsx_bytes() -> Vec<u8> {
    let mut wb = formula_model::Workbook::new();
    wb.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&wb, &mut cursor).expect("write workbook");
    cursor.into_inner()
}

fn derive_rc4_key_md5(h: &[u8], block: u32, key_len: usize) -> Vec<u8> {
    use md5::Digest as _;

    assert!(key_len > 0, "RC4 key length must be non-zero");

    let mut hasher = md5::Md5::new();
    hasher.update(h);
    hasher.update(block.to_le_bytes());
    let digest = hasher.finalize();

    // MS-OFFCRYPTO Standard RC4 uses raw digest truncation (`keyLen = keySize/8`). For
    // `keySize == 0` (40-bit), that is a 5-byte key — *not* a 16-byte key padded with zeros.
    digest[..key_len].to_vec()
}

fn encrypt_rc4_cryptoapi_md5(h: &[u8], key_len: usize, plaintext: &[u8]) -> Vec<u8> {
    const BLOCK: usize = 0x200;

    let mut out = plaintext.to_vec();
    for (block_index, chunk) in out.chunks_mut(BLOCK).enumerate() {
        let key = derive_rc4_key_md5(h, block_index as u32, key_len);
        let mut rc4 = Rc4::new(&key);
        rc4.apply_keystream(chunk);
    }
    out
}

#[test]
fn open_workbook_with_password_decrypts_standard_cryptoapi_rc4_md5() {
    // Build a valid plaintext OOXML ZIP payload.
    let plaintext_xlsx = simple_xlsx_bytes();
    let password = "password";

    // Use MD5 (CALG_MD5) and a 40-bit key (represented as `keySize = 0` in the Standard header).
    let alg_id_hash = CALG_MD5;
    let hash_alg = HashAlg::from_calg_id(alg_id_hash).expect("HashAlg::from_calg_id(CALG_MD5)");
    let key_size_bits: u32 = 0;
    let key_len = 5usize; // 40-bit

    // Deterministic salt + verifier bytes so the ciphertext is stable.
    let salt: [u8; 16] = [
        0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
        0xDC, 0xFE,
    ];
    let verifier: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];

    // Derive the base hash H (fixed 50k spin).
    let pw_utf16le = password_to_utf16le(password);
    let h = hash_password_fixed_spin(&pw_utf16le, &salt, hash_alg);

    // Encrypt the verifier blob with block 0 RC4 key: verifier || Hash(verifier).
    let verifier_hash: [u8; 16] = {
        use md5::Digest as _;
        md5::Md5::digest(&verifier).into()
    };
    let mut verifier_blob = Vec::with_capacity(16 + 16);
    verifier_blob.extend_from_slice(&verifier);
    verifier_blob.extend_from_slice(&verifier_hash);
    let verifier_cipher = encrypt_rc4_cryptoapi_md5(&h, key_len, &verifier_blob);
    let encrypted_verifier = verifier_cipher[..16].to_vec();
    let encrypted_verifier_hash = verifier_cipher[16..].to_vec();

    // Encrypt the EncryptedPackage payload (does NOT include the 8-byte size prefix).
    let encrypted_payload = encrypt_rc4_cryptoapi_md5(&h, key_len, &plaintext_xlsx);
    let mut encrypted_package_stream = Vec::with_capacity(8 + encrypted_payload.len());
    encrypted_package_stream.extend_from_slice(&(plaintext_xlsx.len() as u64).to_le_bytes());
    encrypted_package_stream.extend_from_slice(&encrypted_payload);

    // Build a minimal Standard EncryptionInfo stream for RC4 + MD5.
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // VersionMajor
    encryption_info.extend_from_slice(&2u16.to_le_bytes()); // VersionMinor (Standard)
    // Standard/CryptoAPI EncryptionInfo commonly uses 0x0000_0040 for this outer flags field.
    // The critical bits for decryptors are in the inner `EncryptionHeader.flags`.
    encryption_info.extend_from_slice(&0x0000_0040u32.to_le_bytes()); // Flags

    let header_size: u32 = 32;
    encryption_info.extend_from_slice(&header_size.to_le_bytes());

    // EncryptionHeader (32 bytes fixed part).
    // MS-OFFCRYPTO Standard `EncryptionHeader.flags`:
    // - fCryptoAPI must be set for Standard/CryptoAPI encryption.
    // - fAES must be unset for RC4.
    encryption_info.extend_from_slice(&F_CRYPTOAPI.to_le_bytes()); // flags
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    encryption_info.extend_from_slice(&CALG_RC4.to_le_bytes()); // algId
    encryption_info.extend_from_slice(&alg_id_hash.to_le_bytes()); // algIdHash (MD5)
    encryption_info.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize (0 => 40-bit)
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // providerType
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    // EncryptionVerifier.
    encryption_info.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    encryption_info.extend_from_slice(&salt);
    encryption_info.extend_from_slice(&encrypted_verifier);
    encryption_info.extend_from_slice(&(verifier_hash.len() as u32).to_le_bytes()); // verifierHashSize
    encryption_info.extend_from_slice(&encrypted_verifier_hash);

    // Wrap in OLE container (EncryptionInfo + EncryptedPackage streams).
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo");
        s.write_all(&encryption_info)
            .expect("write EncryptionInfo bytes");
    }
    {
        let mut s = ole
            .create_stream("EncryptedPackage")
            .expect("create EncryptedPackage");
        s.write_all(&encrypted_package_stream)
            .expect("write EncryptedPackage bytes");
    }
    let ole_bytes = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard-rc4-md5.xlsx");
    std::fs::write(&path, &ole_bytes).expect("write encrypted file");

    let wrong = open_workbook_with_password_and_preserved_ole(&path, Some("wrong-password"));
    assert!(
        matches!(wrong, Err(Error::InvalidPassword { .. })),
        "wrong password should return InvalidPassword, got {wrong:?}"
    );

    let opened = open_workbook_with_password_and_preserved_ole(&path, Some(password))
        .expect("decrypt + open (preserving OLE streams)");
    assert!(
        opened.preserved_ole.is_some(),
        "expected encrypted OOXML open path to preserve OLE entries"
    );
    let Workbook::Xlsx(pkg) = opened.workbook else {
        panic!("expected Workbook::Xlsx, got {:?}", opened.workbook);
    };
    assert!(
        pkg.read_part("xl/workbook.xml")
            .expect("read xl/workbook.xml")
            .is_some(),
        "expected decrypted package to be a valid XLSX zip"
    );
}
