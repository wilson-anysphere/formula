//! Regenerate the encrypted workbook fixtures under `fixtures/encryption/`.
//!
//! This is intentionally an *example* (not a library API) so its dependencies
//! remain `dev-dependencies` of `formula-office-crypto`.

use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use aes::cipher::{generic_array::GenericArray, BlockEncryptMut, KeyInit};
use rand09::RngCore;
use rc4::{consts::U16, Rc4, StreamCipher};
use sha1::{Digest as _, Sha1};

const PASSWORD_ASCII: &str = "password";
const PASSWORD_UNICODE: &str = "pässwörd";

// `office-crypto` decrypts Agile payloads in 4096-byte segments. Some versions panic when the
// plaintext package is smaller than one segment, so ensure our Agile fixtures are slightly larger
// than 4096 bytes.
const AGILE_MIN_PLAINTEXT_LEN: usize = 4096;

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .canonicalize()
        .expect("repo root")
}

fn read_file(path: &Path) -> Vec<u8> {
    fs::read(path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()))
}

fn write_file(path: &Path, bytes: &[u8]) {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .unwrap_or_else(|e| panic!("create dir {}: {e}", parent.display()));
    }
    fs::write(path, bytes).unwrap_or_else(|e| panic!("write {}: {e}", path.display()));
}

fn pad_zip_for_agile_min_len(zip_bytes: &[u8]) -> Vec<u8> {
    if zip_bytes.len() >= AGILE_MIN_PLAINTEXT_LEN {
        return zip_bytes.to_vec();
    }

    // Append a stored padding file without rewriting existing entries; this keeps the workbook
    // template intact while ensuring the overall ZIP size exceeds one Agile segment.
    let cursor = Cursor::new(zip_bytes.to_vec());
    let mut writer = zip::ZipWriter::new_append(cursor).expect("open zip for append");

    let pad_len = (AGILE_MIN_PLAINTEXT_LEN - zip_bytes.len()) + 256;
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    writer
        .start_file("xl/padding.bin", options)
        .expect("start padding file");
    let chunk = [0u8; 256];
    let mut remaining = pad_len;
    while remaining > 0 {
        let n = remaining.min(chunk.len());
        writer.write_all(&chunk[..n]).expect("write padding");
        remaining -= n;
    }

    let cursor = writer.finish().expect("finish padded zip");
    cursor.into_inner()
}

fn make_ole_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor).expect("create v3 cfb");

    for (name, data) in streams {
        let mut s = ole.create_stream(name).expect("create stream");
        s.write_all(data).expect("write stream");
    }

    ole.into_inner().into_inner()
}

fn make_ecma376_standard_key(password: &str, salt: &[u8; 16], key_bits: u32) -> Vec<u8> {
    const ITER_COUNT: u32 = 50000;

    let pass_utf16: Vec<u16> = password.encode_utf16().collect();
    let pass_utf16: &[u8] = unsafe {
        std::slice::from_raw_parts(pass_utf16.as_ptr() as *const u8, pass_utf16.len() * 2)
    };

    let mut h = Sha1::digest([salt.as_slice(), pass_utf16].concat());
    for i in 0..ITER_COUNT {
        h = Sha1::digest([&i.to_le_bytes(), h.as_slice()].concat());
    }

    let block_bytes = [0u8; 4];
    let h = Sha1::digest([h.as_slice(), &block_bytes].concat());

    let cb_required_key_length = (key_bits / 8) as usize;

    let mut buf1 = [0x36u8; 64];
    for (dst, src) in buf1.iter_mut().zip(h.iter()) {
        *dst ^= *src;
    }
    let x1 = Sha1::digest(buf1);

    let mut buf2 = [0x5cu8; 64];
    for (dst, src) in buf2.iter_mut().zip(h.iter()) {
        *dst ^= *src;
    }
    let x2 = Sha1::digest(buf2);

    [x1.as_slice(), x2.as_slice()].concat()[..cb_required_key_length].to_vec()
}

fn aes128_ecb_encrypt(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(key.len(), 16, "AES-128 key");
    assert!(
        plaintext.len() % 16 == 0,
        "ECB plaintext must be a multiple of 16 bytes"
    );
    let mut cipher = aes::Aes128::new_from_slice(key).expect("aes128 key");
    let mut out = plaintext.to_vec();
    for block in out.chunks_mut(16) {
        cipher.encrypt_block_mut(GenericArray::from_mut_slice(block));
    }
    out
}

fn encrypt_ooxml_standard(plaintext_zip: &[u8], password: &str) -> Vec<u8> {
    // Parameters chosen to match what `office-crypto` supports:
    // - AES-128
    // - SHA1 key derivation (ITER_COUNT=50000)
    // - EncryptedPackage AES-ECB with a 8-byte header (totalSize + reserved)
    let key_bits = 128u32;

    let mut rng = rand09::rng();

    let mut salt = [0u8; 16];
    rng.fill_bytes(&mut salt);
    let key = make_ecma376_standard_key(password, &salt, key_bits);

    // Verifier
    let mut verifier = [0u8; 16];
    rng.fill_bytes(&mut verifier);
    let verifier_hash = Sha1::digest(verifier);
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..verifier_hash.len()].copy_from_slice(verifier_hash.as_slice());

    let encrypted_verifier = aes128_ecb_encrypt(&key, &verifier);
    let encrypted_verifier_hash = aes128_ecb_encrypt(&key, &verifier_hash_padded);

    // Build EncryptionInfo (Standard)
    let csp_name = "Microsoft Enhanced RSA and AES Cryptographic Provider";
    let mut csp_utf16: Vec<u8> = csp_name
        .encode_utf16()
        .chain(std::iter::once(0)) // null terminator
        .flat_map(|u| u.to_le_bytes())
        .collect();
    // Excel writes an even number of u16s; ensure we don't accidentally omit the terminator.
    if csp_utf16.len() % 2 != 0 {
        csp_utf16.push(0);
    }

    let mut encryption_header = Vec::new();
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    encryption_header.extend_from_slice(&0x0000_660Eu32.to_le_bytes()); // algId (AES-128)
    encryption_header.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash (SHA1)
    encryption_header.extend_from_slice(&key_bits.to_le_bytes()); // keySize
    encryption_header.extend_from_slice(&0x0000_0018u32.to_le_bytes()); // providerType (PROV_RSA_AES)
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    encryption_header.extend_from_slice(&csp_utf16);

    let header_size = encryption_header.len() as u32;

    let mut encryption_info = Vec::new();
    // EncryptionVersionInfo (4.2 => Standard)
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&2u16.to_le_bytes());
    // headerFlags (unused by `office-crypto`).
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    // encryptionHeaderSize
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&encryption_header);
    // EncryptionVerifier
    encryption_info.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    encryption_info.extend_from_slice(&salt);
    encryption_info.extend_from_slice(&encrypted_verifier);
    encryption_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    encryption_info.extend_from_slice(&encrypted_verifier_hash);

    // Build EncryptedPackage
    let total_size = plaintext_zip.len() as u32;
    let mut padded = plaintext_zip.to_vec();
    let pad_len = (16 - (padded.len() % 16)) % 16;
    padded.extend(std::iter::repeat_n(0u8, pad_len));
    let ciphertext = aes128_ecb_encrypt(&key, &padded);

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&total_size.to_le_bytes());
    encrypted_package.extend_from_slice(&0u32.to_le_bytes()); // reserved
    encrypted_package.extend_from_slice(&ciphertext);

    make_ole_with_streams(&[
        ("EncryptionInfo", &encryption_info),
        ("EncryptedPackage", &encrypted_package),
    ])
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve(s.len().saturating_mul(2));
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn sha1_digest(chunks: &[&[u8]]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    out
}

fn derive_key_material_rc4_cryptoapi_sha1(password: &str, salt: &[u8; 16]) -> [u8; 20] {
    // CryptoAPI password hashing [MS-OFFCRYPTO] (SHA1):
    //   H0 = Hash(salt + UTF16LE(password))
    //   for i in 0..49999: H0 = Hash(i_le32 + H0)
    const ITERATIONS: u32 = 50_000;

    let pw_bytes = utf16le_bytes(password);
    let mut h = sha1_digest(&[salt.as_slice(), &pw_bytes]);
    for i in 0..ITERATIONS {
        let iter = i.to_le_bytes();
        h = sha1_digest(&[&iter, &h]);
    }
    h
}

fn derive_block_key_rc4_cryptoapi_sha1(
    key_material: &[u8; 20],
    key_bits: u32,
    block: u32,
) -> [u8; 16] {
    let block_bytes = block.to_le_bytes();
    let digest = sha1_digest(&[key_material, &block_bytes]);

    let key_len = (key_bits / 8) as usize;
    assert!(
        matches!(key_len, 5 | 16),
        "only 40-bit and 128-bit RC4 keys are supported by the fixture generator"
    );

    // [MS-OFFCRYPTO] calls out that 40-bit RC4 uses a 16-byte key where only the low 5 bytes are
    // set and the remaining bytes are 0.
    let mut key = [0u8; 16];
    if key_len == 5 {
        key[..5].copy_from_slice(&digest[..5]);
    } else {
        key.copy_from_slice(&digest[..16]);
    }
    key
}

struct PayloadRc4CryptoApi {
    key_material: [u8; 20],
    key_bits: u32,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4<U16>,
}

impl PayloadRc4CryptoApi {
    const BLOCK_SIZE: usize = 1024;

    fn new(key_material: [u8; 20], key_bits: u32) -> Self {
        let key = derive_block_key_rc4_cryptoapi_sha1(&key_material, key_bits, 0);
        let rc4 = Rc4::<U16>::new(key.as_slice().into());
        Self {
            key_material,
            key_bits,
            block: 0,
            pos_in_block: 0,
            rc4,
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key =
            derive_block_key_rc4_cryptoapi_sha1(&self.key_material, self.key_bits, self.block);
        self.rc4 = Rc4::<U16>::new(key.as_slice().into());
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        while !data.is_empty() {
            if self.pos_in_block == Self::BLOCK_SIZE {
                self.rekey();
            }

            let remaining = Self::BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

fn encrypt_biff_rc4_cryptoapi_workbook_stream(workbook_stream: &[u8], password: &str) -> Vec<u8> {
    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_BOUNDSHEET8: u16 = 0x0085;

    // Generate RC4 CryptoAPI verifier fields.
    let key_bits: u32 = 128;
    let mut rng = rand09::rng();

    let mut salt = [0u8; 16];
    rng.fill_bytes(&mut salt);

    let key_material = derive_key_material_rc4_cryptoapi_sha1(password, &salt);

    let mut verifier = [0u8; 16];
    rng.fill_bytes(&mut verifier);
    let verifier_hash = sha1_digest(&[&verifier]);

    // Encrypt verifier + hash using a single RC4 stream (block=0).
    let key0 = derive_block_key_rc4_cryptoapi_sha1(&key_material, key_bits, 0);
    let mut cipher = Rc4::<U16>::new(key0.as_slice().into());
    let mut encrypted_verifier = verifier;
    cipher.apply_keystream(&mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash.to_vec();
    cipher.apply_keystream(&mut encrypted_verifier_hash);

    // Build CryptoAPI EncryptionInfo (minimal, SHA1 + RC4).
    let csp_name = "Microsoft Enhanced Cryptographic Provider v1.0";
    let csp_utf16: Vec<u8> = csp_name
        .encode_utf16()
        .chain(std::iter::once(0))
        .flat_map(|u| u.to_le_bytes())
        .collect();

    let mut encryption_header = Vec::new();
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    encryption_header.extend_from_slice(&0x0000_6801u32.to_le_bytes()); // algId (RC4)
    encryption_header.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash (SHA1)
    encryption_header.extend_from_slice(&key_bits.to_le_bytes()); // keySize
    encryption_header.extend_from_slice(&0x0000_0001u32.to_le_bytes()); // providerType (PROV_RSA_FULL)
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    encryption_header.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    encryption_header.extend_from_slice(&csp_utf16);

    let mut enc_info = Vec::new();
    enc_info.extend_from_slice(&4u16.to_le_bytes()); // majorVersion (ignored by parser)
    enc_info.extend_from_slice(&2u16.to_le_bytes()); // minorVersion (ignored by parser)
    enc_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    enc_info.extend_from_slice(&(encryption_header.len() as u32).to_le_bytes()); // headerSize
    enc_info.extend_from_slice(&encryption_header);
    // EncryptionVerifier
    enc_info.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    enc_info.extend_from_slice(&salt);
    enc_info.extend_from_slice(&encrypted_verifier);
    enc_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    enc_info.extend_from_slice(&encrypted_verifier_hash);

    // Build FILEPASS record payload [MS-XLS 2.4.117] for RC4 CryptoAPI:
    // - u16 wEncryptionType (0x0001)
    // - u16 wEncryptionSubType (0x0002)
    // - u32 dwEncryptionInfoLen
    // - EncryptionInfo (dwEncryptionInfoLen bytes)
    let mut filepass_payload = Vec::new();
    filepass_payload.extend_from_slice(&0x0001u16.to_le_bytes()); // wEncryptionType (RC4)
    filepass_payload.extend_from_slice(&0x0002u16.to_le_bytes()); // wEncryptionSubType (CryptoAPI)
    filepass_payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
    filepass_payload.extend_from_slice(&enc_info);

    let filepass_len = filepass_payload.len();
    assert!(filepass_len <= u16::MAX as usize);

    let mut filepass_record = Vec::new();
    let _ = filepass_record.try_reserve_exact(4usize.saturating_add(filepass_len));
    filepass_record.extend_from_slice(&RECORD_FILEPASS.to_le_bytes());
    filepass_record.extend_from_slice(&(filepass_len as u16).to_le_bytes());
    filepass_record.extend_from_slice(&filepass_payload);

    // Insert FILEPASS after the first BOF record.
    let bof_len = {
        if workbook_stream.len() < 4 {
            panic!("workbook stream too short");
        }
        let record_id = u16::from_le_bytes([workbook_stream[0], workbook_stream[1]]);
        assert_eq!(record_id, RECORD_BOF_BIFF8, "expected BIFF8 BOF in fixture");
        let len = u16::from_le_bytes([workbook_stream[2], workbook_stream[3]]) as usize;
        4 + len
    };
    let delta = filepass_record.len() as u32;
    let mut with_filepass = Vec::new();
    with_filepass.extend_from_slice(&workbook_stream[..bof_len]);
    with_filepass.extend_from_slice(&filepass_record);
    with_filepass.extend_from_slice(&workbook_stream[bof_len..]);

    // Patch BoundSheet8.lbPlyPos offsets to account for the inserted FILEPASS record.
    let mut off = 0usize;
    while off + 4 <= with_filepass.len() {
        let rid = u16::from_le_bytes([with_filepass[off], with_filepass[off + 1]]);
        let len = u16::from_le_bytes([with_filepass[off + 2], with_filepass[off + 3]]) as usize;
        let data_start = off + 4;
        let data_end = data_start + len;
        if data_end > with_filepass.len() {
            break;
        }
        if rid == RECORD_BOUNDSHEET8 && len >= 4 {
            let orig = u32::from_le_bytes(
                with_filepass[data_start..data_start + 4]
                    .try_into()
                    .unwrap(),
            );
            let patched = orig.wrapping_add(delta);
            with_filepass[data_start..data_start + 4].copy_from_slice(&patched.to_le_bytes());
        }
        off = data_end;
    }

    // Some BIFF workbook streams include trailing padding bytes that do not form a complete record
    // header. Our masking/encryption logic operates on BIFF records, so truncate to the end of the
    // last complete record to keep offsets consistent.
    let mut record_end = 0usize;
    let mut off = 0usize;
    while off + 4 <= with_filepass.len() {
        let len = u16::from_le_bytes([with_filepass[off + 2], with_filepass[off + 3]]) as usize;
        let data_end = off + 4 + len;
        if data_end > with_filepass.len() {
            break;
        }
        record_end = data_end;
        off = data_end;
    }
    if record_end > 0 && record_end < with_filepass.len() {
        with_filepass.truncate(record_end);
    }

    // Encrypt record payloads after FILEPASS in the RC4 CryptoAPI "payload stream" mode:
    // record headers are plaintext and *do not* consume keystream.
    let filepass_end = bof_len + filepass_record.len();
    let mut payload_cipher = PayloadRc4CryptoApi::new(key_material, key_bits);
    let mut off = filepass_end;
    while off < with_filepass.len() {
        let remaining = with_filepass.len().saturating_sub(off);
        if remaining < 4 {
            break;
        }
        let len = u16::from_le_bytes([with_filepass[off + 2], with_filepass[off + 3]]) as usize;
        let data_start = off + 4;
        let data_end = data_start + len;
        if data_end > with_filepass.len() {
            break;
        }
        payload_cipher.apply_keystream(&mut with_filepass[data_start..data_end]);
        off = data_end;
    }

    with_filepass
}

fn encrypt_xls_rc4_cryptoapi(template_xls: &[u8], password: &str) -> Vec<u8> {
    let cursor = Cursor::new(template_xls);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open template xls");
    let mut stream = ole.open_stream("Workbook").expect("Workbook stream");
    let mut workbook_stream = Vec::new();
    stream
        .read_to_end(&mut workbook_stream)
        .expect("read Workbook");

    let encrypted_workbook_stream =
        encrypt_biff_rc4_cryptoapi_workbook_stream(&workbook_stream, password);

    make_ole_with_streams(&[("Workbook", &encrypted_workbook_stream)])
}

fn encrypt_ooxml_agile(plaintext_zip: &[u8], password: &str) -> Vec<u8> {
    // ms-offcrypto-writer produces an OLE file containing EncryptionInfo + EncryptedPackage.
    let cursor = Cursor::new(Vec::new());
    let mut rng = rand09::rng();
    let mut writer = ms_offcrypto_writer::Ecma376AgileWriter::create(&mut rng, password, cursor)
        .expect("create agile writer");
    writer
        .write_all(plaintext_zip)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    cursor.into_inner()
}

fn main() {
    let root = repo_root();
    let out_dir = root.join("fixtures").join("encryption");

    let template_xlsx = read_file(&root.join("fixtures/xlsx/basic/basic.xlsx"));
    let template_xlsx_agile = pad_zip_for_agile_min_len(&template_xlsx);
    // Use the small "basic.xlsm" fixture which contains a real VBA project that `formula-vba`
    // can parse (used by `crates/formula-xlsx/tests/roundtrip_macro.rs`).
    let template_xlsm = read_file(&root.join("fixtures/xlsx/macros/basic.xlsm"));
    let template_xlsm_agile = pad_zip_for_agile_min_len(&template_xlsm);
    let template_xls = read_file(&root.join("crates/formula-xls/tests/fixtures/basic.xls"));

    // OOXML: Agile
    write_file(
        &out_dir.join("encrypted_agile.xlsx"),
        &encrypt_ooxml_agile(&template_xlsx_agile, PASSWORD_ASCII),
    );
    write_file(
        &out_dir.join("encrypted_agile_unicode.xlsx"),
        &encrypt_ooxml_agile(&template_xlsx_agile, PASSWORD_UNICODE),
    );
    write_file(
        &out_dir.join("encrypted_agile.xlsm"),
        &encrypt_ooxml_agile(&template_xlsm_agile, PASSWORD_ASCII),
    );

    // OOXML: Standard
    write_file(
        &out_dir.join("encrypted_standard.xlsx"),
        &encrypt_ooxml_standard(&template_xlsx, PASSWORD_ASCII),
    );

    // Legacy XLS: RC4 CryptoAPI
    write_file(
        &out_dir.join("encrypted_rc4_cryptoapi.xls"),
        &encrypt_xls_rc4_cryptoapi(&template_xls, PASSWORD_ASCII),
    );
    write_file(
        &out_dir.join("encrypted_rc4_cryptoapi_unicode.xls"),
        &encrypt_xls_rc4_cryptoapi(&template_xls, PASSWORD_UNICODE),
    );

    println!("Wrote fixtures to {}", out_dir.display());
}
