use std::io::{Cursor, Read};

use anyhow::{Context, Result};
use base64::Engine as _;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

// MS-OFFCRYPTO: password key derivation block keys (Agile encryption).
const BLOCK_KEY_VERIFIER_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLOCK_KEY_VERIFIER_HASH: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLOCK_KEY_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

const ENCRYPTED_PACKAGE_CHUNK_SIZE: usize = 4096;

#[derive(Debug, Copy, Clone)]
enum HashAlgorithm {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    fn digest(self, bytes: &[u8]) -> Vec<u8> {
        match self {
            HashAlgorithm::Sha1 => {
                use sha1::Digest as _;
                let mut hasher = sha1::Sha1::new();
                hasher.update(bytes);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha256 => {
                use sha2::Digest as _;
                let mut hasher = sha2::Sha256::new();
                hasher.update(bytes);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha384 => {
                use sha2::Digest as _;
                let mut hasher = sha2::Sha384::new();
                hasher.update(bytes);
                hasher.finalize().to_vec()
            }
            HashAlgorithm::Sha512 => {
                use sha2::Digest as _;
                let mut hasher = sha2::Sha512::new();
                hasher.update(bytes);
                hasher.finalize().to_vec()
            }
        }
    }
}

fn parse_hash_algorithm(name: &str) -> Result<HashAlgorithm> {
    let normalized = name.trim().to_ascii_uppercase().replace('-', "");
    match normalized.as_str() {
        "SHA1" => Ok(HashAlgorithm::Sha1),
        "SHA256" => Ok(HashAlgorithm::Sha256),
        "SHA384" => Ok(HashAlgorithm::Sha384),
        "SHA512" => Ok(HashAlgorithm::Sha512),
        other => anyhow::bail!("unsupported hash algorithm {other:?}"),
    }
}

#[derive(Debug)]
struct KeyData {
    block_size: usize,
    key_bits: usize,
    hash_algorithm: HashAlgorithm,
    salt_value: Vec<u8>,
}

#[derive(Debug)]
struct EncryptedKey {
    spin_count: u32,
    salt_size: usize,
    block_size: usize,
    key_bits: usize,
    hash_size: usize,
    hash_algorithm: HashAlgorithm,
    salt_value: Vec<u8>,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

#[derive(Debug)]
struct AgileEncryptionInfo {
    key_data: KeyData,
    encrypted_key: EncryptedKey,
}

pub(crate) fn is_encrypted_ooxml_ole(bytes: &[u8]) -> bool {
    if !bytes.starts_with(&OLE_MAGIC) {
        return false;
    }
    let cursor = Cursor::new(bytes.to_vec());
    let Ok(mut ole) = cfb::CompoundFile::open(cursor) else {
        return false;
    };
    stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")
}

pub(crate) fn decrypt_encrypted_ooxml_ole(bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    let cursor = Cursor::new(bytes.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).context("open OLE compound file")?;

    let encryption_info_bytes = read_stream(&mut ole, "EncryptionInfo")
        .context("read OLE stream EncryptionInfo")?;
    let encrypted_package_bytes = read_stream(&mut ole, "EncryptedPackage")
        .context("read OLE stream EncryptedPackage")?;

    let info = parse_agile_encryption_info(&encryption_info_bytes)?;
    decrypt_agile_encrypted_package(&info, &encrypted_package_bytes, password)
}

pub(crate) fn looks_like_xlsx(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || &bytes[..2] != b"PK" {
        return false;
    }
    let cursor = Cursor::new(bytes);
    let Ok(mut archive) = zip::ZipArchive::new(cursor) else {
        return false;
    };
    for i in 0..archive.len() {
        let Ok(file) = archive.by_index(i) else {
            continue;
        };
        if file.is_dir() {
            continue;
        }
        let name = file.name().trim_start_matches('/').replace('\\', "/");
        if name.eq_ignore_ascii_case("xl/workbook.xml") {
            return true;
        }
    }
    false
}

fn stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    ole.open_stream(&format!("/{name}")).is_ok()
}

fn read_stream<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Vec<u8>> {
    let mut stream = ole
        .open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .with_context(|| format!("open OLE stream {name}"))?;
    let mut out = Vec::new();
    stream
        .read_to_end(&mut out)
        .with_context(|| format!("read OLE stream {name}"))?;
    Ok(out)
}

fn parse_agile_encryption_info(encryption_info: &[u8]) -> Result<AgileEncryptionInfo> {
    if encryption_info.len() < 8 {
        anyhow::bail!("EncryptionInfo stream too short");
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    if major != 4 || minor != 4 {
        anyhow::bail!("unsupported EncryptionInfo version {major}.{minor} (expected 4.4)");
    }

    // Remaining bytes are an XML encryption descriptor (UTF-8), sometimes null-terminated.
    let mut xml_bytes = encryption_info[8..].to_vec();
    while matches!(xml_bytes.last(), Some(0)) {
        xml_bytes.pop();
    }
    let xml = std::str::from_utf8(&xml_bytes).context("decode EncryptionInfo XML as UTF-8")?;

    let doc = roxmltree::Document::parse(xml).context("parse EncryptionInfo XML")?;

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .context("EncryptionInfo missing <keyData>")?;

    let key_data_salt_size: usize = key_data_node
        .attribute("saltSize")
        .context("<keyData> missing saltSize")?
        .parse()
        .context("parse <keyData> saltSize")?;
    let key_data_salt_value = base64::engine::general_purpose::STANDARD
        .decode(
            key_data_node
                .attribute("saltValue")
                .context("<keyData> missing saltValue")?,
        )
        .context("decode <keyData> saltValue base64")?;
    if key_data_salt_value.len() != key_data_salt_size {
        anyhow::bail!(
            "<keyData> saltSize={key_data_salt_size} does not match decoded saltValue len={}",
            key_data_salt_value.len()
        );
    }

    let key_data = KeyData {
        block_size: key_data_node
            .attribute("blockSize")
            .context("<keyData> missing blockSize")?
            .parse()
            .context("parse <keyData> blockSize")?,
        key_bits: key_data_node
            .attribute("keyBits")
            .context("<keyData> missing keyBits")?
            .parse()
            .context("parse <keyData> keyBits")?,
        // Parse hashSize for validation (even though we don't currently use it for decryption).
        hash_algorithm: parse_hash_algorithm(
            key_data_node
                .attribute("hashAlgorithm")
                .context("<keyData> missing hashAlgorithm")?,
        )
        .context("parse <keyData> hashAlgorithm")?,
        salt_value: key_data_salt_value,
    };

    let _key_data_hash_size: usize = key_data_node
        .attribute("hashSize")
        .context("<keyData> missing hashSize")?
        .parse()
        .context("parse <keyData> hashSize")?;

    let key_data = KeyData {
        block_size: key_data.block_size,
        key_bits: key_data.key_bits,
        hash_algorithm: key_data.hash_algorithm,
        salt_value: key_data.salt_value,
    };

    let encrypted_key_node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyEncryptor"
                && n.attribute("uri")
                    .map(|uri| uri.contains("keyEncryptor/password"))
                    .unwrap_or(false)
        })
        .and_then(|node| {
            node.descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        })
        .context("EncryptionInfo missing password <encryptedKey>")?;

    let encrypted_key_salt_size: usize = encrypted_key_node
        .attribute("saltSize")
        .context("<encryptedKey> missing saltSize")?
        .parse()
        .context("parse <encryptedKey> saltSize")?;
    let encrypted_key_salt_value = base64::engine::general_purpose::STANDARD
        .decode(
            encrypted_key_node
                .attribute("saltValue")
                .context("<encryptedKey> missing saltValue")?,
        )
        .context("decode <encryptedKey> saltValue base64")?;
    if encrypted_key_salt_value.len() != encrypted_key_salt_size {
        anyhow::bail!(
            "<encryptedKey> saltSize={encrypted_key_salt_size} does not match decoded saltValue len={}",
            encrypted_key_salt_value.len()
        );
    }

    let encrypted_key_hash_size: usize = encrypted_key_node
        .attribute("hashSize")
        .context("<encryptedKey> missing hashSize")?
        .parse()
        .context("parse <encryptedKey> hashSize")?;

    let encrypted_key = EncryptedKey {
        spin_count: encrypted_key_node
            .attribute("spinCount")
            .context("<encryptedKey> missing spinCount")?
            .parse()
            .context("parse <encryptedKey> spinCount")?,
        salt_size: encrypted_key_salt_size,
        block_size: encrypted_key_node
            .attribute("blockSize")
            .context("<encryptedKey> missing blockSize")?
            .parse()
            .context("parse <encryptedKey> blockSize")?,
        key_bits: encrypted_key_node
            .attribute("keyBits")
            .context("<encryptedKey> missing keyBits")?
            .parse()
            .context("parse <encryptedKey> keyBits")?,
        hash_size: encrypted_key_hash_size,
        hash_algorithm: parse_hash_algorithm(
            encrypted_key_node
                .attribute("hashAlgorithm")
                .context("<encryptedKey> missing hashAlgorithm")?,
        )
        .context("parse <encryptedKey> hashAlgorithm")?,
        salt_value: encrypted_key_salt_value,
        encrypted_verifier_hash_input: base64::engine::general_purpose::STANDARD
            .decode(
                encrypted_key_node
                    .attribute("encryptedVerifierHashInput")
                    .context("<encryptedKey> missing encryptedVerifierHashInput")?,
            )
            .context("decode encryptedVerifierHashInput base64")?,
        encrypted_verifier_hash_value: base64::engine::general_purpose::STANDARD
            .decode(
                encrypted_key_node
                    .attribute("encryptedVerifierHashValue")
                    .context("<encryptedKey> missing encryptedVerifierHashValue")?,
            )
            .context("decode encryptedVerifierHashValue base64")?,
        encrypted_key_value: base64::engine::general_purpose::STANDARD
            .decode(
                encrypted_key_node
                    .attribute("encryptedKeyValue")
                    .context("<encryptedKey> missing encryptedKeyValue")?,
            )
            .context("decode encryptedKeyValue base64")?,
    };

    Ok(AgileEncryptionInfo {
        key_data,
        encrypted_key,
    })
}

fn decrypt_agile_encrypted_package(
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    if encrypted_package.len() < 8 {
        anyhow::bail!("EncryptedPackage stream too short");
    }

    let package_size = u64::from_le_bytes(encrypted_package[..8].try_into().unwrap()) as usize;
    let ciphertext = &encrypted_package[8..];

    let encrypted_key = &info.encrypted_key;
    let password_bytes = password_utf16le(password);

    // Password key derivation (Agile encryption).
    let mut h = {
        let mut buf = Vec::with_capacity(encrypted_key.salt_value.len() + password_bytes.len());
        buf.extend_from_slice(&encrypted_key.salt_value);
        buf.extend_from_slice(&password_bytes);
        encrypted_key.hash_algorithm.digest(&buf)
    };

    let mut spin_buf = Vec::with_capacity(4 + h.len());
    for i in 0..encrypted_key.spin_count {
        spin_buf.clear();
        spin_buf.extend_from_slice(&(i as u32).to_le_bytes());
        spin_buf.extend_from_slice(&h);
        h = encrypted_key.hash_algorithm.digest(&spin_buf);
    }

    let verifier_hash_input = {
        let key = derive_key(
            encrypted_key.hash_algorithm,
            &h,
            &BLOCK_KEY_VERIFIER_INPUT,
            encrypted_key.key_bits,
        )?;
        let iv = derive_iv(
            encrypted_key.hash_algorithm,
            &encrypted_key.salt_value,
            &BLOCK_KEY_VERIFIER_INPUT,
            encrypted_key.block_size,
        );
        let decrypted = aes_cbc_decrypt(&encrypted_key.encrypted_verifier_hash_input, &key, &iv)?;
        decrypted
            .get(..encrypted_key.salt_size)
            .context("decrypted verifier hash input shorter than saltSize")?
            .to_vec()
    };

    let verifier_hash_value = {
        let key = derive_key(
            encrypted_key.hash_algorithm,
            &h,
            &BLOCK_KEY_VERIFIER_HASH,
            encrypted_key.key_bits,
        )?;
        let iv = derive_iv(
            encrypted_key.hash_algorithm,
            &encrypted_key.salt_value,
            &BLOCK_KEY_VERIFIER_HASH,
            encrypted_key.block_size,
        );
        let decrypted = aes_cbc_decrypt(&encrypted_key.encrypted_verifier_hash_value, &key, &iv)?;
        decrypted
            .get(..encrypted_key.hash_size)
            .context("decrypted verifier hash value shorter than hashSize")?
            .to_vec()
    };

    let computed_hash = encrypted_key.hash_algorithm.digest(&verifier_hash_input);
    if computed_hash
        .get(..encrypted_key.hash_size)
        .context("hash output shorter than hashSize")?
        != verifier_hash_value
    {
        anyhow::bail!("invalid password");
    }

    let key_value = {
        let key = derive_key(
            encrypted_key.hash_algorithm,
            &h,
            &BLOCK_KEY_KEY_VALUE,
            encrypted_key.key_bits,
        )?;
        let iv = derive_iv(
            encrypted_key.hash_algorithm,
            &encrypted_key.salt_value,
            &BLOCK_KEY_KEY_VALUE,
            encrypted_key.block_size,
        );
        let decrypted = aes_cbc_decrypt(&encrypted_key.encrypted_key_value, &key, &iv)?;
        let expected_len = info.key_data.key_bits / 8;
        decrypted
            .get(..expected_len)
            .context("decrypted key value shorter than keyBits/8")?
            .to_vec()
    };

    // Decrypt the package stream in 4096-byte chunks using per-chunk IVs derived from keyData.
    let key_data = &info.key_data;
    if key_data.block_size != 16 {
        anyhow::bail!("unsupported keyData blockSize {} (expected 16)", key_data.block_size);
    }
    if key_value.len() * 8 != key_data.key_bits {
        anyhow::bail!(
            "unexpected decrypted key length {} for keyBits {}",
            key_value.len(),
            key_data.key_bits
        );
    }

    let mut plaintext = Vec::with_capacity(package_size);
    for (chunk_index, chunk) in ciphertext.chunks(ENCRYPTED_PACKAGE_CHUNK_SIZE).enumerate() {
        if chunk.is_empty() {
            continue;
        }
        let block_key = (chunk_index as u32).to_le_bytes();
        let iv = derive_iv(
            key_data.hash_algorithm,
            &key_data.salt_value,
            &block_key,
            key_data.block_size,
        );
        let decrypted = aes_cbc_decrypt(chunk, &key_value, &iv)?;
        plaintext.extend_from_slice(&decrypted);
    }
    plaintext.truncate(package_size);
    Ok(plaintext)
}

fn password_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn derive_key(
    hash_algorithm: HashAlgorithm,
    base_hash: &[u8],
    block_key: &[u8],
    key_bits: usize,
) -> Result<Vec<u8>> {
    let key_len = key_bits / 8;
    if key_len == 0 {
        anyhow::bail!("invalid keyBits {key_bits}");
    }

    let mut buf = Vec::with_capacity(base_hash.len() + block_key.len());
    buf.extend_from_slice(base_hash);
    buf.extend_from_slice(block_key);
    let mut digest = hash_algorithm.digest(&buf);
    if digest.len() < key_len {
        digest.resize(key_len, 0);
    } else {
        digest.truncate(key_len);
    }
    Ok(digest)
}

fn derive_iv(
    hash_algorithm: HashAlgorithm,
    salt: &[u8],
    block_key: &[u8],
    block_size: usize,
) -> Vec<u8> {
    let mut buf = Vec::with_capacity(salt.len() + block_key.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(block_key);
    let mut digest = hash_algorithm.digest(&buf);
    if digest.len() < block_size {
        digest.resize(block_size, 0);
    } else {
        digest.truncate(block_size);
    }
    digest
}

fn aes_cbc_decrypt(ciphertext: &[u8], key: &[u8], iv: &[u8]) -> Result<Vec<u8>> {
    use aes::cipher::{BlockDecrypt, KeyInit};
    use aes::cipher::generic_array::GenericArray;

    if ciphertext.len() % 16 != 0 {
        anyhow::bail!("ciphertext length {} is not a multiple of 16", ciphertext.len());
    }
    if iv.len() != 16 {
        anyhow::bail!("IV length {} is not 16", iv.len());
    }

    let mut out = Vec::with_capacity(ciphertext.len());
    let mut prev = [0u8; 16];
    prev.copy_from_slice(iv);

    match key.len() {
        16 => {
            let cipher = aes::Aes128::new(GenericArray::from_slice(key));
            for block in ciphertext.chunks_exact(16) {
                let mut buf = [0u8; 16];
                buf.copy_from_slice(block);
                let ga = GenericArray::from_mut_slice(&mut buf);
                cipher.decrypt_block(ga);
                for i in 0..16 {
                    buf[i] ^= prev[i];
                }
                prev.copy_from_slice(block);
                out.extend_from_slice(&buf);
            }
        }
        32 => {
            let cipher = aes::Aes256::new(GenericArray::from_slice(key));
            for block in ciphertext.chunks_exact(16) {
                let mut buf = [0u8; 16];
                buf.copy_from_slice(block);
                let ga = GenericArray::from_mut_slice(&mut buf);
                cipher.decrypt_block(ga);
                for i in 0..16 {
                    buf[i] ^= prev[i];
                }
                prev.copy_from_slice(block);
                out.extend_from_slice(&buf);
            }
        }
        other => anyhow::bail!("unsupported AES key size {other}"),
    }

    Ok(out)
}

