//! Decrypt an OOXML `EncryptedPackage` container (password-protected `.xlsx` / `.xlsm` / `.xlsb`).
//!
//! Encrypted OOXML files are **not ZIP files on disk** even if they use a `.xlsx` extension. Excel
//! wraps the real ZIP/OPC package in an OLE/CFB container with (at least) two streams:
//!
//! - `EncryptionInfo`
//! - `EncryptedPackage`
//!
//! This example reads those streams, prints a one-line `EncryptionInfo` summary to stderr, and
//! writes the decrypted ZIP bytes to a file or stdout.
//!
//! ## Usage
//!
//! ```bash
//! # Print help
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- --help
//!
//! # Decrypt to a file
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'correct horse battery staple' --output book.zip
//!
//! # Decrypt to stdout (useful for piping)
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'pw' > book.zip
//!
//! # (Agile) Verify the `dataIntegrity` HMAC as well
//! cargo run -p formula-offcrypto --example decrypt_ooxml -- \
//!   --input book.xlsx --password 'pw' --verify-integrity > book.zip
//! ```
//!
//! The output is a ZIP file; you can inspect it with `unzip -l book.zip`.

use std::ffi::OsString;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::PathBuf;

use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use formula_offcrypto::{
    decrypt_encrypted_package, inspect_encryption_info, parse_encryption_info,
    standard_derive_key_zeroizing, standard_verify_key, AgileEncryptionInfo, EncryptionInfo,
    HashAlgorithm, OffcryptoError, StandardEncryptionInfo,
};
use hmac::{Hmac, Mac as _};
use sha1::Digest as _;
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

fn main() {
    let args = match Args::parse() {
        Ok(args) => args,
        Err(ParseOutcome::Help(msg)) => {
            print!("{msg}");
            return;
        }
        Err(ParseOutcome::Error(msg)) => {
            eprintln!("{msg}");
            std::process::exit(2);
        }
    };

    let mut file = match File::open(&args.input) {
        Ok(f) => f,
        Err(err) => {
            eprintln!("error: failed to open {}: {err}", args.input.display());
            std::process::exit(1);
        }
    };

    let mut ole = match cfb::CompoundFile::open(&mut file) {
        Ok(ole) => ole,
        Err(err) => {
            eprintln!(
                "error: failed to parse OLE/CFB compound file {}: {err}",
                args.input.display()
            );
            std::process::exit(1);
        }
    };

    let encryption_info_bytes = match read_stream_best_effort(&mut ole, "EncryptionInfo") {
        Ok(b) => b,
        Err(err) => {
            eprintln!("error: failed to read EncryptionInfo stream: {err}");
            std::process::exit(1);
        }
    };
    let encrypted_package_bytes = match read_stream_best_effort(&mut ole, "EncryptedPackage") {
        Ok(b) => b,
        Err(err) => {
            eprintln!("error: failed to read EncryptedPackage stream: {err}");
            std::process::exit(1);
        }
    };

    match inspect_encryption_info(&encryption_info_bytes) {
        Ok(summary) => eprintln!("EncryptionInfo: {summary:?}"),
        Err(err) => eprintln!("warning: failed to inspect EncryptionInfo: {err}"),
    }

    let decrypted_zip = match parse_encryption_info(&encryption_info_bytes) {
        Ok(EncryptionInfo::Standard {
            header, verifier, ..
        }) => {
            let info = StandardEncryptionInfo { header, verifier };
            match decrypt_standard_encrypted_package(
                &info,
                &encrypted_package_bytes,
                &args.password,
            ) {
                Ok(b) => b,
                Err(err) => {
                    eprintln!("error: failed to decrypt Standard EncryptedPackage: {err}");
                    std::process::exit(1);
                }
            }
        }
        Ok(EncryptionInfo::Agile { info, .. }) => {
            match decrypt_agile_encrypted_package(
                &info,
                &encrypted_package_bytes,
                &args.password,
                args.verify_integrity,
            ) {
                Ok(b) => b,
                Err(err) => {
                    eprintln!("error: failed to decrypt Agile EncryptedPackage: {err}");
                    std::process::exit(1);
                }
            }
        }
        Ok(EncryptionInfo::Unsupported { version }) => {
            eprintln!(
                "error: unsupported EncryptionInfo version {}.{}",
                version.major, version.minor
            );
            std::process::exit(1);
        }
        Err(err) => {
            eprintln!("error: failed to parse EncryptionInfo: {err}");
            std::process::exit(1);
        }
    };

    if let Some(out_path) = &args.output {
        if let Err(err) = std::fs::write(out_path, &decrypted_zip) {
            eprintln!("error: failed to write {}: {err}", out_path.display());
            std::process::exit(1);
        }
    } else {
        let mut stdout = std::io::stdout().lock();
        if let Err(err) = stdout.write_all(&decrypted_zip) {
            eprintln!("error: failed to write decrypted bytes to stdout: {err}");
            std::process::exit(1);
        }
    }
}

struct Args {
    input: PathBuf,
    password: String,
    verify_integrity: bool,
    output: Option<PathBuf>,
}

enum ParseOutcome {
    Help(String),
    Error(String),
}

impl Args {
    fn parse() -> Result<Self, ParseOutcome> {
        let mut input: Option<PathBuf> = None;
        let mut password: Option<String> = None;
        let mut verify_integrity = false;
        let mut output: Option<PathBuf> = None;

        let mut argv = std::env::args_os();
        let exe = argv
            .next()
            .unwrap_or_else(|| OsString::from("decrypt_ooxml"));

        while let Some(arg) = argv.next() {
            match arg.to_string_lossy().as_ref() {
                "-h" | "--help" => {
                    return Err(ParseOutcome::Help(Self::help(&exe)));
                }
                "--input" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --input requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    input = Some(PathBuf::from(v));
                }
                "--password" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --password requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    password = Some(v.to_string_lossy().to_string());
                }
                "--verify-integrity" => {
                    verify_integrity = true;
                }
                "--output" => {
                    let Some(v) = argv.next() else {
                        return Err(ParseOutcome::Error(format!(
                            "error: --output requires a value\n\n{}",
                            Self::help(&exe)
                        )));
                    };
                    output = Some(PathBuf::from(v));
                }
                other => {
                    return Err(ParseOutcome::Error(format!(
                        "error: unrecognized argument `{other}`\n\n{}",
                        Self::help(&exe)
                    )));
                }
            }
        }

        let input = input.ok_or_else(|| {
            ParseOutcome::Error(format!(
                "error: missing required --input\n\n{}",
                Self::help(&exe)
            ))
        })?;
        let password = password.ok_or_else(|| {
            ParseOutcome::Error(format!(
                "error: missing required --password\n\n{}",
                Self::help(&exe)
            ))
        })?;

        Ok(Self {
            input,
            password,
            verify_integrity,
            output,
        })
    }

    fn help(exe: &OsString) -> String {
        let exe = exe.to_string_lossy();
        format!(
            "Usage: {exe} --input <path> --password <pw> [--verify-integrity] [--output <path>]\n\
             \n\
             Decrypt an OOXML encrypted container (OLE/CFB with EncryptionInfo + EncryptedPackage).\n\
             \n\
             Options:\n\
               --input <path>           Path to the encrypted OLE/CFB file (.xlsx/.xlsm/.xlsb)\n\
               --password <pw>          Password to open the workbook\n\
               --verify-integrity       (Agile) verify dataIntegrity HMAC\n\
               --output <path>          Write decrypted ZIP bytes to a file (defaults to stdout)\n\
               -h, --help               Print help\n"
        )
    }
}

fn read_stream_best_effort<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Vec<u8>, std::io::Error> {
    let mut stream = open_stream_best_effort(ole, name)?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf)?;
    Ok(buf)
}

fn open_stream_best_effort<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, std::io::Error> {
    let want = name.trim_start_matches('/');

    if let Ok(s) = ole.open_stream(want) {
        return Ok(s);
    }
    let with_leading_slash = format!("/{want}");
    if let Ok(s) = ole.open_stream(&with_leading_slash) {
        return Ok(s);
    }

    // Case-insensitive fallback: walk the directory tree and match stream paths.
    let mut found_path: Option<String> = None;
    let mut found_normalized: Option<String> = None;
    for entry in ole.walk() {
        if !entry.is_stream() {
            continue;
        }
        let path = entry.path().to_string_lossy().to_string();
        let normalized = path.trim_start_matches('/').to_string();
        if normalized.eq_ignore_ascii_case(want) {
            found_path = Some(path);
            found_normalized = Some(normalized);
            break;
        }
    }

    if let Some(normalized) = found_normalized {
        if let Ok(s) = ole.open_stream(&normalized) {
            return Ok(s);
        }
        let with_slash = format!("/{normalized}");
        if let Ok(s) = ole.open_stream(&with_slash) {
            return Ok(s);
        }
        if let Some(path) = found_path {
            if let Ok(s) = ole.open_stream(&path) {
                return Ok(s);
            }
        }
    }

    Err(std::io::Error::new(
        std::io::ErrorKind::NotFound,
        format!("stream not found: `{want}`"),
    ))
}

// --- Agile encryption constants (MS-OFFCRYPTO) ---------------------------------------------------

const VERIFIER_HASH_INPUT_BLOCK: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const VERIFIER_HASH_VALUE_BLOCK: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const KEY_VALUE_BLOCK: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
const HMAC_KEY_BLOCK: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
const HMAC_VALUE_BLOCK: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

fn decrypt_standard_encrypted_package(
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let key = standard_derive_key_zeroizing(info, password)?;
    standard_verify_key(info, &key)?;

    let mut iv_seed = Vec::with_capacity(info.verifier.salt.len() + 4);
    decrypt_encrypted_package(encrypted_package, |segment_index, ciphertext, plaintext| {
        iv_seed.clear();
        iv_seed.extend_from_slice(&info.verifier.salt);
        iv_seed.extend_from_slice(&segment_index.to_le_bytes());

        let digest = sha1::Sha1::digest(&iv_seed);
        let mut iv = [0u8; 16];
        iv.copy_from_slice(&digest[..16]);

        plaintext.copy_from_slice(ciphertext);
        aes_cbc_decrypt_in_place(&key, &iv, plaintext)
    })
}

fn decrypt_agile_encrypted_package(
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
    verify_integrity: bool,
) -> Result<Vec<u8>, OffcryptoError> {
    // 1) Derive the iterated password hash H.
    let h = derive_iterated_hash_from_password(
        password,
        &info.password_salt,
        info.password_hash_algorithm,
        info.spin_count,
    );

    // 2) Decrypt/verifier check.
    let verifier_hash_input = decrypt_agile_value(
        &h,
        &info.password_salt,
        info.password_hash_algorithm,
        info.password_key_bits,
        &VERIFIER_HASH_INPUT_BLOCK,
        &info.encrypted_verifier_hash_input,
    )?;
    let verifier_hash_value = decrypt_agile_value(
        &h,
        &info.password_salt,
        info.password_hash_algorithm,
        info.password_key_bits,
        &VERIFIER_HASH_VALUE_BLOCK,
        &info.encrypted_verifier_hash_value,
    )?;
    verify_agile_password(
        &verifier_hash_input,
        &verifier_hash_value,
        info.password_hash_algorithm,
    )?;

    // 3) Decrypt keyValue (the package AES key).
    let key_value = decrypt_agile_value(
        &h,
        &info.password_salt,
        info.password_hash_algorithm,
        info.password_key_bits,
        &KEY_VALUE_BLOCK,
        &info.encrypted_key_value,
    )?;
    let key_len =
        info.password_key_bits
            .checked_div(8)
            .ok_or(OffcryptoError::InvalidEncryptionInfo {
                context: "encryptedKey.keyBits must be a multiple of 8",
            })?;
    if key_value.len() < key_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile decrypted keyValue is too short",
        });
    }
    let secret_key = &key_value[..key_len];

    // 4) Decrypt the segmented package ciphertext.
    let mut iv_seed = Vec::with_capacity(info.key_data_salt.len() + 4);
    let decrypted =
        decrypt_encrypted_package(encrypted_package, |segment_index, ciphertext, plaintext| {
            iv_seed.clear();
            iv_seed.extend_from_slice(&info.key_data_salt);
            iv_seed.extend_from_slice(&segment_index.to_le_bytes());

            let digest = hash_alg_digest(info.key_data_hash_algorithm, &iv_seed);
            let mut iv = [0u8; 16];
            iv.copy_from_slice(&digest[..16]);

            plaintext.copy_from_slice(ciphertext);
            aes_cbc_decrypt_in_place(secret_key, &iv, plaintext)
        })?;

    // 5) Optional `dataIntegrity` verification.
    if verify_integrity {
        verify_agile_data_integrity(info, secret_key, &decrypted)?;
    }

    Ok(decrypted)
}

fn verify_agile_password(
    verifier_hash_input: &[u8],
    verifier_hash_value: &[u8],
    hash_alg: HashAlgorithm,
) -> Result<(), OffcryptoError> {
    let expected = Zeroizing::new(hash_alg_digest(hash_alg, verifier_hash_input));
    if verifier_hash_value.len() < expected.len() {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile verifierHashValue is too short",
        });
    }
    if !bool::from(verifier_hash_value[..expected.len()].ct_eq(&expected[..])) {
        return Err(OffcryptoError::InvalidPassword);
    }
    Ok(())
}

fn verify_agile_data_integrity(
    info: &AgileEncryptionInfo,
    secret_key: &[u8],
    decrypted_package: &[u8],
) -> Result<(), OffcryptoError> {
    let key_bits = info.password_key_bits;

    let iv = iv_from_salt_16(&info.key_data_salt)?;

    let hmac_key_encryption_key = derive_encryption_key(
        secret_key,
        &HMAC_KEY_BLOCK,
        info.key_data_hash_algorithm,
        key_bits,
    )?;
    let hmac_key_buf = aes_cbc_decrypt(&info.encrypted_hmac_key, &hmac_key_encryption_key, &iv)?;

    let digest_len = hash_alg_digest_len(info.key_data_hash_algorithm);
    if hmac_key_buf.len() < digest_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile decrypted HMAC key is too short",
        });
    }
    let hmac_key = &hmac_key_buf[..digest_len];

    let hmac_value_encryption_key = derive_encryption_key(
        secret_key,
        &HMAC_VALUE_BLOCK,
        info.key_data_hash_algorithm,
        key_bits,
    )?;
    let hmac_value_buf =
        aes_cbc_decrypt(&info.encrypted_hmac_value, &hmac_value_encryption_key, &iv)?;
    if hmac_value_buf.len() < digest_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile decrypted HMAC value is too short",
        });
    }
    let expected_hmac = &hmac_value_buf[..digest_len];

    let computed =
        Zeroizing::new(compute_hmac(info.key_data_hash_algorithm, hmac_key, decrypted_package)?);
    if !bool::from(computed.as_slice().ct_eq(expected_hmac)) {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile dataIntegrity HMAC mismatch",
        });
    }
    Ok(())
}

fn decrypt_agile_value(
    h: &[u8],
    salt: &[u8],
    hash_alg: HashAlgorithm,
    key_bits: usize,
    block_key: &[u8],
    ciphertext: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    let iv = iv_from_salt_16(salt)?;
    let key = derive_encryption_key(h, block_key, hash_alg, key_bits)?;
    aes_cbc_decrypt(ciphertext, &key, &iv)
}

fn derive_iterated_hash_from_password(
    password: &str,
    salt: &[u8],
    hash_alg: HashAlgorithm,
    spin: u32,
) -> Vec<u8> {
    let pw_utf16 = password_to_utf16le_bytes(password);

    let mut buf = Vec::with_capacity(salt.len() + pw_utf16.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&pw_utf16);
    let mut h = hash_alg_digest(hash_alg, &buf);

    let mut tmp = Vec::new();
    for i in 0..spin {
        tmp.clear();
        tmp.extend_from_slice(&i.to_le_bytes());
        tmp.extend_from_slice(&h);
        h = hash_alg_digest(hash_alg, &tmp);
    }

    h
}

fn derive_encryption_key(
    h: &[u8],
    block_key: &[u8],
    hash_alg: HashAlgorithm,
    key_bits: usize,
) -> Result<Vec<u8>, OffcryptoError> {
    if key_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "keyBits must be a multiple of 8",
        });
    }
    let key_len = key_bits / 8;

    let mut buf = Vec::with_capacity(h.len() + block_key.len());
    buf.extend_from_slice(h);
    buf.extend_from_slice(block_key);
    let mut out = hash_alg_digest(hash_alg, &buf);

    if key_len <= out.len() {
        out.truncate(key_len);
    } else {
        out.resize(key_len, 0);
    }

    Ok(out)
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn iv_from_salt_16(salt: &[u8]) -> Result<[u8; 16], OffcryptoError> {
    if salt.len() < 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "saltValue is too short for AES-CBC IV",
        });
    }
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&salt[..16]);
    Ok(iv)
}

fn hash_alg_digest_len(hash_alg: HashAlgorithm) -> usize {
    match hash_alg {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    }
}

fn hash_alg_digest(hash_alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match hash_alg {
        HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

fn aes_cbc_decrypt(
    ciphertext: &[u8],
    key: &[u8],
    iv: &[u8; 16],
) -> Result<Vec<u8>, OffcryptoError> {
    let mut buf = ciphertext.to_vec();
    aes_cbc_decrypt_in_place(key, iv, &mut buf)?;
    Ok(buf)
}

fn aes_cbc_decrypt_in_place(
    key: &[u8],
    iv: &[u8; 16],
    buf: &mut [u8],
) -> Result<(), OffcryptoError> {
    if buf.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength { len: buf.len() });
    }
    let len = buf.len();

    match key.len() {
        16 => {
            let decryptor = Decryptor::<Aes128>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        24 => {
            let decryptor = Decryptor::<Aes192>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        32 => {
            let decryptor = Decryptor::<Aes256>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    Ok(())
}

fn compute_hmac(
    hash_alg: HashAlgorithm,
    key: &[u8],
    data: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    let out = match hash_alg {
        HashAlgorithm::Sha1 => {
            let mut mac = <Hmac<sha1::Sha1> as hmac::Mac>::new_from_slice(key).map_err(|_| {
                OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid HMAC key",
                }
            })?;
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha256 => {
            let mut mac = <Hmac<sha2::Sha256> as hmac::Mac>::new_from_slice(key).map_err(|_| {
                OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid HMAC key",
                }
            })?;
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha384 => {
            let mut mac = <Hmac<sha2::Sha384> as hmac::Mac>::new_from_slice(key).map_err(|_| {
                OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid HMAC key",
                }
            })?;
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
        HashAlgorithm::Sha512 => {
            let mut mac = <Hmac<sha2::Sha512> as hmac::Mac>::new_from_slice(key).map_err(|_| {
                OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid HMAC key",
                }
            })?;
            mac.update(data);
            mac.finalize().into_bytes().to_vec()
        }
    };
    Ok(out)
}
