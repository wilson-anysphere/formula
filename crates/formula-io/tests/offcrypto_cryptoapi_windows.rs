#![cfg(windows)]

use formula_io::offcrypto::cryptoapi::{crypt_derive_key, HashAlg};
use sha1::{Digest as _, Sha1};
use windows_sys::Win32::Foundation::GetLastError;
use windows_sys::Win32::Security::Cryptography::{
    CryptAcquireContextW, CryptCreateHash, CryptDeriveKey, CryptDestroyHash, CryptDestroyKey,
    CryptExportKey, CryptGetHashParam, CryptGetKeyParam, CryptHashData, CryptReleaseContext,
    CALG_AES_128,
    CALG_AES_192, CALG_AES_256, CALG_MD5, CALG_SHA1, CRYPT_EXPORTABLE, CRYPT_VERIFYCONTEXT,
    HP_HASHVAL, KP_KEYVAL, PROV_RSA_AES,
};

// https://learn.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
const NTE_BAD_ALGID: u32 = 0x8009_0008;

fn last_err() -> u32 {
    // SAFETY: safe FFI call.
    unsafe { GetLastError() }
}

#[derive(Debug)]
struct CryptoProvider(usize);

impl CryptoProvider {
    fn acquire() -> Self {
        let mut hprov = 0usize;
        // SAFETY: FFI call; passing null container/provider to request an ephemeral context.
        let ok = unsafe {
            CryptAcquireContextW(
                &mut hprov,
                std::ptr::null(),
                std::ptr::null(),
                PROV_RSA_AES,
                CRYPT_VERIFYCONTEXT,
            )
        };
        assert_ne!(
            ok, 0,
            "CryptAcquireContextW(PROV_RSA_AES, CRYPT_VERIFYCONTEXT) failed: {}",
            last_err()
        );
        Self(hprov)
    }
}

impl Drop for CryptoProvider {
    fn drop(&mut self) {
        // SAFETY: FFI call; ignore failures in Drop to avoid panics during unwinding.
        unsafe {
            CryptReleaseContext(self.0, 0);
        }
    }
}

#[derive(Debug)]
struct CryptoHash(usize);

impl CryptoHash {
    fn try_new(provider: &CryptoProvider, alg_id_hash: u32) -> Result<Self, u32> {
        let mut hhash = 0usize;
        // SAFETY: FFI call.
        let ok = unsafe { CryptCreateHash(provider.0, alg_id_hash, 0, 0, &mut hhash) };
        if ok == 0 {
            return Err(last_err());
        }
        Ok(Self(hhash))
    }

    fn new(provider: &CryptoProvider, alg_id_hash: u32) -> Self {
        Self::try_new(provider, alg_id_hash).unwrap_or_else(|err| {
            panic!("CryptCreateHash(alg_id_hash={alg_id_hash:#x}) failed: {err}")
        })
    }

    fn hash_data(&mut self, data: &[u8]) {
        // SAFETY: FFI call.
        let ok = unsafe { CryptHashData(self.0, data.as_ptr(), data.len() as u32, 0) };
        assert_ne!(ok, 0, "CryptHashData failed: {}", last_err());
    }

    fn get_hash_value(&self) -> Vec<u8> {
        let mut len: u32 = 0;
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetHashParam(self.0, HP_HASHVAL, std::ptr::null_mut(), &mut len, 0) };
        assert_ne!(
            ok, 0,
            "CryptGetHashParam(HP_HASHVAL, size query) failed: {}",
            last_err()
        );
        let mut buf = vec![0u8; len as usize];
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetHashParam(self.0, HP_HASHVAL, buf.as_mut_ptr(), &mut len, 0) };
        assert_ne!(
            ok, 0,
            "CryptGetHashParam(HP_HASHVAL) failed: {}",
            last_err()
        );
        buf.truncate(len as usize);
        buf
    }
}

impl Drop for CryptoHash {
    fn drop(&mut self) {
        // SAFETY: FFI call; ignore failures in Drop to avoid panics during unwinding.
        unsafe {
            CryptDestroyHash(self.0);
        }
    }
}

#[derive(Debug)]
struct CryptoKey(usize);

impl CryptoKey {
    fn try_derive(
        provider: &CryptoProvider,
        alg_id_key: u32,
        hash: &CryptoHash,
    ) -> Result<Self, u32> {
        let mut hkey = 0usize;
        // SAFETY: FFI call.
        let ok = unsafe { CryptDeriveKey(provider.0, alg_id_key, hash.0, CRYPT_EXPORTABLE, &mut hkey) };
        if ok == 0 {
            return Err(last_err());
        }
        Ok(Self(hkey))
    }

    fn derive(provider: &CryptoProvider, alg_id_key: u32, hash: &CryptoHash) -> Self {
        Self::try_derive(provider, alg_id_key, hash).unwrap_or_else(|err| {
            panic!("CryptDeriveKey(alg_id_key={alg_id_key:#x}) failed: {err}")
        })
    }

    fn get_key_value(&self) -> Vec<u8> {
        match self.get_key_value_kp_keyval() {
            Ok(buf) => buf,
            Err(kp_err) => {
                // Some CryptoAPI providers reject KP_KEYVAL even for exportable session keys.
                // Fall back to exporting a PLAINTEXTKEYBLOB.
                self.export_key_plaintext_blob().unwrap_or_else(|export_err| {
                    panic!(
                        "failed to extract key bytes: KP_KEYVAL error={kp_err} CryptExportKey error={export_err}"
                    )
                })
            }
        }
    }

    fn get_key_value_kp_keyval(&self) -> Result<Vec<u8>, u32> {
        let mut len: u32 = 0;
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetKeyParam(self.0, KP_KEYVAL, std::ptr::null_mut(), &mut len, 0) };
        if ok == 0 {
            return Err(last_err());
        }
        let mut buf = vec![0u8; len as usize];
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetKeyParam(self.0, KP_KEYVAL, buf.as_mut_ptr(), &mut len, 0) };
        if ok == 0 {
            return Err(last_err());
        }
        buf.truncate(len as usize);
        Ok(buf)
    }

    fn export_key_plaintext_blob(&self) -> Result<Vec<u8>, u32> {
        // `PLAINTEXTKEYBLOB` is not exported by windows-sys, so define the WinCrypt constant here.
        const PLAINTEXTKEYBLOB: u32 = 0x8;

        let mut len: u32 = 0;
        // SAFETY: FFI call.
        let ok = unsafe {
            CryptExportKey(
                self.0,
                0, // no exchange key needed for PLAINTEXTKEYBLOB
                PLAINTEXTKEYBLOB,
                0,
                std::ptr::null_mut(),
                &mut len,
            )
        };
        if ok == 0 {
            return Err(last_err());
        }

        let mut buf = vec![0u8; len as usize];
        // SAFETY: FFI call.
        let ok = unsafe {
            CryptExportKey(
                self.0,
                0,
                PLAINTEXTKEYBLOB,
                0,
                buf.as_mut_ptr(),
                &mut len,
            )
        };
        if ok == 0 {
            return Err(last_err());
        }
        buf.truncate(len as usize);
        Ok(buf)
    }
}

impl Drop for CryptoKey {
    fn drop(&mut self) {
        // SAFETY: FFI call; ignore failures in Drop to avoid panics during unwinding.
        unsafe {
            CryptDestroyKey(self.0);
        }
    }
}

fn extract_session_key_bytes(kp_keyval: &[u8], expected_len: usize, expected_alg_id: u32) -> Vec<u8> {
    // Some providers return raw key bytes, others return a PLAINTEXTKEYBLOB:
    //   BLOBHEADER (8) || DWORD key_bits (4) || key_bytes
    if kp_keyval.len() == expected_len {
        return kp_keyval.to_vec();
    }

    const PLAINTEXTKEYBLOB: u8 = 0x8;
    const CUR_BLOB_VERSION: u8 = 0x2;

    if kp_keyval.len() == 12 + expected_len
        && kp_keyval.first().copied() == Some(PLAINTEXTKEYBLOB)
        && kp_keyval.get(1).copied() == Some(CUR_BLOB_VERSION)
    {
        let reserved = u16::from_le_bytes([kp_keyval[2], kp_keyval[3]]);
        assert_eq!(
            reserved, 0,
            "KP_KEYVAL PLAINTEXTKEYBLOB reserved field must be 0 (got {reserved})"
        );
        let alg_id = u32::from_le_bytes([
            kp_keyval[4],
            kp_keyval[5],
            kp_keyval[6],
            kp_keyval[7],
        ]);
        assert_eq!(
            alg_id, expected_alg_id,
            "KP_KEYVAL PLAINTEXTKEYBLOB aiKeyAlg mismatch (expected {expected_alg_id:#x}, got {alg_id:#x})"
        );
        let key_bits = u32::from_le_bytes([
            kp_keyval[8],
            kp_keyval[9],
            kp_keyval[10],
            kp_keyval[11],
        ]);
        assert_eq!(
            key_bits as usize,
            expected_len * 8,
            "KP_KEYVAL PLAINTEXTKEYBLOB key_bits mismatch (expected {} bits, got {key_bits})",
            expected_len * 8
        );
        return kp_keyval[12..].to_vec();
    }

    panic!(
        "unexpected KP_KEYVAL format: len={} expected_key_len={expected_len}",
        kp_keyval.len()
    );
}

fn derive_key_and_hash(
    provider: &CryptoProvider,
    alg_id_hash: u32,
    alg_id_key: u32,
    data: &[u8],
) -> (Vec<u8>, Vec<u8>) {
    let mut hash = CryptoHash::new(provider, alg_id_hash);
    hash.hash_data(data);
    let hash_value = hash.get_hash_value();

    let key = CryptoKey::derive(provider, alg_id_key, &hash);
    let key_bytes = key.get_key_value();
    (hash_value, key_bytes)
}

fn try_derive_key_and_hash(
    provider: &CryptoProvider,
    alg_id_hash: u32,
    alg_id_key: u32,
    data: &[u8],
) -> Result<(Vec<u8>, Vec<u8>), u32> {
    let mut hash = CryptoHash::try_new(provider, alg_id_hash)?;
    hash.hash_data(data);
    let hash_value = hash.get_hash_value();

    let key = CryptoKey::try_derive(provider, alg_id_key, &hash)?;
    let key_bytes = key.get_key_value();
    Ok((hash_value, key_bytes))
}

#[test]
fn cryptderivekey_matches_cryptoapi_for_sha1_aes_key_sizes() {
    let provider = CryptoProvider::acquire();
    let alg_id_hash = CALG_SHA1;
    let hash_alg = HashAlg::from_calg_id(alg_id_hash).unwrap();

    // Arbitrary-but-stable input for the hash object.
    let data = b"formula-io offcrypto CryptDeriveKey cross-check";

    for (alg_id_key, key_len) in [
        (CALG_AES_256, 32usize),
        (CALG_AES_192, 24usize),
        (CALG_AES_128, 16usize),
    ] {
        let (hash_value, key_blob) = derive_key_and_hash(&provider, alg_id_hash, alg_id_key, data);
        let expected_hash_value: [u8; 20] = Sha1::digest(data).into();
        assert_eq!(
            hash_value,
            expected_hash_value,
            "unexpected SHA-1 hash value from CryptoAPI (hash derivation mismatch)"
        );
        let key_bytes = extract_session_key_bytes(&key_blob, key_len, alg_id_key);

        let ours = crypt_derive_key(&hash_value, key_len, hash_alg);
        assert_eq!(
            key_bytes, ours,
            "derived key mismatch for alg_id_key={alg_id_key:#x} key_len={key_len}"
        );
    }
}

#[test]
fn cryptderivekey_matches_cryptoapi_for_md5_aes_key_sizes() {
    let provider = CryptoProvider::acquire();
    let alg_id_hash = CALG_MD5;
    let hash_alg = HashAlg::from_calg_id(alg_id_hash).unwrap();

    // Arbitrary-but-stable input for the hash object.
    let data = b"formula-io offcrypto CryptDeriveKey cross-check";

    for (alg_id_key, key_len) in [
        (CALG_AES_256, 32usize),
        (CALG_AES_192, 24usize),
        (CALG_AES_128, 16usize),
    ] {
        let (hash_value, key_blob) =
            match try_derive_key_and_hash(&provider, alg_id_hash, alg_id_key, data) {
                Ok(v) => v,
                Err(err) if err == NTE_BAD_ALGID => {
                    // Some environments/providers disable MD5 (e.g. via FIPS policy). Skip the
                    // cross-check in that case; the library's MD5 implementation is still covered
                    // by deterministic unit tests.
                    eprintln!("skipping MD5 CryptDeriveKey cross-check: CALG_MD5 unsupported (err={err})");
                    return;
                }
                Err(err) => panic!(
                    "CryptDeriveKey cross-check failed for alg_id_key={alg_id_key:#x} key_len={key_len}: {err}"
                ),
            };
        let key_bytes = extract_session_key_bytes(&key_blob, key_len, alg_id_key);

        let ours = crypt_derive_key(&hash_value, key_len, hash_alg);
        assert_eq!(
            key_bytes, ours,
            "derived key mismatch for alg_id_key={alg_id_key:#x} key_len={key_len}"
        );
    }
}
