#![cfg(windows)]

use formula_io::offcrypto::cryptoapi::{crypt_derive_key, HashAlg};
use windows_sys::Win32::Foundation::GetLastError;
use windows_sys::Win32::Security::Cryptography::{
    CryptAcquireContextW, CryptCreateHash, CryptDeriveKey, CryptDestroyHash, CryptDestroyKey,
    CryptGetHashParam, CryptGetKeyParam, CryptHashData, CryptReleaseContext, CALG_AES_128,
    CALG_AES_192, CALG_AES_256, CALG_SHA1, CRYPT_EXPORTABLE, CRYPT_VERIFYCONTEXT, HP_HASHVAL,
    KP_KEYVAL, PROV_RSA_AES,
};

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
    fn new(provider: &CryptoProvider, alg_id_hash: u32) -> Self {
        let mut hhash = 0usize;
        // SAFETY: FFI call.
        let ok = unsafe { CryptCreateHash(provider.0, alg_id_hash, 0, 0, &mut hhash) };
        assert_ne!(
            ok, 0,
            "CryptCreateHash(alg_id_hash={alg_id_hash:#x}) failed: {}",
            last_err()
        );
        Self(hhash)
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
    fn derive(provider: &CryptoProvider, alg_id_key: u32, hash: &CryptoHash) -> Self {
        let mut hkey = 0usize;
        // SAFETY: FFI call.
        let ok = unsafe { CryptDeriveKey(provider.0, alg_id_key, hash.0, CRYPT_EXPORTABLE, &mut hkey) };
        assert_ne!(
            ok, 0,
            "CryptDeriveKey(alg_id_key={alg_id_key:#x}) failed: {}",
            last_err()
        );
        Self(hkey)
    }

    fn get_key_value(&self) -> Vec<u8> {
        let mut len: u32 = 0;
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetKeyParam(self.0, KP_KEYVAL, std::ptr::null_mut(), &mut len, 0) };
        assert_ne!(
            ok, 0,
            "CryptGetKeyParam(KP_KEYVAL, size query) failed: {}",
            last_err()
        );
        let mut buf = vec![0u8; len as usize];
        // SAFETY: FFI call.
        let ok = unsafe { CryptGetKeyParam(self.0, KP_KEYVAL, buf.as_mut_ptr(), &mut len, 0) };
        assert_ne!(ok, 0, "CryptGetKeyParam(KP_KEYVAL) failed: {}", last_err());
        buf.truncate(len as usize);
        buf
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

#[test]
fn cryptderivekey_matches_cryptoapi_for_sha1_aes_key_sizes() {
    let provider = CryptoProvider::acquire();
    let alg_id_hash = CALG_SHA1;

    // Arbitrary-but-stable input for the hash object.
    let data = b"formula-io offcrypto CryptDeriveKey cross-check";

    for (alg_id_key, key_len) in [
        (CALG_AES_256, 32usize),
        (CALG_AES_192, 24usize),
        (CALG_AES_128, 16usize),
    ] {
        let (hash_value, key_bytes) = derive_key_and_hash(&provider, alg_id_hash, alg_id_key, data);
        assert_eq!(
            key_bytes.len(),
            key_len,
            "CryptoAPI returned unexpected key length for alg_id_key={alg_id_key:#x}"
        );

        let ours = crypt_derive_key(&hash_value, key_len, HashAlg::from_calg_id(alg_id_hash).unwrap());
        assert_eq!(
            key_bytes, ours,
            "derived key mismatch for alg_id_key={alg_id_key:#x} key_len={key_len}"
        );
    }
}
