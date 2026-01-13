use std::io::{Read, Seek, SeekFrom};

use sha1::{Digest as _, Sha1};

const RC4_BLOCK_SIZE: usize = 0x200;

/// Streaming decryptor for Standard/CryptoAPI RC4 `EncryptedPackage` (MS-OFFCRYPTO).
///
/// This reader exposes the decrypted plaintext bytes of the encrypted package without buffering the
/// whole ZIP payload in memory.
///
/// ## Block restart behavior
///
/// The `EncryptedPackage` payload is split into 0x200-byte blocks. Each block is encrypted with a
/// fresh RC4 key derived as:
///
/// `rc4_key_b = Hash(H || LE32(b))[0..key_len]`
///
/// **40-bit note:** CryptoAPI/Office represent a "40-bit" RC4 key as a 128-bit RC4 key where the
/// low 40 bits are set and the remaining 88 bits are zero. Concretely, when `key_len == 5`, the
/// RC4 key bytes are:
///
/// `rc4_key_b = Hash(H || LE32(b))[0..5] || 0x00 * 11` (16 bytes total)
///
/// where:
/// - `H` is the base hash bytes (typically `Hfinal`, 20-byte SHA-1 output)
/// - `b` is the 0-based block index.
///
/// Seeking is supported by re-deriving the block key and discarding `o = pos % 0x200` bytes of
/// RC4 keystream.
#[derive(Debug)]
pub struct Rc4CryptoApiDecryptReader<R: Read + Seek> {
    inner: R,

    /// Absolute stream offset of the first encrypted byte (i.e. after the 8-byte package size
    /// prefix in `EncryptedPackage`).
    ciphertext_start: u64,
    /// Current inner position *relative* to `ciphertext_start`.
    inner_pos: u64,

    /// Total plaintext size (from the 8-byte `EncryptedPackage` length prefix).
    package_size: u64,
    /// Current plaintext offset.
    pos: u64,

    /// Base hash bytes used for per-block key derivation.
    h: Vec<u8>,
    /// RC4 key length in bytes (e.g. `keySize / 8` from EncryptionHeader).
    key_len: usize,

    rc4: Option<Rc4>,
    block_index: Option<u32>,
    /// Offset within the current block that `rc4` is aligned to.
    block_offset: usize,
}

impl<R: Read + Seek> Rc4CryptoApiDecryptReader<R> {
    /// Create a decrypting reader wrapping `inner`.
    ///
    /// `inner` must be positioned at the start of the ciphertext payload (i.e. just after reading
    /// the 8-byte `package_size` prefix from the `EncryptedPackage` stream).
    pub fn new(
        mut inner: R,
        package_size: u64,
        h: Vec<u8>,
        key_len: usize,
    ) -> std::io::Result<Self> {
        if key_len == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "RC4 key_len must be non-zero",
            ));
        }

        // RC4 key bytes are derived by hashing (SHA-1) and truncating, so key_len cannot exceed
        // the SHA-1 output size.
        if key_len > 20 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("RC4 key_len {key_len} exceeds SHA-1 digest size"),
            ));
        }

        let ciphertext_start = inner.stream_position()?;
        Ok(Self {
            inner,
            ciphertext_start,
            inner_pos: 0,
            package_size,
            pos: 0,
            h,
            key_len,
            rc4: None,
            block_index: None,
            block_offset: 0,
        })
    }

    /// Return the plaintext length of the encrypted package.
    pub fn package_size(&self) -> u64 {
        self.package_size
    }

    /// Consume the wrapper and return the underlying reader.
    pub fn into_inner(self) -> R {
        self.inner
    }

    fn ensure_inner_position(&mut self) -> std::io::Result<()> {
        if self.inner_pos == self.pos {
            return Ok(());
        }
        self.inner
            .seek(SeekFrom::Start(self.ciphertext_start + self.pos))?;
        self.inner_pos = self.pos;
        Ok(())
    }

    fn ensure_block(&mut self) -> std::io::Result<()> {
        let block_index_u64 = self.pos / RC4_BLOCK_SIZE as u64;
        let block_index = u32::try_from(block_index_u64).map_err(|_| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "encrypted package position exceeds 32-bit block address space",
            )
        })?;
        let offset = (self.pos % RC4_BLOCK_SIZE as u64) as usize;

        if self.block_index == Some(block_index) && self.block_offset == offset {
            return Ok(());
        }

        // Derive per-block RC4 key: SHA1(H || LE32(block_index)) truncated to key_len.
        let mut hasher = Sha1::new();
        hasher.update(&self.h);
        hasher.update(block_index.to_le_bytes());
        let digest = hasher.finalize();
        let mut rc4 = if self.key_len == 5 {
            // CryptoAPI 40-bit RC4 uses a 128-bit key with the high 88 bits zero.
            let mut padded = [0u8; 16];
            padded[..5].copy_from_slice(&digest[..5]);
            Rc4::new(&padded)
        } else {
            Rc4::new(&digest[..self.key_len])
        };
        rc4.skip(offset);

        self.rc4 = Some(rc4);
        self.block_index = Some(block_index);
        self.block_offset = offset;
        Ok(())
    }

    fn invalidate_cipher_state(&mut self) {
        self.rc4 = None;
        self.block_index = None;
        self.block_offset = 0;
    }
}

impl<R: Read + Seek> Read for Rc4CryptoApiDecryptReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }
        if self.pos >= self.package_size {
            return Ok(0);
        }

        let mut remaining = (self.package_size - self.pos) as usize;
        remaining = remaining.min(buf.len());

        let mut written = 0usize;
        while remaining > 0 {
            self.ensure_block()?;
            self.ensure_inner_position()?;

            let in_block_offset = (self.pos % RC4_BLOCK_SIZE as u64) as usize;
            let block_remaining = RC4_BLOCK_SIZE - in_block_offset;
            let chunk_len = remaining.min(block_remaining);

            let out = &mut buf[written..written + chunk_len];
            let n = self.inner.read(out)?;
            if n == 0 {
                if written == 0 {
                    return Err(std::io::Error::new(
                        std::io::ErrorKind::UnexpectedEof,
                        "unexpected EOF while reading EncryptedPackage ciphertext",
                    ));
                }
                break;
            }
            self.inner_pos = self
                .inner_pos
                .checked_add(n as u64)
                .expect("inner_pos should not overflow u64");

            self.rc4
                .as_mut()
                .expect("rc4 state must be initialized by ensure_block")
                .apply_keystream(&mut out[..n]);

            self.pos += n as u64;
            self.block_offset += n;

            written += n;
            remaining -= n;

            // If the read ended early (common for some Read impls), return what we have.
            if n < chunk_len {
                break;
            }

            // Move to next block when we've fully consumed this one.
            if self.block_offset >= RC4_BLOCK_SIZE {
                self.invalidate_cipher_state();
            }
        }

        Ok(written)
    }
}

impl<R: Read + Seek> Seek for Rc4CryptoApiDecryptReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let base: i128 = match pos {
            SeekFrom::Start(n) => n as i128,
            SeekFrom::Current(off) => self.pos as i128 + off as i128,
            SeekFrom::End(off) => self.package_size as i128 + off as i128,
        };
        if base < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "invalid seek to a negative position",
            ));
        }
        let mut new_pos = base as u64;

        // Clamp to the plaintext EOF boundary. This avoids seeking the underlying stream past the
        // meaningful ciphertext range while still satisfying "seek beyond EOF behaves like EOF".
        if new_pos > self.package_size {
            new_pos = self.package_size;
        }

        self.pos = new_pos;
        self.invalidate_cipher_state();

        self.inner
            .seek(SeekFrom::Start(self.ciphertext_start + self.pos))?;
        self.inner_pos = self.pos;

        Ok(self.pos)
    }
}

#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl std::fmt::Debug for Rc4 {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Avoid dumping the full internal permutation in debug output.
        f.debug_struct("Rc4")
            .field("i", &self.i)
            .field("j", &self.j)
            .finish()
    }
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");

        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }

        let mut j: u8 = 0;
        for i in 0..256u16 {
            let si = s[i as usize];
            j = j.wrapping_add(si).wrapping_add(key[i as usize % key.len()]);
            s.swap(i as usize, j as usize);
        }

        Self { s, i: 0, j: 0 }
    }

    fn next_byte(&mut self) -> u8 {
        self.i = self.i.wrapping_add(1);
        self.j = self.j.wrapping_add(self.s[self.i as usize]);
        self.s.swap(self.i as usize, self.j as usize);
        let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
        self.s[idx as usize]
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            *b ^= self.next_byte();
        }
    }

    fn skip(&mut self, n: usize) {
        for _ in 0..n {
            let _ = self.next_byte();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn encrypt_rc4_cryptoapi(plaintext: &[u8], h: &[u8], key_len: usize) -> Vec<u8> {
        assert!(key_len <= 20);

        let mut out = vec![0u8; plaintext.len()];
        let mut offset = 0usize;
        let mut block_index = 0u32;
        while offset < plaintext.len() {
            let mut hasher = Sha1::new();
            hasher.update(h);
            hasher.update(block_index.to_le_bytes());
            let digest = hasher.finalize();
            let mut rc4 = if key_len == 5 {
                // CryptoAPI 40-bit RC4 uses a 128-bit key with the high 88 bits zero.
                let mut padded = [0u8; 16];
                padded[..5].copy_from_slice(&digest[..5]);
                Rc4::new(&padded)
            } else {
                Rc4::new(&digest[..key_len])
            };

            let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
            out[offset..offset + block_len].copy_from_slice(&plaintext[offset..offset + block_len]);
            rc4.apply_keystream(&mut out[offset..offset + block_len]);

            offset += block_len;
            block_index += 1;
        }
        out
    }

    #[test]
    fn sequential_reads_across_block_boundary() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 16;

        // Ensure plaintext crosses a 0x200 boundary.
        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let mut cursor = Cursor::new(stream);
        cursor.seek(SeekFrom::Start(8)).unwrap();

        let mut reader =
            Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                .unwrap();

        let mut out = vec![0u8; plaintext.len()];
        // Read in small chunks to force multiple calls and cross-block behavior.
        let out_len = out.len();
        let mut read = 0usize;
        while read < out_len {
            let end = read + 33.min(out_len - read);
            let n = reader.read(&mut out[read..end]).unwrap();
            assert!(n > 0, "unexpected EOF while reading");
            read += n;
        }
        assert_eq!(out, plaintext);
    }

    #[test]
    fn sequential_reads_across_block_boundary_with_40_bit_key() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 5; // 40-bit (must be padded to 16 bytes for RC4)

        // Ensure plaintext crosses a 0x200 boundary.
        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE + 64];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len);

        // Simulate EncryptedPackage stream layout: [u64 package_size] + ciphertext.
        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let mut cursor = Cursor::new(stream);
        cursor.seek(SeekFrom::Start(8)).unwrap();

        let mut reader =
            Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                .unwrap();

        let mut out = vec![0u8; plaintext.len()];
        reader.read_exact(&mut out).unwrap();
        assert_eq!(out, plaintext);
    }

    #[test]
    fn seek_into_middle_of_block_and_read() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 16;

        let mut plaintext = vec![0u8; RC4_BLOCK_SIZE * 3];
        for (i, b) in plaintext.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len);

        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let mut cursor = Cursor::new(stream);
        cursor.seek(SeekFrom::Start(8)).unwrap();

        let mut reader =
            Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                .unwrap();

        let seek_pos = (RC4_BLOCK_SIZE as u64) + 0x10;
        reader.seek(SeekFrom::Start(seek_pos)).unwrap();

        let mut buf = [0u8; 64];
        reader.read_exact(&mut buf).unwrap();

        assert_eq!(
            &buf[..],
            &plaintext[seek_pos as usize..seek_pos as usize + buf.len()]
        );
    }

    #[test]
    fn seek_beyond_package_size_behaves_like_eof() {
        let h = b"0123456789ABCDEFGHIJ".to_vec(); // 20 bytes
        let key_len = 16;

        let plaintext = b"hello world".to_vec();
        let ciphertext = encrypt_rc4_cryptoapi(&plaintext, &h, key_len);

        let mut stream = Vec::new();
        stream.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
        stream.extend_from_slice(&ciphertext);

        let mut cursor = Cursor::new(stream);
        cursor.seek(SeekFrom::Start(8)).unwrap();

        let mut reader =
            Rc4CryptoApiDecryptReader::new(cursor, plaintext.len() as u64, h.clone(), key_len)
                .unwrap();

        // Seek beyond EOF; reader should clamp to EOF and reads should return 0.
        reader
            .seek(SeekFrom::Start(plaintext.len() as u64 + 100))
            .unwrap();

        let mut buf = [0u8; 32];
        let n = reader.read(&mut buf).unwrap();
        assert_eq!(n, 0);
    }
}
