//! Regenerates the encrypted `.xls` fixtures under `tests/fixtures/encrypted/`.
//!
//! This is an ignored test so it doesn't run in CI; it's a convenient, in-repo way to keep the
//! binary fixture blobs reproducible and auditable.
//!
//! This generator produces **decryptable** BIFF8 `.xls` workbooks for each supported `FILEPASS`
//! encryption scheme:
//! - XOR obfuscation (`wEncryptionType=0x0000`)
//! - RC4 "Standard Encryption" (`wEncryptionType=0x0001`, `subType=0x0001`)
//! - RC4 CryptoAPI (`wEncryptionType=0x0001`, `subType=0x0002`)
//!
//! Run:
//!   cargo test -p formula-xls --test regenerate_encrypted_xls_fixtures -- --ignored

use md5::{Digest as _, Md5};
use sha1::Sha1;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use std::path::PathBuf;

const RECORD_BOF: u16 = 0x0809;
const RECORD_EOF: u16 = 0x000A;
const RECORD_FILEPASS: u16 = 0x002F;

const RECORD_CODEPAGE: u16 = 0x0042;
const RECORD_WINDOW1: u16 = 0x003D;
const RECORD_FONT: u16 = 0x0031;
const RECORD_XF: u16 = 0x00E0;
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_INTERFACEHDR: u16 = 0x00E1;
const RECORD_RRDINFO: u16 = 0x0138;
const RECORD_RRDHEAD: u16 = 0x0139;
const RECORD_USREXCL: u16 = 0x0194;
const RECORD_FILELOCK: u16 = 0x0195;

const RECORD_DIMENSIONS: u16 = 0x0200;
const RECORD_WINDOW2: u16 = 0x023E;
const RECORD_NUMBER: u16 = 0x0203;

const BOF_VERSION_BIFF8: u16 = 0x0600;
const BOF_DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const BOF_DT_WORKSHEET: u16 = 0x0010;

const XF_FLAG_LOCKED: u16 = 0x0001;
const XF_FLAG_STYLE: u16 = 0x0004;

fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn push_record(out: &mut Vec<u8>, record_id: u16, payload: &[u8]) {
    out.extend_from_slice(&record(record_id, payload));
}

fn bof_biff8(dt: u16) -> [u8; 16] {
    // BOF record payload (BIFF8) matching `tests/common/xls_fixture_builder.rs`.
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&BOF_VERSION_BIFF8.to_le_bytes());
    out[2..4].copy_from_slice(&dt.to_le_bytes());
    out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
    out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
    out
}

fn build_xls_bytes(workbook_stream: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(workbook_stream)
            .expect("write Workbook stream bytes");
    }
    ole.into_inner().into_inner()
}

fn read_workbook_stream_from_xls(path: &Path) -> Vec<u8> {
    let mut comp = cfb::open(path).expect("open xls cfb");
    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(mut stream) = comp.open_stream(candidate) {
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf).expect("read workbook stream");
            // Some writers may include trailing padding bytes after the final EOF record. Legacy
            // RC4 standard decryption expects the workbook stream to end on a record boundary, so
            // trim any non-record tail bytes to keep regenerated fixtures decryptable.
            let end = last_complete_record_end(&buf);
            if end > 0 && end < buf.len() {
                buf.truncate(end);
            }
            return buf;
        }
    }
    panic!("xls fixture missing Workbook/Book stream: {path:?}");
}

fn last_complete_record_end(stream: &[u8]) -> usize {
    let mut offset = 0usize;
    let mut last_end = 0usize;
    while let Some((_, _, next)) = super_read_record(stream, offset) {
        last_end = next;
        offset = next;
    }
    last_end
}

fn window1() -> [u8; 18] {
    // WINDOW1 record payload (BIFF8, 18 bytes).
    let mut out = [0u8; 18];
    // cTabSel = 1
    out[14..16].copy_from_slice(&1u16.to_le_bytes());
    // wTabRatio = 600 (arbitrary non-zero)
    out[16..18].copy_from_slice(&600u16.to_le_bytes());
    out
}

fn window2() -> [u8; 18] {
    // WINDOW2 record payload (BIFF8).
    let mut out = [0u8; 18];
    let grbit: u16 = 0x02B6;
    out[0..2].copy_from_slice(&grbit.to_le_bytes());
    out
}

fn write_short_unicode_string(out: &mut Vec<u8>, s: &str) {
    // BIFF8 short XLUnicodeString: cch (u8) + flags (u8) + chars.
    // Use compressed (flags=0) for ASCII strings.
    let bytes = s.as_bytes();
    assert!(
        bytes.len() <= u8::MAX as usize,
        "string too long for BIFF8 short string"
    );
    out.push(bytes.len() as u8);
    out.push(0); // flags: compressed ANSI
    out.extend_from_slice(bytes);
}

fn font(name: &str) -> Vec<u8> {
    // Minimal BIFF8 FONT record payload (enough for calamine).
    const COLOR_AUTOMATIC: u16 = 0x7FFF;

    let mut out = Vec::<u8>::new();
    out.extend_from_slice(&200u16.to_le_bytes()); // height (10pt)
    out.extend_from_slice(&0u16.to_le_bytes()); // option flags
    out.extend_from_slice(&COLOR_AUTOMATIC.to_le_bytes()); // color
    out.extend_from_slice(&400u16.to_le_bytes()); // weight
    out.extend_from_slice(&0u16.to_le_bytes()); // escapement
    out.push(0); // underline
    out.push(0); // family
    out.push(0); // charset
    out.push(0); // reserved
    write_short_unicode_string(&mut out, name);
    out
}

fn xf_record_with_alignment(
    font_idx: u16,
    fmt_idx: u16,
    is_style_xf: bool,
    alignment: u8,
) -> [u8; 20] {
    let mut out = [0u8; 20];
    out[0..2].copy_from_slice(&font_idx.to_le_bytes());
    out[2..4].copy_from_slice(&fmt_idx.to_le_bytes());

    // Protection / type / parent:
    // bit0: locked (1)
    // bit2: xfType (1 = style XF, 0 = cell XF)
    // bits4-15: parent style XF index (0)
    let flags: u16 = XF_FLAG_LOCKED | if is_style_xf { XF_FLAG_STYLE } else { 0 };
    out[4..6].copy_from_slice(&flags.to_le_bytes());

    // BIFF8 alignment byte (horizontal + vertical + wrap).
    //
    // - bits 0-2: horizontal alignment (0 = General)
    // - bit 3: wrap
    // - bits 4-6: vertical alignment (0 = Top, 2 = Bottom (default))
    //
    // The CryptoAPI fixture uses a non-default vertical alignment ("Top", 0x00) for the *cell*
    // XF so decrypted workbooks exercise post-decryption style parsing. Style XFs keep Excel's
    // defaults ("Bottom", 0x20).
    out[6] = alignment;

    // Attribute flags: apply all so fixture cell XFs don't rely on inheritance.
    out[9] = 0x3F;
    out
}

fn xf_record(font_idx: u16, fmt_idx: u16, is_style_xf: bool) -> [u8; 20] {
    // Default BIFF8 alignment: General + Bottom.
    xf_record_with_alignment(font_idx, fmt_idx, is_style_xf, 0x20)
}

fn number_cell(row: u16, col: u16, xf: u16, v: f64) -> [u8; 14] {
    let mut out = [0u8; 14];
    out[0..2].copy_from_slice(&row.to_le_bytes());
    out[2..4].copy_from_slice(&col.to_le_bytes());
    out[4..6].copy_from_slice(&xf.to_le_bytes());
    out[6..14].copy_from_slice(&v.to_le_bytes());
    out
}

fn build_plain_biff8_workbook_stream(filepass_payload: &[u8], extra_sheet_payload_len: usize) -> Vec<u8> {
    // Common workbook structure used across all encrypted fixtures. The record payload bytes after
    // FILEPASS are encrypted by the scheme-specific builder.

    // -- Globals ----------------------------------------------------------------
    let mut globals = Vec::<u8>::new();
    push_record(&mut globals, RECORD_BOF, &bof_biff8(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_FILEPASS, filepass_payload);
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: keep the usual 16 style XFs so BIFF consumers stay happy.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // Cell XF used by the sheet's NUMBER record. Make it non-default (vertical alignment = Top)
    // so the decrypt + BIFF parser tests can assert that XF metadata after FILEPASS was imported.
    let xf_cell: u16 = 16;
    push_record(&mut globals, RECORD_XF, &xf_record_with_alignment(0, 0, false, 0x00));

    // BoundSheet with placeholder offset.
    let boundsheet_start = globals.len();
    let mut boundsheet = Vec::<u8>::new();
    boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
    boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
    write_short_unicode_string(&mut boundsheet, "Sheet1");
    push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
    let boundsheet_offset_pos = boundsheet_start + 4;

    push_record(&mut globals, RECORD_EOF, &[]);

    // -- Sheet ------------------------------------------------------------------
    let sheet_offset = globals.len();
    let mut sheet = Vec::<u8>::new();
    push_record(&mut sheet, RECORD_BOF, &bof_biff8(BOF_DT_WORKSHEET));

    // DIMENSIONS: rows [0, 1) cols [0, 1) => A1.
    let mut dims = Vec::<u8>::new();
    dims.extend_from_slice(&0u32.to_le_bytes()); // first row
    dims.extend_from_slice(&1u32.to_le_bytes()); // last row + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // first col
    dims.extend_from_slice(&1u16.to_le_bytes()); // last col + 1
    dims.extend_from_slice(&0u16.to_le_bytes()); // reserved
    push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
    push_record(&mut sheet, RECORD_WINDOW2, &window2());
    push_record(&mut sheet, RECORD_NUMBER, &number_cell(0, 0, xf_cell, 42.0));

    if extra_sheet_payload_len > 0 {
        // Add an extra unknown record with a large payload so RC4 Standard fixtures cross the
        // 1024-byte payload boundary and exercise per-block rekeying.
        const RECORD_DUMMY_UNKNOWN: u16 = 0xFFFF;
        let mut payload = Vec::with_capacity(extra_sheet_payload_len);
        for i in 0..extra_sheet_payload_len {
            payload.push(((i * 31) % 251) as u8);
        }
        push_record(&mut sheet, RECORD_DUMMY_UNKNOWN, &payload);
    }

    push_record(&mut sheet, RECORD_EOF, &[]);

    // Patch BoundSheet offset to point at the sheet BOF.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn sha1_bytes(chunks: &[&[u8]]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    out
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len().saturating_mul(2));
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn utf16le_bytes_truncated_15(password: &str) -> Vec<u8> {
    // Excel 97-2003 legacy RC4 encryption truncates passwords to 15 UTF-16 code units.
    let mut out = Vec::with_capacity(password.len().min(15) * 2);
    for unit in password.encode_utf16().take(15) {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn md5_bytes(chunks: &[&[u8]]) -> [u8; 16] {
    let mut h = Md5::new();
    for chunk in chunks {
        h.update(chunk);
    }
    h.finalize().into()
}

fn derive_rc4_standard_intermediate_key(password: &str, salt: &[u8; 16]) -> [u8; 16] {
    // [MS-OFFCRYPTO] "Standard Encryption" key derivation (Excel 97-2003 RC4):
    // - password_hash = MD5(UTF16LE(password)[..15])
    // - intermediate_key = MD5(password_hash + salt)
    let pw_bytes = utf16le_bytes_truncated_15(password);
    let password_hash = md5_bytes(&[&pw_bytes]);
    md5_bytes(&[&password_hash, salt])
}

fn derive_rc4_standard_block_key(intermediate_key: &[u8; 16], block: u32) -> [u8; 16] {
    // block_key = MD5(intermediate_key + block_index_le32)
    md5_bytes(&[intermediate_key, &block.to_le_bytes()])
}

fn build_filepass_rc4_standard_payload(password: &str) -> (Vec<u8>, [u8; 16]) {
    // FILEPASS payload layout (BIFF8 RC4 "Standard Encryption") [MS-XLS 2.4.105]:
    // - u16 wEncryptionType = 0x0001 (RC4)
    // - u16 major = 0x0001 (RC4 standard subtype)
    // - u16 minor = 0x0002 (128-bit)
    // - 16-byte salt
    // - 16-byte encrypted verifier
    // - 16-byte encrypted verifier hash (MD5)
    //
    // Note: Passwords are truncated to the first 15 UTF-16 code units per [MS-OFFCRYPTO].

    let salt: [u8; 16] = [
        0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
        0xDC, 0xFE,
    ];
    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];
    let verifier_hash_plain = md5_bytes(&[&verifier_plain]);

    let intermediate_key = derive_rc4_standard_intermediate_key(password, &salt);
    let block0_key = derive_rc4_standard_block_key(&intermediate_key, 0);
    let mut rc4 = Rc4::new(&block0_key);
    let mut verifier_buf = [0u8; 32];
    verifier_buf[..16].copy_from_slice(&verifier_plain);
    verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
    rc4.apply_keystream(&mut verifier_buf);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
    let mut encrypted_verifier_hash = [0u8; 16];
    encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

    let mut payload = Vec::<u8>::new();
    payload.extend_from_slice(&[0x01, 0x00]); // wEncryptionType (RC4)
    payload.extend_from_slice(&[0x01, 0x00]); // major (subType)
    payload.extend_from_slice(&[0x02, 0x00]); // minor (128-bit)
    payload.extend_from_slice(&salt);
    payload.extend_from_slice(&encrypted_verifier);
    payload.extend_from_slice(&encrypted_verifier_hash);

    (payload, intermediate_key)
}

fn derive_cryptoapi_key_material(password: &str, salt: &[u8; 16]) -> [u8; 20] {
    // Matches `crates/formula-xls/src/biff/encryption/cryptoapi.rs` (CryptoAPI RC4 password KDF, SHA-1).
    const PASSWORD_HASH_ITERATIONS: u32 = 50_000;

    let pw_bytes = utf16le_bytes(password);
    let mut hash = sha1_bytes(&[salt, &pw_bytes]);

    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        hash = sha1_bytes(&[&iter, &hash]);
    }

    hash
}

fn derive_cryptoapi_key_material_md5(password: &str, salt: &[u8; 16]) -> [u8; 16] {
    // Matches `crates/formula-xls/src/biff/encryption/cryptoapi.rs` (CryptoAPI RC4 password KDF, MD5).
    const PASSWORD_HASH_ITERATIONS: u32 = 50_000;

    let pw_bytes = utf16le_bytes(password);
    let mut hash = md5_bytes(&[salt, &pw_bytes]);

    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        hash = md5_bytes(&[&iter, &hash]);
    }

    hash
}

fn derive_cryptoapi_key_material_legacy(password: &str, salt: &[u8; 16]) -> [u8; 20] {
    // Legacy BIFF8 CryptoAPI FILEPASS layout (wEncryptionInfo=0x0004):
    //   key_material = SHA1(salt + UTF16LE(password))
    //
    // Unlike the modern CryptoAPI encoding, this does *not* apply the 50,000-iteration hashing step.
    let pw_bytes = utf16le_bytes(password);
    sha1_bytes(&[salt, &pw_bytes])
}

fn derive_cryptoapi_block_key(key_material: &[u8; 20], block: u32, key_len: usize) -> Vec<u8> {
    let block_bytes = block.to_le_bytes();
    let digest = sha1_bytes(&[key_material, &block_bytes]);
    digest[..key_len].to_vec()
}

fn derive_cryptoapi_block_key_md5(key_material: &[u8; 16], block: u32, key_len: usize) -> Vec<u8> {
    let block_bytes = block.to_le_bytes();
    let digest = md5_bytes(&[key_material, &block_bytes]);
    if key_len == 5 {
        // CryptoAPI "40-bit" RC4 uses a 16-byte RC4 key with the high 88 bits set to zero.
        let mut key = digest[..5].to_vec();
        key.resize(16, 0);
        return key;
    }
    digest[..key_len].to_vec()
}

#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
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
        for b in data.iter_mut() {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

fn rc4_discard(rc4: &mut Rc4, mut n: usize) {
    // Advance the internal RC4 state without caring about the output bytes. This is used by the
    // absolute-offset legacy BIFF8 CryptoAPI RC4 variant to jump to `pos_in_block` within a
    // 1024-byte rekey segment.
    let mut scratch = [0u8; 64];
    while n > 0 {
        let take = n.min(scratch.len());
        rc4.apply_keystream(&mut scratch[..take]);
        n -= take;
    }
}

fn apply_cryptoapi_legacy_keystream_by_offset(
    bytes: &mut [u8],
    start_offset: usize,
    key_material: &[u8; 20],
    key_len: usize,
) {
    // Symmetric encrypt/decrypt helper matching `decrypt_range_by_offset` in
    // `crates/formula-xls/src/biff/encryption/cryptoapi.rs`.
    const BLOCK_SIZE: usize = 1024;

    let mut stream_pos = start_offset;
    let mut remaining = bytes.len();
    let mut pos = 0usize;
    while remaining > 0 {
        let block = (stream_pos / BLOCK_SIZE) as u32;
        let in_block = stream_pos % BLOCK_SIZE;
        let take = remaining.min(BLOCK_SIZE - in_block);

        let key = derive_cryptoapi_block_key(key_material, block, key_len);
        let mut rc4 = Rc4::new(&key);
        rc4_discard(&mut rc4, in_block);
        rc4.apply_keystream(&mut bytes[pos..pos + take]);

        stream_pos += take;
        pos += take;
        remaining -= take;
    }
}

struct PayloadRc4 {
    key_material: [u8; 20],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

struct PayloadRc4Md5 {
    key_material: [u8; 16],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4Md5 {
    fn new(key_material: [u8; 16], key_len: usize) -> Self {
        let key0 = derive_cryptoapi_block_key_md5(&key_material, 0, key_len);
        Self {
            key_material,
            key_len,
            block: 0,
            pos_in_block: 0,
            rc4: Rc4::new(&key0),
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_cryptoapi_block_key_md5(&self.key_material, self.block, self.key_len);
        self.rc4 = Rc4::new(&key);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        const PAYLOAD_BLOCK_SIZE: usize = 1024;
        while !data.is_empty() {
            if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

impl PayloadRc4 {
    fn new(key_material: [u8; 20], key_len: usize) -> Self {
        let key0 = derive_cryptoapi_block_key(&key_material, 0, key_len);
        Self {
            key_material,
            key_len,
            block: 0,
            pos_in_block: 0,
            rc4: Rc4::new(&key0),
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_cryptoapi_block_key(&self.key_material, self.block, self.key_len);
        self.rc4 = Rc4::new(&key);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        const PAYLOAD_BLOCK_SIZE: usize = 1024;
        while !data.is_empty() {
            if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

struct PayloadRc4Standard {
    intermediate_key: [u8; 16],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4Standard {
    fn new(intermediate_key: [u8; 16], key_len: usize) -> Self {
        let key0 = derive_rc4_standard_block_key(&intermediate_key, 0);
        Self {
            intermediate_key,
            key_len,
            block: 0,
            pos_in_block: 0,
            rc4: Rc4::new(&key0[..key_len]),
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_rc4_standard_block_key(&self.intermediate_key, self.block);
        self.rc4 = Rc4::new(&key[..self.key_len]);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        const PAYLOAD_BLOCK_SIZE: usize = 1024;
        while !data.is_empty() {
            if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

// -------------------------------------------------------------------------------------------------
// BIFF8 XOR obfuscation (MS-OFFCRYPTO/MS-XLS "Method 1")
// -------------------------------------------------------------------------------------------------

// [MS-OFFCRYPTO] 2.3.7.2 (CreateXorArray_Method1) constants.
const XOR_PAD_ARRAY: [u8; 15] = [
    0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
];

const XOR_INITIAL_CODE: [u16; 15] = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F,
    0x84F9, 0x280C, 0xA96A, 0x4EC3,
];

const XOR_MATRIX: [u16; 105] = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B,
    0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375,
    0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D,
    0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667,
    0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68,
    0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1,
    0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
    0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1,
    0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021,
    0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4,
];

fn xor_ror(byte1: u8, byte2: u8) -> u8 {
    (byte1 ^ byte2).rotate_right(1)
}

fn create_password_verifier_method1(password: &[u8]) -> u16 {
    let mut verifier: u16 = 0;
    let mut password_array = Vec::<u8>::with_capacity(password.len().saturating_add(1));
    password_array.push(password.len() as u8);
    password_array.extend_from_slice(password);

    for &b in password_array.iter().rev() {
        let intermediate1 = if (verifier & 0x4000) == 0 { 0u16 } else { 1u16 };
        let intermediate2 = verifier.wrapping_mul(2) & 0x7FFF;
        let intermediate3 = intermediate1 | intermediate2;
        verifier = intermediate3 ^ (b as u16);
    }

    verifier ^ 0xCE4B
}

fn create_xor_key_method1(password: &[u8]) -> u16 {
    if password.is_empty() || password.len() > 15 {
        return 0;
    }

    let mut xor_key = XOR_INITIAL_CODE[password.len() - 1];
    let mut current_element: i32 = 0x68;

    for &byte in password.iter().rev() {
        let mut ch = byte;
        for _ in 0..7 {
            if (ch & 0x40) != 0 {
                if current_element < 0 || current_element as usize >= XOR_MATRIX.len() {
                    return xor_key;
                }
                xor_key ^= XOR_MATRIX[current_element as usize];
            }
            ch = ch.wrapping_mul(2);
            current_element -= 1;
        }
    }

    xor_key
}

fn create_xor_array_method1(password: &[u8], xor_key: u16) -> [u8; 16] {
    let mut out = [0u8; 16];
    let mut index = password.len();

    let key_high = (xor_key >> 8) as u8;
    let key_low = (xor_key & 0x00FF) as u8;

    if index % 2 == 1 {
        if index < out.len() {
            out[index] = xor_ror(XOR_PAD_ARRAY[0], key_high);
        }

        index = index.saturating_sub(1);

        if !password.is_empty() && index < out.len() {
            let password_last = password[password.len() - 1];
            out[index] = xor_ror(password_last, key_low);
        }
    }

    while index > 0 {
        index = index.saturating_sub(1);
        if index < password.len() {
            out[index] = xor_ror(password[index], key_high);
        }

        index = index.saturating_sub(1);
        if index < password.len() {
            out[index] = xor_ror(password[index], key_low);
        }
    }

    let mut out_index: i32 = 15;
    let mut pad_index: i32 = 15i32 - (password.len() as i32);
    while pad_index > 0 {
        if out_index < 0 {
            break;
        }

        let pi = pad_index as usize;
        if pi < XOR_PAD_ARRAY.len() {
            out[out_index as usize] = xor_ror(XOR_PAD_ARRAY[pi], key_high);
        }
        out_index -= 1;
        pad_index -= 1;

        if out_index < 0 {
            break;
        }

        let pi = pad_index.max(0) as usize;
        if pi < XOR_PAD_ARRAY.len() {
            out[out_index as usize] = xor_ror(XOR_PAD_ARRAY[pi], key_low);
        }
        out_index -= 1;
        pad_index -= 1;
    }

    out
}

fn encrypt_payloads_after_filepass_xor_method1(
    workbook_stream: &mut [u8],
    filepass_data_end: usize,
    xor_array: &[u8; 16],
) {
    let mut cursor = filepass_data_end;
    while cursor < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(cursor);
        if remaining < 4 {
            break;
        }

        let record_id = u16::from_le_bytes([workbook_stream[cursor], workbook_stream[cursor + 1]]);
        let len = u16::from_le_bytes([workbook_stream[cursor + 2], workbook_stream[cursor + 3]]) as usize;
        let data_start = cursor + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= workbook_stream.len(),
            "generated BIFF record extends past end of stream"
        );

        // Per [MS-XLS] 2.2.10, some record payloads are not encrypted or partially encrypted even
        // after FILEPASS.
        let mut encrypt_from = 0usize;
        let skip_entire_payload = matches!(
            record_id,
            RECORD_BOF
                | RECORD_FILEPASS
                | RECORD_INTERFACEHDR
                | RECORD_FILELOCK
                | RECORD_USREXCL
                | RECORD_RRDINFO
                | RECORD_RRDHEAD
        );

        if !skip_entire_payload {
            if record_id == RECORD_BOUNDSHEET {
                // BoundSheet.lbPlyPos MUST NOT be encrypted.
                encrypt_from = 4.min(len);
            }

            let payload = &mut workbook_stream[data_start..data_end];
            for i in encrypt_from..payload.len() {
                let abs_pos = data_start + i;
                let mut value = payload[i];
                value = value.rotate_left(5);
                value ^= xor_array[abs_pos % 16];
                payload[i] = value;
            }
        }

        cursor = data_end;
    }
}

fn build_filepass_cryptoapi_payload(
    password: &str,
    salt: &[u8; 16],
    version_major: u16,
    version_minor: u16,
    provider_type: u32,
    csp_name: Option<&str>,
) -> Vec<u8> {
    // FILEPASS payload layout (CryptoAPI) [MS-XLS 2.4.105]:
    //   u16 wEncryptionType = 0x0001 (RC4)
    //   u16 wEncryptionSubType = 0x0002 (CryptoAPI)
    //   u32 dwEncryptionInfoLen
    //   EncryptionInfo bytes
    //
    // EncryptionInfo matches [MS-OFFCRYPTO] RC4 CryptoAPI.
    const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
    const ENCRYPTION_SUBTYPE_CRYPTOAPI: u16 = 0x0002;
    const CALG_RC4: u32 = 0x0000_6801;
    const CALG_SHA1: u32 = 0x0000_8004;

    // Deterministic verifier bytes.
    let verifier_plain: [u8; 16] = [
        0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D, 0x1E,
        0x0F,
    ];
    let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

    // Encrypt verifier + hash using block 0 key.
    let key_material = derive_cryptoapi_key_material(password, salt);
    let key0 = derive_cryptoapi_block_key(&key_material, 0, 16);
    let mut rc4 = Rc4::new(&key0);

    let mut verifier_buf = [0u8; 36];
    verifier_buf[..16].copy_from_slice(&verifier_plain);
    verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
    rc4.apply_keystream(&mut verifier_buf);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
    let mut encrypted_verifier_hash = [0u8; 20];
    encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

    // EncryptionHeader (fixed 32 bytes + optional CSPName UTF-16LE NUL-terminated) [MS-OFFCRYPTO].
    let mut header = Vec::<u8>::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&128u32.to_le_bytes()); // KeySize bits
    header.extend_from_slice(&provider_type.to_le_bytes()); // ProviderType
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
    if let Some(csp_name) = csp_name {
        // CSPName is stored as a null-terminated UTF-16LE string.
        for cu in csp_name.encode_utf16() {
            header.extend_from_slice(&cu.to_le_bytes());
        }
        header.extend_from_slice(&0u16.to_le_bytes());
    }

    // EncryptionVerifier.
    let mut verifier = Vec::<u8>::new();
    verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier.extend_from_slice(salt);
    verifier.extend_from_slice(&encrypted_verifier);
    verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
    verifier.extend_from_slice(&encrypted_verifier_hash);

    // EncryptionInfo:
    //   u16 MajorVersion, u16 MinorVersion, u32 Flags, u32 HeaderSize, EncryptionHeader, EncryptionVerifier.
    let mut enc_info = Vec::<u8>::new();
    enc_info.extend_from_slice(&version_major.to_le_bytes()); // Major
    enc_info.extend_from_slice(&version_minor.to_le_bytes()); // Minor
    enc_info.extend_from_slice(&0u32.to_le_bytes()); // Flags
    enc_info.extend_from_slice(&(header.len() as u32).to_le_bytes()); // HeaderSize
    enc_info.extend_from_slice(&header);
    enc_info.extend_from_slice(&verifier);

    let mut payload = Vec::<u8>::new();
    payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
    payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
    payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
    payload.extend_from_slice(&enc_info);
    payload
}

fn build_filepass_cryptoapi_md5_payload(
    password: &str,
    salt: &[u8; 16],
    version_major: u16,
    version_minor: u16,
    provider_type: u32,
    csp_name: Option<&str>,
) -> Vec<u8> {
    // CryptoAPI FILEPASS payload that uses MD5 for password hashing + verifier hashing.
    const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
    const ENCRYPTION_SUBTYPE_CRYPTOAPI: u16 = 0x0002;
    const CALG_RC4: u32 = 0x0000_6801;
    const CALG_MD5: u32 = 0x0000_8003;

    // Deterministic verifier bytes.
    let verifier_plain: [u8; 16] = [
        0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D,
        0x1E, 0x0F,
    ];
    let verifier_hash_plain = md5_bytes(&[&verifier_plain]);

    // Encrypt verifier + hash using block 0 key.
    let key_material = derive_cryptoapi_key_material_md5(password, salt);
    let key0 = derive_cryptoapi_block_key_md5(&key_material, 0, 16);
    let mut rc4 = Rc4::new(&key0);

    let mut verifier_buf = [0u8; 32];
    verifier_buf[..16].copy_from_slice(&verifier_plain);
    verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
    rc4.apply_keystream(&mut verifier_buf);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
    let mut encrypted_verifier_hash = [0u8; 16];
    encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

    // EncryptionHeader (fixed 32 bytes + optional CSPName UTF-16LE NUL-terminated) [MS-OFFCRYPTO].
    let mut header = Vec::<u8>::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_MD5.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&128u32.to_le_bytes()); // KeySize bits
    header.extend_from_slice(&provider_type.to_le_bytes()); // ProviderType
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
    if let Some(csp_name) = csp_name {
        for cu in csp_name.encode_utf16() {
            header.extend_from_slice(&cu.to_le_bytes());
        }
        header.extend_from_slice(&0u16.to_le_bytes());
    }

    // EncryptionVerifier.
    let mut verifier = Vec::<u8>::new();
    verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier.extend_from_slice(salt);
    verifier.extend_from_slice(&encrypted_verifier);
    verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
    verifier.extend_from_slice(&encrypted_verifier_hash);

    // EncryptionInfo:
    //   u16 MajorVersion, u16 MinorVersion, u32 Flags, u32 HeaderSize, EncryptionHeader, EncryptionVerifier.
    let mut enc_info = Vec::<u8>::new();
    enc_info.extend_from_slice(&version_major.to_le_bytes()); // Major
    enc_info.extend_from_slice(&version_minor.to_le_bytes()); // Minor
    enc_info.extend_from_slice(&0u32.to_le_bytes()); // Flags
    enc_info.extend_from_slice(&(header.len() as u32).to_le_bytes()); // HeaderSize
    enc_info.extend_from_slice(&header);
    enc_info.extend_from_slice(&verifier);

    let mut payload = Vec::<u8>::new();
    payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
    payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
    payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
    payload.extend_from_slice(&enc_info);
    payload
}

fn build_filepass_cryptoapi_legacy_payload(
    password: &str,
    salt: &[u8; 16],
    version_major: u16,
    version_minor: u16,
    provider_type: u32,
    csp_name: Option<&str>,
) -> (Vec<u8>, [u8; 20]) {
    // Legacy BIFF8 RC4 CryptoAPI FILEPASS layout ("wEncryptionInfo=0x0004") supported by
    // `crates/formula-xls/src/biff/encryption/cryptoapi.rs`.
    //
    // FILEPASS payload layout:
    // - u16 wEncryptionType = 0x0001 (RC4)
    // - u16 wEncryptionInfo = 0x0004 (legacy CryptoAPI)
    // - u16 vMajor
    // - u16 vMinor
    // - u16 reserved (0)
    // - u32 headerSize
    // - EncryptionHeader bytes
    // - EncryptionVerifier bytes
    const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
    const ENCRYPTION_INFO_CRYPTOAPI_LEGACY: u16 = 0x0004;
    const CALG_RC4: u32 = 0x0000_6801;
    const CALG_SHA1: u32 = 0x0000_8004;

    // Deterministic verifier bytes.
    let verifier_plain: [u8; 16] = [
        0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D,
        0x1E, 0x0F,
    ];
    let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

    // Encrypt verifier + hash using the legacy key material (no 50k-iteration hardening).
    let key_material = derive_cryptoapi_key_material_legacy(password, salt);
    let key0 = derive_cryptoapi_block_key(&key_material, 0, 16);
    let mut rc4 = Rc4::new(&key0);

    let mut verifier_buf = [0u8; 36];
    verifier_buf[..16].copy_from_slice(&verifier_plain);
    verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
    rc4.apply_keystream(&mut verifier_buf);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
    let mut encrypted_verifier_hash = [0u8; 20];
    encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

    // EncryptionHeader (fixed 32 bytes + optional CSPName UTF-16LE NUL-terminated) [MS-OFFCRYPTO].
    let mut header = Vec::<u8>::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&128u32.to_le_bytes()); // KeySize bits
    header.extend_from_slice(&provider_type.to_le_bytes()); // ProviderType
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
    if let Some(csp_name) = csp_name {
        for cu in csp_name.encode_utf16() {
            header.extend_from_slice(&cu.to_le_bytes());
        }
        header.extend_from_slice(&0u16.to_le_bytes());
    }

    // EncryptionVerifier (salt is stored plaintext).
    let mut verifier = Vec::<u8>::new();
    verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier.extend_from_slice(salt);
    verifier.extend_from_slice(&encrypted_verifier);
    verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
    verifier.extend_from_slice(&encrypted_verifier_hash);

    let mut payload = Vec::<u8>::new();
    payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
    payload.extend_from_slice(&ENCRYPTION_INFO_CRYPTOAPI_LEGACY.to_le_bytes());
    payload.extend_from_slice(&version_major.to_le_bytes());
    payload.extend_from_slice(&version_minor.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
    payload.extend_from_slice(&(header.len() as u32).to_le_bytes());
    payload.extend_from_slice(&header);
    payload.extend_from_slice(&verifier);

    (payload, key_material)
}

fn encrypt_payloads_after_filepass<T>(
    workbook_stream: &mut [u8],
    filepass_data_end: usize,
    mut cipher: T,
) where
    T: FnMut(&mut [u8]),
{
    let mut cursor = filepass_data_end;
    while cursor < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(cursor);
        if remaining < 4 {
            break;
        }

        let len = u16::from_le_bytes([workbook_stream[cursor + 2], workbook_stream[cursor + 3]]) as usize;
        let data_start = cursor + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= workbook_stream.len(),
            "generated BIFF record extends past end of stream"
        );

        cipher(&mut workbook_stream[data_start..data_end]);
        cursor = data_end;
    }
}

fn encrypt_payloads_after_filepass_cryptoapi_legacy(
    workbook_stream: &mut [u8],
    filepass_data_end: usize,
    key_material: &[u8; 20],
    key_len: usize,
) {
    // Legacy BIFF8 CryptoAPI RC4 encrypts record payloads using an absolute-offset stream position
    // that includes record headers. Mirror the decryptor's behavior in
    // `crates/formula-xls/src/biff/encryption/cryptoapi.rs` so regenerated fixtures remain decryptable.
    let mut offset = filepass_data_end;
    let mut stream_pos = filepass_data_end;

    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(offset);
        if remaining < 4 {
            break;
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len = u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= workbook_stream.len(),
            "generated BIFF record extends past end of stream"
        );

        // Record headers are not encrypted but still advance the CryptoAPI stream position.
        stream_pos = stream_pos.saturating_add(4);

        // Mirror `is_never_encrypted_record` from the decryptor.
        let skip_entire_payload = matches!(record_id, RECORD_BOF | RECORD_FILEPASS | RECORD_INTERFACEHDR);

        if !skip_entire_payload {
            match record_id {
                RECORD_BOUNDSHEET => {
                    // BoundSheet.lbPlyPos MUST NOT be encrypted.
                    if len > 4 {
                        let decrypt_start = stream_pos.saturating_add(4);
                        apply_cryptoapi_legacy_keystream_by_offset(
                            &mut workbook_stream[data_start + 4..data_end],
                            decrypt_start,
                            key_material,
                            key_len,
                        );
                    }
                }
                _ => apply_cryptoapi_legacy_keystream_by_offset(
                    &mut workbook_stream[data_start..data_end],
                    stream_pos,
                    key_material,
                    key_len,
                ),
            }
        }

        // Advance past the record payload, regardless of whether we encrypted it.
        stream_pos = stream_pos.saturating_add(len);
        offset = data_end;
    }
}

fn patch_boundsheet_offsets_for_inserted_record(workbook_stream: &mut [u8], inserted_len: u32) {
    // BOUNDSHEET payload layout (BIFF8) [MS-XLS 2.4.28]:
    // - u32 lbPlyPos: absolute stream offset of the sheet substream BOF
    // - u16 grbit: visibility + sheet type
    // - sheet name (short XLUnicodeString)
    //
    // When inserting bytes into the workbook globals stream before worksheet substreams, these
    // absolute offsets must be shifted forward by the insertion length.
    let mut offset = 0usize;
    let mut patched = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len = u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]])
            as usize;
        let data_start = offset + 4;
        let data_end = data_start.saturating_add(len);
        if data_end > workbook_stream.len() {
            break;
        }

        if record_id == RECORD_BOUNDSHEET && len >= 4 {
            let raw = u32::from_le_bytes([
                workbook_stream[data_start],
                workbook_stream[data_start + 1],
                workbook_stream[data_start + 2],
                workbook_stream[data_start + 3],
            ]);
            let adjusted = raw.wrapping_add(inserted_len);
            workbook_stream[data_start..data_start + 4].copy_from_slice(&adjusted.to_le_bytes());
            patched += 1;
        }

        offset = data_end;
    }

    assert!(patched > 0, "expected to patch at least one BOUNDSHEET record");
}

fn build_rc4_standard_encrypted_xls_from_plain_stream(
    workbook_stream_plain: &[u8],
    password: &str,
) -> Vec<u8> {
    // Encrypt an existing BIFF8 workbook stream by inserting FILEPASS and applying RC4 "standard"
    // encryption to all subsequent record payloads.
    //
    // This is used for the password edge-case fixtures derived from `basic.xls`.
    let (filepass_payload, intermediate_key) = build_filepass_rc4_standard_payload(password);

    let Some((record_id, _, bof_end)) = super_read_record(workbook_stream_plain, 0) else {
        panic!("unexpected empty workbook stream");
    };
    assert_eq!(record_id, RECORD_BOF, "expected BOF record at workbook stream offset 0");

    // Ensure the input stream is not already encrypted.
    let mut scan = 0usize;
    while let Some((rid, _, next)) = super_read_record(workbook_stream_plain, scan) {
        assert_ne!(
            rid, RECORD_FILEPASS,
            "expected plaintext stream without FILEPASS; pass an unencrypted fixture"
        );
        scan = next;
    }

    let filepass_record = record(RECORD_FILEPASS, &filepass_payload);
    let inserted_len = filepass_record.len() as u32;

    let mut workbook_stream = Vec::with_capacity(workbook_stream_plain.len() + filepass_record.len());
    workbook_stream.extend_from_slice(&workbook_stream_plain[..bof_end]);
    workbook_stream.extend_from_slice(&filepass_record);
    workbook_stream.extend_from_slice(&workbook_stream_plain[bof_end..]);

    // Patch BOUNDSHEET offsets to account for the inserted FILEPASS bytes.
    patch_boundsheet_offsets_for_inserted_record(&mut workbook_stream, inserted_len);

    // Encrypt payload bytes after FILEPASS.
    let filepass_data_end = bof_end + filepass_record.len();
    let mut cipher = PayloadRc4Standard::new(intermediate_key, 16);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, |data| {
        cipher.apply_keystream(data);
    });

    build_xls_bytes(&workbook_stream)
}

fn xor_password_bytes_method2(password: &str) -> Vec<u8> {
    // MS-OFFCRYPTO 2.3.7.4 "method 2": copy low byte unless zero, else high byte.
    //
    // This encoding is used by some BIFF8 XOR writers when deriving the Method-1 key/verifier.
    let mut bytes = Vec::with_capacity(15);
    for ch in password.encode_utf16() {
        if bytes.len() >= 15 {
            break;
        }
        let lo = (ch & 0x00FF) as u8;
        let hi = (ch >> 8) as u8;
        bytes.push(if lo != 0 { lo } else { hi });
    }
    bytes
}

fn build_xor_encrypted_xls_bytes_with_password_bytes(mut pw_bytes: Vec<u8>) -> Vec<u8> {
    // BIFF8 XOR obfuscation fixture using the MS-OFFCRYPTO/MS-XLS "Method 1" algorithm (the real
    // Excel-compatible scheme).
    pw_bytes.truncate(15);

    let key = create_xor_key_method1(&pw_bytes);
    let verifier = create_password_verifier_method1(&pw_bytes);
    let xor_array = create_xor_array_method1(&pw_bytes, key);

    let filepass_payload = [
        0x00, 0x00, // wEncryptionType (XOR)
        key.to_le_bytes()[0],
        key.to_le_bytes()[1],
        verifier.to_le_bytes()[0],
        verifier.to_le_bytes()[1],
    ];

    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 0);

    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) = super_read_record(&workbook_stream, offset) {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    encrypt_payloads_after_filepass_xor_method1(&mut workbook_stream, filepass_data_end, &xor_array);

    build_xls_bytes(&workbook_stream)
}

fn build_xor_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    use encoding_rs::WINDOWS_1252;

    let (pw_bytes, _, _) = WINDOWS_1252.encode(password);
    build_xor_encrypted_xls_bytes_with_password_bytes(pw_bytes.into_owned())
}

fn build_xor_encrypted_xls_bytes_method2(password: &str) -> Vec<u8> {
    build_xor_encrypted_xls_bytes_with_password_bytes(xor_password_bytes_method2(password))
}

fn build_rc4_standard_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    // BIFF8 RC4 "Standard Encryption" fixture. Ensure payload-after-FILEPASS crosses 1024 bytes by
    // adding a large dummy record in the worksheet substream.

    let (filepass_payload, intermediate_key) = build_filepass_rc4_standard_payload(password);

    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 2048);

    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) = super_read_record(&workbook_stream, offset) {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    let mut cipher = PayloadRc4Standard::new(intermediate_key, 16);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, |data| {
        cipher.apply_keystream(data);
    });

    build_xls_bytes(&workbook_stream)
}

fn build_cryptoapi_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    // Deterministic parameters to keep the fixture stable.
    let salt = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
        0x1E, 0x1F,
    ];

    build_cryptoapi_encrypted_xls_bytes_with_config(
        password,
        &salt,
        4,
        2,
        0,
        None,
    )
}

fn build_cryptoapi_md5_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    let salt: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];
    build_cryptoapi_md5_encrypted_xls_bytes_with_config(password, &salt, 4, 2, 0, None)
}

fn build_cryptoapi_legacy_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    // Deterministic parameters to keep the fixture stable.
    let salt = [
        0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2A, 0x2B, 0x2C, 0x2D, 0x2E,
        0x2F, 0x30,
    ];

    let (filepass_payload, key_material) =
        build_filepass_cryptoapi_legacy_payload(password, &salt, 1, 1, 0, None);
    // Ensure payload-after-FILEPASS crosses 1024 bytes so the legacy CryptoAPI absolute-offset
    // decryptor has to re-key mid-stream.
    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 2048);

    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) = super_read_record(&workbook_stream, offset) {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    encrypt_payloads_after_filepass_cryptoapi_legacy(
        &mut workbook_stream,
        filepass_data_end,
        &key_material,
        16,
    );

    build_xls_bytes(&workbook_stream)
}

fn build_cryptoapi_encrypted_xls_bytes_with_config(
    password: &str,
    salt: &[u8; 16],
    version_major: u16,
    version_minor: u16,
    provider_type: u32,
    csp_name: Option<&str>,
) -> Vec<u8> {
    // Build a minimal BIFF8 workbook stream with one sheet containing A1=42, then encrypt all
    // record payload bytes after FILEPASS using RC4 CryptoAPI.

    let filepass_payload = build_filepass_cryptoapi_payload(
        password,
        salt,
        version_major,
        version_minor,
        provider_type,
        csp_name,
    );
    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 0);
    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) = super_read_record(&workbook_stream, offset) {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    let key_material = derive_cryptoapi_key_material(password, salt);
    let mut cipher = PayloadRc4::new(key_material, 16);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, |data| {
        cipher.apply_keystream(data);
    });

    build_xls_bytes(&workbook_stream)
}

fn build_cryptoapi_md5_encrypted_xls_bytes_with_config(
    password: &str,
    salt: &[u8; 16],
    version_major: u16,
    version_minor: u16,
    provider_type: u32,
    csp_name: Option<&str>,
) -> Vec<u8> {
    let filepass_payload = build_filepass_cryptoapi_md5_payload(
        password,
        salt,
        version_major,
        version_minor,
        provider_type,
        csp_name,
    );
    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 0);
    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) = super_read_record(&workbook_stream, offset) {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    let key_material = derive_cryptoapi_key_material_md5(password, salt);
    let mut cipher = PayloadRc4Md5::new(key_material, 16);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, |data| {
        cipher.apply_keystream(data);
    });
    build_xls_bytes(&workbook_stream)
}

// Local copy of the record reader used by `import_encrypted.rs` so we can locate FILEPASS offsets
// without depending on that integration test module.
fn super_read_record(stream: &[u8], offset: usize) -> Option<(u16, &[u8], usize)> {
    if offset + 4 > stream.len() {
        return None;
    }
    let record_id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
    let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
    let data_start = offset + 4;
    let data_end = data_start.checked_add(len)?;
    let data = stream.get(data_start..data_end)?;
    Some((record_id, data, data_end))
}

#[test]
#[ignore]
fn regenerate_encrypted_xls_fixtures() {
    let fixtures_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted");
    std::fs::create_dir_all(&fixtures_dir).expect("create encrypted fixtures dir");

    // Base workbook used by password edge-case fixtures (multi-sheet, strings, formulas).
    let basic_fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("basic.xls");
    let basic_workbook_stream = read_workbook_stream_from_xls(&basic_fixture_path);

    // Decryptable BIFF8 XOR fixture.
    let xor_path = fixtures_dir.join("biff8_xor_pw_open.xls");
    let xor_bytes = build_xor_encrypted_xls_bytes("password");
    std::fs::write(&xor_path, xor_bytes)
        .unwrap_or_else(|err| panic!("write encrypted fixture {xor_path:?} failed: {err}"));

    // Unicode-password variant (non-ASCII).
    let xor_unicode_path = fixtures_dir.join("biff8_xor_unicode_pw_open.xls");
    let xor_unicode_bytes = build_xor_encrypted_xls_bytes("pässwörd");
    std::fs::write(&xor_unicode_path, xor_unicode_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {xor_unicode_path:?} failed: {err}");
    });

    // XOR long-password fixture (15-byte truncation semantics).
    let xor_long_path = fixtures_dir.join("biff8_xor_pw_open_long_password.xls");
    let xor_long_bytes = build_xor_encrypted_xls_bytes("0123456789abcdef");
    std::fs::write(&xor_long_path, xor_long_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {xor_long_path:?} failed: {err}");
    });

    // XOR Method-1 password derivation method-2 fixture: uses a Unicode password that is not
    // representable in Windows-1252, so decryptors must fall back to MS-OFFCRYPTO 2.3.7.4 "method 2"
    // password bytes.
    let xor_unicode_method2_path = fixtures_dir.join("biff8_xor_pw_open_unicode_method2.xls");
    let xor_unicode_method2_bytes = build_xor_encrypted_xls_bytes_method2("Ā");
    std::fs::write(&xor_unicode_method2_path, xor_unicode_method2_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {xor_unicode_method2_path:?} failed: {err}");
    });

    // Empty-password XOR fixture (some writers can emit this; Excel UI may refuse to create it).
    let xor_empty_path = fixtures_dir.join("biff8_xor_pw_open_empty_password.xls");
    let xor_empty_bytes = build_xor_encrypted_xls_bytes("");
    std::fs::write(&xor_empty_path, xor_empty_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {xor_empty_path:?} failed: {err}");
    });

    // Decryptable BIFF8 RC4 Standard fixture.
    let rc4_standard_path = fixtures_dir.join("biff8_rc4_standard_pw_open.xls");
    let rc4_standard_bytes = build_rc4_standard_encrypted_xls_bytes("password");
    std::fs::write(&rc4_standard_path, rc4_standard_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {rc4_standard_path:?} failed: {err}");
    });

    // Unicode-password variant (non-ASCII). This ensures RC4 Standard derives keys from UTF-16LE,
    // not a narrow codepage.
    let rc4_standard_unicode_path = fixtures_dir.join("biff8_rc4_standard_unicode_pw_open.xls");
    let rc4_standard_unicode_bytes = build_rc4_standard_encrypted_xls_bytes("pässwörd");
    std::fs::write(&rc4_standard_unicode_path, rc4_standard_unicode_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {rc4_standard_unicode_path:?} failed: {err}");
    });

    // Unicode + emoji password variant (non-BMP / UTF-16 surrogate pair).
    let rc4_standard_unicode_emoji_path =
        fixtures_dir.join("biff8_rc4_standard_unicode_emoji_pw_open.xls");
    let rc4_standard_unicode_emoji_bytes = build_rc4_standard_encrypted_xls_bytes("pässwörd🔒");
    std::fs::write(&rc4_standard_unicode_emoji_path, rc4_standard_unicode_emoji_bytes)
        .unwrap_or_else(|err| {
            panic!("write encrypted fixture {rc4_standard_unicode_emoji_path:?} failed: {err}");
        });

    // Surrogate-pair truncation edge case: RC4 Standard only uses the first 15 UTF-16 code units of
    // the password. With 14 ASCII chars + a non-BMP emoji, the truncation boundary falls *inside*
    // the emoji's surrogate pair (high surrogate retained, low surrogate ignored).
    let rc4_standard_surrogate_truncation_path =
        fixtures_dir.join("biff8_rc4_standard_surrogate_truncation_pw_open.xls");
    let rc4_standard_surrogate_truncation_bytes =
        build_rc4_standard_encrypted_xls_bytes("0123456789ABCD🔒");
    std::fs::write(
        &rc4_standard_surrogate_truncation_path,
        rc4_standard_surrogate_truncation_bytes,
    )
    .unwrap_or_else(|err| {
        panic!(
            "write encrypted fixture {rc4_standard_surrogate_truncation_path:?} failed: {err}"
        );
    });

    // RC4 Standard edge-case fixtures derived from `basic.xls`.
    let rc4_standard_long_path = fixtures_dir.join("biff8_rc4_standard_pw_open_long_password.xls");
    let rc4_standard_long_bytes =
        build_rc4_standard_encrypted_xls_from_plain_stream(&basic_workbook_stream, "0123456789abcdef");
    std::fs::write(&rc4_standard_long_path, rc4_standard_long_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {rc4_standard_long_path:?} failed: {err}");
    });

    let rc4_standard_empty_path =
        fixtures_dir.join("biff8_rc4_standard_pw_open_empty_password.xls");
    let rc4_standard_empty_bytes =
        build_rc4_standard_encrypted_xls_from_plain_stream(&basic_workbook_stream, "");
    std::fs::write(&rc4_standard_empty_path, rc4_standard_empty_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {rc4_standard_empty_path:?} failed: {err}");
    });

    // Decryptable BIFF8 RC4 CryptoAPI fixture.
    let cryptoapi_path = fixtures_dir.join("biff8_rc4_cryptoapi_pw_open.xls");
    let cryptoapi_bytes = build_cryptoapi_encrypted_xls_bytes("correct horse battery staple");
    std::fs::write(&cryptoapi_path, cryptoapi_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_path:?} failed: {err}");
    });

    // Empty-password CryptoAPI fixture (some writers can emit this; Excel UI may refuse to create it).
    let cryptoapi_empty_path = fixtures_dir.join("biff8_rc4_cryptoapi_pw_open_empty_password.xls");
    let cryptoapi_empty_bytes = build_cryptoapi_encrypted_xls_bytes("");
    std::fs::write(&cryptoapi_empty_path, cryptoapi_empty_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_empty_path:?} failed: {err}");
    });

    // Decryptable BIFF8 RC4 CryptoAPI fixture using MD5 for password hashing / verifier hashing.
    let cryptoapi_md5_path = fixtures_dir.join("biff8_rc4_cryptoapi_md5_pw_open.xls");
    let cryptoapi_md5_bytes = build_cryptoapi_md5_encrypted_xls_bytes("password");
    std::fs::write(&cryptoapi_md5_path, cryptoapi_md5_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_md5_path:?} failed: {err}");
    });

    // Unicode-password variant (non-ASCII).
    let cryptoapi_unicode_path = fixtures_dir.join("biff8_rc4_cryptoapi_unicode_pw_open.xls");
    let cryptoapi_unicode_bytes = build_cryptoapi_encrypted_xls_bytes("pässwörd");
    std::fs::write(&cryptoapi_unicode_path, cryptoapi_unicode_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_unicode_path:?} failed: {err}");
    });

    // Unicode + emoji password variant (non-BMP / UTF-16 surrogate pair).
    let cryptoapi_unicode_emoji_path =
        fixtures_dir.join("biff8_rc4_cryptoapi_unicode_emoji_pw_open.xls");
    let cryptoapi_unicode_emoji_bytes = build_cryptoapi_encrypted_xls_bytes("pässwörd🔒");
    std::fs::write(&cryptoapi_unicode_emoji_path, cryptoapi_unicode_emoji_bytes)
        .unwrap_or_else(|err| {
            panic!("write encrypted fixture {cryptoapi_unicode_emoji_path:?} failed: {err}");
        });

    // Legacy FILEPASS layout variant (`wEncryptionInfo = 0x0004`).
    let cryptoapi_legacy_unicode_emoji_path = fixtures_dir
        .join("biff8_rc4_cryptoapi_legacy_unicode_emoji_pw_open.xls");
    let cryptoapi_legacy_unicode_emoji_bytes = build_cryptoapi_legacy_encrypted_xls_bytes("pässwörd🔒");
    std::fs::write(
        &cryptoapi_legacy_unicode_emoji_path,
        cryptoapi_legacy_unicode_emoji_bytes,
    )
    .unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_legacy_unicode_emoji_path:?} failed: {err}");
    });

    let cryptoapi_legacy_empty_path =
        fixtures_dir.join("biff8_rc4_cryptoapi_legacy_pw_open_empty_password.xls");
    let cryptoapi_legacy_empty_bytes = build_cryptoapi_legacy_encrypted_xls_bytes("");
    std::fs::write(&cryptoapi_legacy_empty_path, cryptoapi_legacy_empty_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_legacy_empty_path:?} failed: {err}");
    });

    // Non-ASCII password fixture used to validate Unicode password handling.
    let unicode_fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_rc4cryptoapi_non_ascii_password.xls");
    // Use stable, deterministic encryption parameters but with an "Excel-like" CryptoAPI header
    // (version 1.1 + CSPName) to exercise the parser.
    let unicode_salt: [u8; 16] = [
        0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
        0xDC, 0xFE,
    ];
    let unicode_bytes = build_cryptoapi_encrypted_xls_bytes_with_config(
        "pässwörd",
        &unicode_salt,
        1,
        1,
        1,
        Some("Microsoft Base Cryptographic Provider v1.0"),
    );
    std::fs::write(&unicode_fixture_path, unicode_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {unicode_fixture_path:?} failed: {err}");
    });
}
