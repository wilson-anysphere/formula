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

use formula_model::hash_legacy_password;
use md5::{Digest as _, Md5};
use sha1::Sha1;
use std::io::{Cursor, Write};
use std::path::PathBuf;

const RECORD_BOF: u16 = 0x0809;
const RECORD_EOF: u16 = 0x000A;
const RECORD_FILEPASS: u16 = 0x002F;

const RECORD_CODEPAGE: u16 = 0x0042;
const RECORD_WINDOW1: u16 = 0x003D;
const RECORD_FONT: u16 = 0x0031;
const RECORD_XF: u16 = 0x00E0;
const RECORD_BOUNDSHEET: u16 = 0x0085;

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
    assert!(bytes.len() <= u8::MAX as usize, "string too long for BIFF8 short string");
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

fn xf_record(font_idx: u16, fmt_idx: u16, is_style_xf: bool, alignment: u8) -> [u8; 20] {
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
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true, 0x20));
    }
    // Cell XF used by the sheet's NUMBER record. Make it non-default (vertical alignment = Top)
    // so the decrypt + BIFF parser tests can assert that XF metadata after FILEPASS was imported.
    let xf_cell: u16 = 16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false, 0x00));

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

fn derive_cryptoapi_key_material(password: &str, salt: &[u8; 16]) -> [u8; 20] {
    // Matches `crates/formula-xls/src/decrypt.rs`.
    const PASSWORD_HASH_ITERATIONS: u32 = 50_000;

    let pw_bytes = utf16le_bytes(password);
    let mut hash = sha1_bytes(&[salt, &pw_bytes]);

    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        hash = sha1_bytes(&[&iter, &hash]);
    }

    hash
}

fn derive_cryptoapi_block_key(key_material: &[u8; 20], block: u32, key_len: usize) -> Vec<u8> {
    let block_bytes = block.to_le_bytes();
    let digest = sha1_bytes(&[key_material, &block_bytes]);
    if key_len == 5 {
        // CryptoAPI 40-bit RC4 keys are expressed as a 128-bit key where the high 88 bits are zero.
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

struct PayloadRc4 {
    key_material: [u8; 20],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
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

struct PayloadXor {
    key_bytes: [u8; 2],
    xor_array: [u8; 16],
    pos: usize,
}

impl PayloadXor {
    fn new(key: u16, password: &str) -> Self {
        Self {
            key_bytes: key.to_le_bytes(),
            xor_array: derive_xor_array(password),
            pos: 0,
        }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data.iter_mut() {
            let ks = self.xor_array[self.pos % self.xor_array.len()] ^ self.key_bytes[self.pos % 2];
            *b ^= ks;
            self.pos = self.pos.saturating_add(1);
        }
    }
}

fn derive_xor_array(password: &str) -> [u8; 16] {
    // Mirrors the legacy BIFF XOR obfuscation "XOR array" used by `crates/formula-xls`.
    const PAD: [u8; 16] = [
        0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0xFF, 0xFF, 0xB8, 0xFF, 0xFF, 0xB7, 0xFF,
        0xFF, 0xB6,
    ];

    let mut out = PAD;
    for (i, ch) in password.encode_utf16().take(out.len()).enumerate() {
        out[i] ^= (ch & 0xFF) as u8;
    }
    out
}

fn build_filepass_cryptoapi_payload(password: &str) -> Vec<u8> {
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

    let salt: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
        0x1E, 0x1F,
    ];

    // Deterministic verifier bytes.
    let verifier_plain: [u8; 16] = [
        0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D,
        0x1E, 0x0F,
    ];
    let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

    // Encrypt verifier + hash using block 0 key.
    let key_material = derive_cryptoapi_key_material(password, &salt);
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

    // EncryptionHeader (32 bytes) [MS-OFFCRYPTO].
    let mut header = Vec::<u8>::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&128u32.to_le_bytes()); // KeySize bits
    header.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

    // EncryptionVerifier.
    let mut verifier = Vec::<u8>::new();
    verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier.extend_from_slice(&salt);
    verifier.extend_from_slice(&encrypted_verifier);
    verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
    verifier.extend_from_slice(&encrypted_verifier_hash);

    // EncryptionInfo:
    //   u16 MajorVersion, u16 MinorVersion, u32 Flags, u32 HeaderSize, EncryptionHeader, EncryptionVerifier.
    let mut enc_info = Vec::<u8>::new();
    enc_info.extend_from_slice(&4u16.to_le_bytes()); // Major
    enc_info.extend_from_slice(&2u16.to_le_bytes()); // Minor
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

fn build_xor_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    // BIFF8 XOR obfuscation fixture.
    const KEY: u16 = 0x1234;
    let verifier = hash_legacy_password(password);
    let filepass_payload = [
        0x00, 0x00, // wEncryptionType (XOR)
        KEY.to_le_bytes()[0],
        KEY.to_le_bytes()[1],
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

    let mut cipher = PayloadXor::new(KEY, password);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, |data| {
        cipher.apply_keystream(data);
    });

    build_xls_bytes(&workbook_stream)
}

fn build_rc4_standard_encrypted_xls_bytes(password: &str) -> Vec<u8> {
    // BIFF8 RC4 "Standard Encryption" fixture. Ensure payload-after-FILEPASS crosses 1024 bytes by
    // adding a large dummy record in the worksheet substream.

    // Deterministic "DocId" / salt bytes.
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

    let mut filepass_payload = Vec::<u8>::new();
    filepass_payload.extend_from_slice(&[0x01, 0x00]); // wEncryptionType (RC4)
    filepass_payload.extend_from_slice(&[0x01, 0x00]); // major (subType)
    filepass_payload.extend_from_slice(&[0x02, 0x00]); // minor (128-bit)
    filepass_payload.extend_from_slice(&salt);
    filepass_payload.extend_from_slice(&encrypted_verifier);
    filepass_payload.extend_from_slice(&encrypted_verifier_hash);

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
    // Build a minimal BIFF8 workbook stream with one sheet containing A1=42, then encrypt all
    // record payload bytes after FILEPASS using RC4 CryptoAPI.

    let filepass_payload = build_filepass_cryptoapi_payload(password);
    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload, 0);
    let mut offset = 0usize;
    let mut filepass_data_end = None::<usize>;
    while let Some((record_id, payload, next)) =
        super_read_record(&workbook_stream, offset)
    {
        if record_id == RECORD_FILEPASS {
            filepass_data_end = Some(offset + 4 + payload.len());
            break;
        }
        offset = next;
    }
    let filepass_data_end =
        filepass_data_end.expect("generated workbook stream should contain FILEPASS");

    let salt = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
        0x1E, 0x1F,
    ];
    let key_material = derive_cryptoapi_key_material(password, &salt);
    let mut cipher = PayloadRc4::new(key_material, 16);
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

    // Decryptable BIFF8 XOR fixture.
    let xor_path = fixtures_dir.join("biff8_xor_pw_open.xls");
    let xor_bytes = build_xor_encrypted_xls_bytes("password");
    std::fs::write(&xor_path, xor_bytes)
        .unwrap_or_else(|err| panic!("write encrypted fixture {xor_path:?} failed: {err}"));

    // Decryptable BIFF8 RC4 Standard fixture.
    let rc4_standard_path = fixtures_dir.join("biff8_rc4_standard_pw_open.xls");
    let rc4_standard_bytes = build_rc4_standard_encrypted_xls_bytes("password");
    std::fs::write(&rc4_standard_path, rc4_standard_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {rc4_standard_path:?} failed: {err}");
    });

    // Decryptable BIFF8 RC4 CryptoAPI fixture.
    let cryptoapi_path = fixtures_dir.join("biff8_rc4_cryptoapi_pw_open.xls");
    let cryptoapi_bytes = build_cryptoapi_encrypted_xls_bytes("correct horse battery staple");
    std::fs::write(&cryptoapi_path, cryptoapi_bytes).unwrap_or_else(|err| {
        panic!("write encrypted fixture {cryptoapi_path:?} failed: {err}");
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
}
