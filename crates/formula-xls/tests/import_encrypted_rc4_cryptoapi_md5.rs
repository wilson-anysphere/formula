use std::io::{Cursor, Write};

use formula_model::{CellValue};

use md5::{Digest as _, Md5};

// Record ids.
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

// BOF substream types.
const BOF_VERSION_BIFF8: u16 = 0x0600;
const BOF_DT_WORKBOOK_GLOBALS: u16 = 0x0005;
const BOF_DT_WORKSHEET: u16 = 0x0010;

const XF_FLAG_LOCKED: u16 = 0x0001;
const XF_FLAG_STYLE: u16 = 0x0004;

// CryptoAPI constants.
const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
const ENCRYPTION_SUBTYPE_CRYPTOAPI: u16 = 0x0002;
const CALG_RC4: u32 = 0x0000_6801;
const CALG_MD5: u32 = 0x0000_8003;
const PASSWORD_HASH_ITERATIONS: u32 = 50_000;
const PAYLOAD_BLOCK_SIZE: usize = 1024;

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
    // Matches the minimal BOF payload used across other `.xls` fixtures in this crate.
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&BOF_VERSION_BIFF8.to_le_bytes());
    out[2..4].copy_from_slice(&dt.to_le_bytes());
    out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
    out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
    out
}

fn window1() -> [u8; 18] {
    let mut out = [0u8; 18];
    out[14..16].copy_from_slice(&1u16.to_le_bytes()); // cTabSel
    out[16..18].copy_from_slice(&600u16.to_le_bytes()); // wTabRatio
    out
}

fn window2() -> [u8; 18] {
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

fn xf_record(font_idx: u16, fmt_idx: u16, is_style_xf: bool) -> [u8; 20] {
    let mut out = [0u8; 20];
    out[0..2].copy_from_slice(&font_idx.to_le_bytes());
    out[2..4].copy_from_slice(&fmt_idx.to_le_bytes());

    let flags: u16 = XF_FLAG_LOCKED | if is_style_xf { XF_FLAG_STYLE } else { 0 };
    out[4..6].copy_from_slice(&flags.to_le_bytes());

    // Default BIFF8 alignment: General + Bottom.
    out[6] = 0x20;
    // Attribute flags: apply all so XFs don't rely on inheritance.
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

fn build_plain_biff8_workbook_stream(filepass_payload: &[u8]) -> Vec<u8> {
    // Minimal workbook structure mirroring the encrypted fixtures generator.
    let mut globals = Vec::<u8>::new();
    push_record(&mut globals, RECORD_BOF, &bof_biff8(BOF_DT_WORKBOOK_GLOBALS));
    push_record(&mut globals, RECORD_FILEPASS, filepass_payload);
    push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
    push_record(&mut globals, RECORD_WINDOW1, &window1());
    push_record(&mut globals, RECORD_FONT, &font("Arial"));

    // XF table: many readers expect 16 style XFs before any cell XFs.
    for _ in 0..16 {
        push_record(&mut globals, RECORD_XF, &xf_record(0, 0, true));
    }
    // One default cell XF (General).
    let xf_cell: u16 = 16;
    push_record(&mut globals, RECORD_XF, &xf_record(0, 0, false));

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
    push_record(&mut sheet, RECORD_EOF, &[]);

    // Patch BoundSheet offset to point at the sheet BOF.
    globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
        .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
    globals.extend_from_slice(&sheet);
    globals
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len().saturating_mul(2));
    for unit in s.encode_utf16() {
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

fn derive_cryptoapi_key_material_md5(password: &str, salt: &[u8; 16]) -> [u8; 16] {
    let pw_bytes = utf16le_bytes(password);
    let mut hash = md5_bytes(&[salt, &pw_bytes]);
    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        hash = md5_bytes(&[&iter, &hash]);
    }
    hash
}

fn derive_cryptoapi_block_key_md5(key_material: &[u8; 16], block: u32) -> [u8; 16] {
    md5_bytes(&[key_material, &block.to_le_bytes()])
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

struct PayloadRc4Md5 {
    key_material: [u8; 16],
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4Md5 {
    fn new(key_material: [u8; 16]) -> Self {
        let key0 = derive_cryptoapi_block_key_md5(&key_material, 0);
        Self {
            key_material,
            block: 0,
            pos_in_block: 0,
            rc4: Rc4::new(&key0),
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_cryptoapi_block_key_md5(&self.key_material, self.block);
        self.rc4 = Rc4::new(&key);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
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

fn build_filepass_cryptoapi_md5_payload(password: &str) -> (Vec<u8>, [u8; 16]) {
    // FILEPASS payload layout (CryptoAPI) [MS-XLS 2.4.105]:
    //   u16 wEncryptionType = 0x0001 (RC4)
    //   u16 wEncryptionSubType = 0x0002 (CryptoAPI)
    //   u32 dwEncryptionInfoLen
    //   EncryptionInfo bytes
    //
    // EncryptionInfo matches [MS-OFFCRYPTO] RC4 CryptoAPI.
    let salt: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];

    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];
    let verifier_hash_plain = md5_bytes(&[&verifier_plain]);

    // Encrypt verifier + hash using block 0 key.
    let key_material = derive_cryptoapi_key_material_md5(password, &salt);
    let key0 = derive_cryptoapi_block_key_md5(&key_material, 0);
    let mut rc4 = Rc4::new(&key0);

    let mut verifier_buf = [0u8; 32];
    verifier_buf[..16].copy_from_slice(&verifier_plain);
    verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
    rc4.apply_keystream(&mut verifier_buf);

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
    let mut encrypted_verifier_hash = [0u8; 16];
    encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

    // EncryptionHeader (32 bytes) [MS-OFFCRYPTO].
    let mut header = Vec::<u8>::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_MD5.to_le_bytes()); // AlgIDHash
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
    (payload, key_material)
}

fn find_filepass_data_end(workbook_stream: &[u8]) -> usize {
    let mut offset = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(data_end <= workbook_stream.len(), "truncated record");
        if record_id == RECORD_FILEPASS {
            return data_end;
        }
        offset = data_end;
    }
    panic!("FILEPASS not found");
}

fn encrypt_payloads_after_filepass(workbook_stream: &mut [u8], filepass_data_end: usize, key_material: [u8; 16]) {
    let mut cipher = PayloadRc4Md5::new(key_material);
    let mut cursor = filepass_data_end;
    while cursor < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(cursor);
        if remaining < 4 {
            break;
        }
        let len = u16::from_le_bytes([workbook_stream[cursor + 2], workbook_stream[cursor + 3]]) as usize;
        let data_start = cursor + 4;
        let data_end = data_start + len;
        assert!(data_end <= workbook_stream.len(), "generated BIFF record extends past end of stream");
        cipher.apply_keystream(&mut workbook_stream[data_start..data_end]);
        cursor = data_end;
    }
}

fn build_xls_bytes(workbook_stream: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(workbook_stream)
            .expect("write Workbook stream bytes");
    }
    ole.into_inner().into_inner()
}

#[test]
fn decrypts_rc4_cryptoapi_md5_end_to_end() {
    const PASSWORD: &str = "password";
    const WRONG_PASSWORD: &str = "wrong password";

    let (filepass_payload, key_material) = build_filepass_cryptoapi_md5_payload(PASSWORD);
    let mut workbook_stream = build_plain_biff8_workbook_stream(&filepass_payload);
    let filepass_data_end = find_filepass_data_end(&workbook_stream);
    encrypt_payloads_after_filepass(&mut workbook_stream, filepass_data_end, key_material);

    let xls_bytes = build_xls_bytes(&workbook_stream);
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&xls_bytes).expect("write xls bytes");

    // Correct password decrypts and imports.
    let result =
        formula_xls::import_xls_path_with_password(tmp.path(), Some(PASSWORD)).expect("import ok");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));

    // Wrong password is surfaced as ImportError::InvalidPassword.
    let err = formula_xls::import_xls_path_with_password(tmp.path(), Some(WRONG_PASSWORD))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

