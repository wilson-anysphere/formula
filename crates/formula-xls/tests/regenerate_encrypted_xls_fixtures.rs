//! Regenerates the encrypted `.xls` fixtures under `tests/fixtures/encrypted/`.
//!
//! This is an ignored test so it doesn't run in CI; it's a convenient, in-repo way to keep the
//! binary fixture blobs reproducible and auditable.
//!
//! Run:
//!   cargo test -p formula-xls --test regenerate_encrypted_xls_fixtures -- --ignored

use std::io::{Cursor, Write};
use std::path::PathBuf;

fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn bof_workbook_globals_biff8() -> [u8; 16] {
    // Match the BOF payload used by `tests/common/xls_fixture_builder.rs` so the resulting BIFF
    // streams are recognizable as BIFF8.
    let mut out = [0u8; 16];
    out[0..2].copy_from_slice(&0x0600u16.to_le_bytes()); // BIFF8
    out[2..4].copy_from_slice(&0x0005u16.to_le_bytes()); // workbook globals substream
    out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
    out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
    out
}

fn workbook_stream_with_filepass(filepass_payload: &[u8]) -> Vec<u8> {
    const RECORD_BOF: u16 = 0x0809;
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_EOF: u16 = 0x000A;

    [
        record(RECORD_BOF, &bof_workbook_globals_biff8()),
        record(RECORD_FILEPASS, filepass_payload),
        record(RECORD_EOF, &[]),
    ]
    .concat()
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

#[test]
#[ignore]
fn regenerate_encrypted_xls_fixtures() {
    let fixtures_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted");
    std::fs::create_dir_all(&fixtures_dir).expect("create encrypted fixtures dir");

    // FILEPASS payloads are intentionally minimal; `formula-xls` currently only needs to observe
    // that `FILEPASS` exists to treat the workbook as encrypted.
    //
    // FILEPASS payload layouts we care about for classification:
    //
    // - BIFF8 XOR obfuscation:
    //   wEncryptionType (0x0000) + key (u16) + verifier (u16)
    //
    // - BIFF8 RC4:
    //   wEncryptionType (0x0001) + subType (0x0001) + opaque algorithm payload
    //
    // - BIFF8 RC4 CryptoAPI:
    //   wEncryptionType (0x0001) + subType (0x0002) + opaque algorithm payload
    //
    // We intentionally keep the algorithm-specific bytes synthetic/deterministic; the importer
    // currently only needs to classify the variant.
    let xor_payload = [0x00, 0x00, 0x34, 0x12, 0x78, 0x56]; // type + key + verifier

    // Use a 52-byte payload for RC4: 4-byte header + 48 bytes of deterministic filler.
    let mut rc4_standard_payload = Vec::with_capacity(4 + 48);
    rc4_standard_payload.extend_from_slice(&[
        0x01, 0x00, // wEncryptionType (RC4)
        0x01, 0x00, // subType (RC4)
    ]);
    rc4_standard_payload.extend(0u8..48u8);

    // Use a 68-byte payload for CryptoAPI: 4-byte header + 64 bytes of deterministic filler.
    let mut rc4_cryptoapi_payload = Vec::with_capacity(4 + 64);
    rc4_cryptoapi_payload.extend_from_slice(&[
        0x01, 0x00, // wEncryptionType (RC4)
        0x02, 0x00, // subType (CryptoAPI)
    ]);
    rc4_cryptoapi_payload.extend(0xA0u8..0xE0u8);

    let fixtures: [(&str, Vec<u8>); 3] = [
        ("biff8_xor_pw_open.xls", xor_payload.to_vec()),
        ("biff8_rc4_standard_pw_open.xls", rc4_standard_payload),
        ("biff8_rc4_cryptoapi_pw_open.xls", rc4_cryptoapi_payload),
    ];

    for (filename, filepass_payload) in fixtures {
        let workbook_stream = workbook_stream_with_filepass(&filepass_payload);
        let bytes = build_xls_bytes(&workbook_stream);

        let path = fixtures_dir.join(filename);
        std::fs::write(&path, bytes).unwrap_or_else(|err| {
            panic!("write encrypted fixture {path:?} failed: {err}");
        });
    }
}
