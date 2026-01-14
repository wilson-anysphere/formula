#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_office_crypto::{
    extract_ole_entries, OfficeCryptoError, MAX_OLE_PRESERVED_STREAM_BYTES,
};

#[test]
fn extract_ole_entries_rejects_oversized_preserved_stream() {
    // Build an OLE container with the expected encryption streams plus one large extra stream.
    // The preservation helper should reject it instead of allocating unbounded memory.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage");

    {
        let mut stream = ole.create_stream("BigStream").expect("create stream");
        let chunk = vec![0xA5u8; 64 * 1024];
        let mut remaining = MAX_OLE_PRESERVED_STREAM_BYTES + 1;
        while remaining > 0 {
            let n = remaining.min(chunk.len());
            stream.write_all(&chunk[..n]).expect("write chunk");
            remaining -= n;
        }
    }

    let bytes = ole.into_inner().into_inner();
    let mut ole = cfb::CompoundFile::open(Cursor::new(bytes)).expect("open cfb");

    let err = extract_ole_entries(&mut ole).expect_err("expected size limit error");
    assert!(
        matches!(
            err,
            OfficeCryptoError::SizeLimitExceeded {
                context: "OLE preserved stream",
                limit
            } if limit == MAX_OLE_PRESERVED_STREAM_BYTES
        ),
        "unexpected error: {err:?}"
    );
}
