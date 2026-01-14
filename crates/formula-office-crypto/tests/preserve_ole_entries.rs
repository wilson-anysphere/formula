use std::io::{Cursor, Read, Seek, Write};

use formula_office_crypto::{
    decrypt_encrypted_package_ole, encrypt_package_to_ole, encrypt_package_to_ole_with_entries,
    extract_ole_entries, is_encrypted_ooxml_ole, EncryptOptions,
};

const SUMMARY_INFORMATION: &str = "\u{0005}SummaryInformation";

fn has_stream(entries: &formula_office_crypto::OleEntries, name: &str) -> bool {
    entries.streams.iter().any(|s| {
        let path = s.path.to_string_lossy();
        let path = path.strip_prefix('/').unwrap_or(&path);
        path.eq_ignore_ascii_case(name)
    })
}

fn has_storage(entries: &formula_office_crypto::OleEntries, name: &str) -> bool {
    entries.storages.iter().any(|p| {
        let path = p.to_string_lossy();
        let path = path.strip_prefix('/').unwrap_or(&path);
        path.eq_ignore_ascii_case(name)
    })
}

fn read_stream_best_effort<R: Read + Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Vec<u8> {
    let mut out = Vec::new();
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .unwrap_or_else(|err| panic!("open stream {name}: {err}"))
        .read_to_end(&mut out)
        .unwrap_or_else(|err| panic!("read stream {name}: {err}"));
    out
}

#[test]
fn encrypt_package_to_ole_with_entries_preserves_extra_ole_streams_and_storages() {
    let password = "password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Keep tests reasonably fast while still exercising the full encrypt/decrypt pipeline.
    let mut opts = EncryptOptions::default();
    opts.spin_count = 1_000;

    let encrypted = encrypt_package_to_ole(plaintext, password, opts.clone()).expect("encrypt");
    assert!(is_encrypted_ooxml_ole(&encrypted));

    // Add some extra OLE metadata streams/storages (these should be preserved when we re-wrap).
    let cursor = Cursor::new(encrypted);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open encrypted OLE");

    ole.create_stream(SUMMARY_INFORMATION)
        .expect("create SummaryInformation")
        .write_all(b"dummy summary information bytes")
        .expect("write SummaryInformation");

    ole.create_storage("foo").expect("create foo storage");
    ole.create_stream("foo/bar.txt")
        .expect("create foo/bar.txt stream")
        .write_all(b"nested stream payload")
        .expect("write foo/bar.txt");

    let encrypted_with_extra = ole.into_inner().into_inner();
    assert!(is_encrypted_ooxml_ole(&encrypted_with_extra));

    // Extract extra entries (should exclude the encryption streams).
    let mut ole = cfb::CompoundFile::open(Cursor::new(&encrypted_with_extra))
        .expect("open encrypted OLE for extraction");
    let entries = extract_ole_entries(&mut ole).expect("extract OLE entries");
    assert!(
        !has_stream(&entries, "EncryptionInfo"),
        "extract_ole_entries must exclude EncryptionInfo"
    );
    assert!(
        !has_stream(&entries, "EncryptedPackage"),
        "extract_ole_entries must exclude EncryptedPackage"
    );
    assert!(has_stream(&entries, SUMMARY_INFORMATION));
    assert!(has_storage(&entries, "foo"));
    assert!(has_stream(&entries, "foo/bar.txt"));

    // Re-encrypt with the preserved entries copied into the new container.
    let rewrapped = encrypt_package_to_ole_with_entries(plaintext, password, opts, Some(&entries))
        .expect("encrypt_package_to_ole_with_entries");
    assert!(is_encrypted_ooxml_ole(&rewrapped));

    // Confirm the extra streams were copied into the output OLE.
    let mut out_ole = cfb::CompoundFile::open(Cursor::new(&rewrapped)).expect("open output OLE");
    let summary = read_stream_best_effort(&mut out_ole, SUMMARY_INFORMATION);
    assert_eq!(summary, b"dummy summary information bytes");
    let nested = read_stream_best_effort(&mut out_ole, "foo/bar.txt");
    assert_eq!(nested, b"nested stream payload");

    // Sanity check: decrypting should yield the original plaintext ZIP bytes.
    let decrypted = decrypt_encrypted_package_ole(&rewrapped, password).expect("decrypt rewrapped");
    assert_eq!(decrypted, plaintext);
    assert!(
        decrypted.starts_with(b"PK"),
        "expected decrypted payload to be a ZIP package"
    );
}
