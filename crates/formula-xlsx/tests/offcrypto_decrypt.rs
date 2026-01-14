use std::io::{Cursor, Write as _};

use formula_xlsx::{decrypt_agile_ooxml_from_cfb, OffCryptoError};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng};

fn small_zip_bytes() -> Vec<u8> {
    use zip::write::FileOptions;

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let options = FileOptions::<()>::default();
        zip.start_file("hello.txt", options)
            .expect("start zip file");
        zip.write_all(b"hello offcrypto")
            .expect("write zip file bytes");
        zip.finish().expect("finish zip");
    }
    cursor.into_inner()
}

#[test]
fn decrypt_agile_ooxml_from_cfb_roundtrip() {
    let password = "Password1234_";
    let plaintext_zip = small_zip_bytes();

    let rng_seed = [0u8; 32];
    let mut rng = StdRng::from_seed(rng_seed);
    let cursor = Cursor::new(Vec::new());

    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext_zip)
        .expect("write plaintext zip bytes");

    let encrypted_cursor = writer.into_inner().expect("finalize writer");
    let encrypted_ole_bytes = encrypted_cursor.into_inner();

    let mut cfb = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");
    let decrypted = decrypt_agile_ooxml_from_cfb(&mut cfb, password).expect("decrypt");
    assert_eq!(decrypted, plaintext_zip);
}

#[test]
fn decrypt_agile_ooxml_from_cfb_errors_on_missing_streams() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Missing `EncryptionInfo`/`EncryptedPackage` streams should return a
    // structured error so callers can distinguish "not encrypted" from "bad
    // password" etc.
    let err = decrypt_agile_ooxml_from_cfb(&mut ole, "pw").expect_err("expected error");
    assert!(
        matches!(err, OffCryptoError::MissingRequiredStream { ref stream } if stream == "EncryptionInfo"),
        "expected MissingRequiredStream(EncryptionInfo), got {err:?}"
    );
}
