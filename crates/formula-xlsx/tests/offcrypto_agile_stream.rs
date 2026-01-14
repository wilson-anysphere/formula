#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Seek, Write};

use formula_xlsx::offcrypto::decrypt_agile_encrypted_package_stream;
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};
use zip::write::FileOptions;

fn make_zip_bytes(payload_len: usize) -> Vec<u8> {
    let payload: Vec<u8> = (0..payload_len).map(|i| (i % 251) as u8).collect();

    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut cursor);
        let opts = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file("data.bin", opts).expect("start zip entry");
        zip.write_all(&payload).expect("write payload");
        zip.finish().expect("finish zip");
    }
    cursor.into_inner()
}

fn open_stream<R: Read + Seek + Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> cfb::Stream<R> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .expect("open OLE stream")
}

#[test]
fn decrypts_agile_encrypted_package_streaming() {
    let password = "correct horse battery staple";
    // Ensure we span multiple 4096-byte chunks and require truncation (not a multiple of 16).
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt agile encrypted package");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);
}

#[test]
fn decrypts_agile_encrypted_package_streaming_without_data_integrity() {
    let password = "correct horse battery staple";
    let plaintext = make_zip_bytes(12_345);

    let mut rng = StdRng::seed_from_u64(0xD15EA5E_u64);
    let cursor = Cursor::new(Vec::new());
    let mut writer =
        Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create agile writer");
    writer
        .write_all(&plaintext)
        .expect("write plaintext package bytes");
    let cursor = writer.into_inner().expect("finalize agile writer");
    let encrypted_ole_bytes = cursor.into_inner();

    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_ole_bytes)).expect("open cfb");

    let mut encryption_info_stream = open_stream(&mut ole, "EncryptionInfo");
    let mut encryption_info = Vec::new();
    encryption_info_stream
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    // Patch the EncryptionInfo XML to remove `<dataIntegrity .../>`.
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("EncryptionInfo must contain XML");
    let header = encryption_info[..xml_start].to_vec();
    let xml = std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");
    let start = xml
        .find("<dataIntegrity")
        .expect("expected <dataIntegrity> element");
    let end = if let Some(end_rel) = xml[start..].find("/>") {
        start + end_rel + 2
    } else if let Some(end_rel) = xml[start..].find("</dataIntegrity>") {
        start + end_rel + "</dataIntegrity>".len()
    } else {
        panic!("expected </dataIntegrity> or />");
    };
    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..start]);
    patched_xml.push_str(&xml[end..]);
    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    let mut encrypted_package_stream = open_stream(&mut ole, "EncryptedPackage");
    let mut out = Vec::new();
    let declared_len = decrypt_agile_encrypted_package_stream(
        &encryption_info,
        &mut encrypted_package_stream,
        password,
        &mut out,
    )
    .expect("decrypt agile encrypted package without integrity");

    assert_eq!(declared_len as usize, plaintext.len());
    assert_eq!(out, plaintext);
}
