#![cfg(all(feature = "vba", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write};

use formula_xlsx::vba::{VbaCertificateTrust, VbaSignatureTrustOptions, VbaSignatureVerification};
use formula_xlsx::XlsxPackage;
use openssl::x509::X509;
use zip::write::FileOptions;

const TEST_KEY_PEM: &str = r#"-----BEGIN PRIVATE KEY-----
MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC8kN1a0raWt6a7
MzszVTIVgdZHbie+mkVWDoMrgTQYX8tm/3yqTLQMXWhuV0hZtrUydWlsRB8k0aTS
aXFCzmmNgAqFh13uQ/rFW82zh5UCWXuaX43uc5JWebD4TzkN2b4vye3s/S3QCmZK
5kT6jWPDaRyngOvaHgcBB9meMS6QT9Efb2SdV/a6QkrGm0nhMfJyZEY00FKEhxJf
A4JlVDVhmmQdpCoXb++cqK/xo9DehmrivP1CL/dFPjy3wkbtHpb+uAatzBNtaqmE
bYwtaw0rqxlkbKZT6baayf9klTXFah4bEzRDSJQrzM6HjhNYDiCBM9omNSowkyb9
PVJqkRRvAgMBAAECggEAAalfIflAXaShpf2mFGY4SkM6IagBVqciXEdFdaEuVdam
QrKWpSOG5KMAFBTV0OCQyTCKrMcO5TKpuqbuNhH+kR4jOZj/RWW49HtCHUZhFEO4
mJwl8od3LybkXdPI42vbRq2HWLLEcBRfYWKVEgYj7mljNMCok2P3WNV6X+/8Ao6F
n9+NDnE/K4e3xr/7pS4hldm/b67KZh92Rhzfoezdpk+uVXtU6ccTeyO10YCnng2w
Qhls2Hkrx92GspAp8gdK5Hnrk2y/Lx8EmThUUSSP7h2uxvoUs3RNevZQYp2vv6fc
0ffD1M9fI9iz79UKkad+1VGwjO2SPPK28LZWUNgO1QKBgQDjX/4W7ZlbeYNN1sxa
pbdEB0eNFs5jk3B4JH986h0lTPZwdkcaEDwJ9sp/pxceYPFZ8ul/IGg9nlDNIa88
BccrH/o1gZjB68UdM2fu6jWYDC9dzscUYjFkFDndQoH8ACg1Nt27UZi5TKN1DTGM
dnPf4Tb7VDOsVpPiCxZDCsgFbQKBgQDUThqr43mgAaB8nlmCVRFIP4Wn3BQScXt/
J5xLKsI5AadvCbNUPw/gO518qhcsFNSKRUzBx1d10AR1h/NX820td/swjHkWm31V
PjbNl+5G8RwHib2miqdI4KRe5RdGeWfSW9wEU1epkeCGPBbApNfgEec0/PN0T+7q
xNBiaSaDywKBgGPMNT0hCkexHOWkWsuKota0Dz6o/OuNwjapZl+Qbjx5/Ey+TVTu
PTvuW1EOKMKHsEdXrA7FTZuGClcO6tgAfTu7bFnhyQeMkVbQwlSF7gIPjxawdIbI
1n7jtcYcs+rEsuEwdMAL/2mNbs0ofk/1icSBGF3VxlxlH8F+NkY0zDg9AoGAOQi6
dY7or6mAObo4haDgwa3+8/dVlRbTfHdhr3fPMY1WM6hBetJuK2kYh9MR4o++AV9Y
nX416rp1WDWrk+cbX2mqG4LBTOd8phfOlTDJnFlNlGDWiBUbl6JxxeR5ej9HOuXe
l3LkS/Oag7VEz3/5VoK4wC1sIcUPhBZXfPiOlj0CgYAVBqxAtjYV+Of4nzYXlvyD
nKgzkiBZCPvjLuINLxl02hMkl5L1rkYYFlBonRXkBZi/qi/sy5yWJFD4bNdXADjx
l6I38mljR1b525IXYYgxl70AE5/oiURtl3rzv4gzYvm7lhV7/c7ZTwY0X43vTO7d
0TiTGpZ2jyGWBsNrW2X+Rw==
-----END PRIVATE KEY-----"#;

const TEST_CERT_PEM: &str = r#"-----BEGIN CERTIFICATE-----
MIIDFzCCAf+gAwIBAgIUQZEa3yk9CWWcytfnuDxC4+5iaPUwDQYJKoZIhvcNAQEL
BQAwGzEZMBcGA1UEAwwQRm9ybXVsYSBWQkEgVGVzdDAeFw0yNjAxMTExMDM2NDBa
Fw0zNjAxMDkxMDM2NDBaMBsxGTAXBgNVBAMMEEZvcm11bGEgVkJBIFRlc3QwggEi
MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC8kN1a0raWt6a7MzszVTIVgdZH
bie+mkVWDoMrgTQYX8tm/3yqTLQMXWhuV0hZtrUydWlsRB8k0aTSaXFCzmmNgAqF
h13uQ/rFW82zh5UCWXuaX43uc5JWebD4TzkN2b4vye3s/S3QCmZK5kT6jWPDaRyn
gOvaHgcBB9meMS6QT9Efb2SdV/a6QkrGm0nhMfJyZEY00FKEhxJfA4JlVDVhmmQd
pCoXb++cqK/xo9DehmrivP1CL/dFPjy3wkbtHpb+uAatzBNtaqmEbYwtaw0rqxlk
bKZT6baayf9klTXFah4bEzRDSJQrzM6HjhNYDiCBM9omNSowkyb9PVJqkRRvAgMB
AAGjUzBRMB0GA1UdDgQWBBSyceRXYQd4wvXncCr1AcYneVlpWTAfBgNVHSMEGDAW
gBSyceRXYQd4wvXncCr1AcYneVlpWTAPBgNVHRMBAf8EBTADAQH/MA0GCSqGSIb3
DQEBCwUAA4IBAQBbcQVLwUMdKA5xj2woUkEe9kcTtS9YOMeCoBE48Fw8KfgkbKtK
lte7yIBdgHdjjAke88g9Dh64OlcRQigu0fS025bXcw1g7AKc0fkBDro8j8GHqdi6
APR5O9xnfdslBSX1cDN/530Q+vRpha/LxLfSG2UXovmb163110RD6ina9gTIvy9r
plrbDIYpuR+SiI0uaQtcwCdbXPtHLlEUUp0ZbnW3i+RHmt9DnwQM1B/hAv9zdg9m
ls5Xirz7pTI39gHpSd86SfJWBbPPcJHabdmgRTJW8AbxMjS2xBDU3pxzGw52MgfK
Kj4ozoiZRiNvvWvqUGOt1yKu7S7nbEPuW3rX
-----END CERTIFICATE-----"#;

fn make_pkcs7_signed_message(data: &[u8]) -> Vec<u8> {
    use openssl::pkcs7::{Pkcs7, Pkcs7Flags};
    use openssl::pkey::PKey;
    use openssl::stack::Stack;
    use openssl::x509::X509;

    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");
    let cert = X509::from_pem(TEST_CERT_PEM.as_bytes()).expect("parse certificate");
    let extra_certs = Stack::new().expect("create cert stack");

    let pkcs7 = Pkcs7::sign(
        &cert,
        &pkey,
        &extra_certs,
        data,
        // NOATTR keeps the output deterministic (avoids adding a SigningTime attribute).
        Pkcs7Flags::BINARY | Pkcs7Flags::NOATTR,
    )
    .expect("pkcs7 sign");
    pkcs7.to_der().expect("pkcs7 DER")
}

fn build_ole_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create compound file");

    for (path, bytes) in streams {
        let mut stream = ole.create_stream(path).expect("create stream");
        stream.write_all(bytes).expect("write stream bytes");
    }

    ole.into_inner().into_inner()
}

fn build_empty_ole() -> Vec<u8> {
    build_ole_with_streams(&[])
}

fn build_vba_project_bin_with_signature(signature_blob: &[u8]) -> Vec<u8> {
    build_ole_with_streams(&[("\u{0005}DigitalSignature", signature_blob)])
}

fn build_zip(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip file");
    }

    zip.finish().expect("finish zip").into_inner()
}

fn make_unrelated_root_cert_der() -> Vec<u8> {
    use openssl::asn1::Asn1Time;
    use openssl::bn::BigNum;
    use openssl::hash::MessageDigest;
    use openssl::pkey::PKey;
    use openssl::x509::{X509Builder, X509NameBuilder};

    let pkey = PKey::private_key_from_pem(TEST_KEY_PEM.as_bytes()).expect("parse private key");

    let mut name_builder = X509NameBuilder::new().expect("x509 name builder");
    name_builder
        .append_entry_by_text("CN", "Formula VBA Unrelated Root")
        .expect("CN");
    let name = name_builder.build();

    let mut builder = X509Builder::new().expect("x509 builder");
    builder.set_version(2).expect("set version");

    let serial = BigNum::from_u32(2)
        .expect("serial bn")
        .to_asn1_integer()
        .expect("serial integer");
    builder.set_serial_number(&serial).expect("serial");

    builder.set_subject_name(&name).expect("subject name");
    builder.set_issuer_name(&name).expect("issuer name");
    builder.set_pubkey(&pkey).expect("pubkey");

    builder
        .set_not_before(&Asn1Time::days_from_now(0).expect("not before"))
        .expect("set not before");
    builder
        .set_not_after(&Asn1Time::days_from_now(3650).expect("not after"))
        .expect("set not after");

    builder
        .sign(&pkey, MessageDigest::sha256())
        .expect("sign");
    builder.build().to_der().expect("DER")
}

#[test]
fn embedded_signature_stream_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let vba_project_bin = build_vba_project_bin_with_signature(&pkcs7);
    let zip_bytes = build_zip(&[("xl/vbaProject.bin", &vba_project_bin)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn embedded_signature_stream_trust_is_reported_with_leading_slash_part_names() {
    // Some producers incorrectly store OPC part names with a leading `/` in the ZIP.
    // Ensure signature verification and trust evaluation still work.
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let vba_project_bin = build_vba_project_bin_with_signature(&pkcs7);
    let zip_bytes = build_zip(&[("/xl/vbaProject.bin", &vba_project_bin)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");

    // `xl/vbaProject.bin` must be a valid OLE file (even if unsigned) so the fallback embedded
    // signature scan can run without errors.
    let vba_project_bin = build_empty_ole();

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &pkcs7),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_trust_is_reported_with_leading_slash_part_names() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");

    // `xl/vbaProject.bin` must be a valid OLE file (even if unsigned) so the fallback embedded
    // signature scan can run without errors.
    let vba_project_bin = build_empty_ole();

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("/xl/vbaProject.bin", &vba_project_bin),
        ("/xl/_rels/vbaProject.bin.rels", vba_rels),
        ("/xl/vbaProjectSignature.bin", &pkcs7),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn external_signature_part_ole_container_trust_is_reported() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let signature_part_ole = build_vba_project_bin_with_signature(&pkcs7);

    // Use a minimal (unsigned) OLE container for `xl/vbaProject.bin` so binding verification can
    // run if needed.
    let vba_project_bin = build_empty_ole();

    let vba_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

    let zip_bytes = build_zip(&[
        ("xl/vbaProject.bin", &vba_project_bin),
        ("xl/_rels/vbaProject.bin.rels", vba_rels),
        ("xl/vbaProjectSignature.bin", &signature_part_ole),
    ]);
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let signer_der = X509::from_pem(TEST_CERT_PEM.as_bytes())
        .expect("parse test certificate")
        .to_der()
        .expect("convert test certificate to DER");
    let wrong_root_der = make_unrelated_root_cert_der();

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Unknown);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![signer_der.clone()],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Trusted);

    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions {
            trusted_root_certs_der: vec![wrong_root_der],
        })
        .expect("verify signature")
        .expect("signature should be present");
    assert_eq!(sig.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.cert_trust, VbaCertificateTrust::Untrusted);
}

#[test]
fn returns_none_when_vba_project_bin_is_missing() {
    let pkcs7 = make_pkcs7_signed_message(b"formula-xlsx-trust-test");
    let zip_bytes = build_zip(&[("xl/vbaProjectSignature.bin", &pkcs7)]);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");
    let sig = pkg
        .verify_vba_digital_signature_with_trust(&VbaSignatureTrustOptions::default())
        .expect("verify signature");
    assert!(sig.is_none());
}
