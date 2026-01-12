#![cfg(feature = "vba")]

use std::io::{Cursor, Write};

use base64::{engine::general_purpose, Engine as _};
use formula_vba::VbaSignatureVerification;
use formula_xlsx::XlsxPackage;

mod vba_signature_test_utils;
use vba_signature_test_utils::build_vba_signature_ole;

const TEST_PKCS7_DER_B64: &str = concat!(
    "MIIExQYJKoZIhvcNAQcCoIIEtjCCBLICAQExDTALBglghkgBZQMEAgEwHwYJKoZIhvcNAQcBoBIEEGZv",
    "cm11bGEtdmJhLXRlc3SgggMbMIIDFzCCAf+gAwIBAgIUQZEa3yk9CWWcytfnuDxC4+5iaPUwDQYJKoZI",
    "hvcNAQELBQAwGzEZMBcGA1UEAwwQRm9ybXVsYSBWQkEgVGVzdDAeFw0yNjAxMTExMDM2NDBaFw0zNjAx",
    "MDkxMDM2NDBaMBsxGTAXBgNVBAMMEEZvcm11bGEgVkJBIFRlc3QwggEiMA0GCSqGSIb3DQEBAQUAA4IB",
    "DwAwggEKAoIBAQC8kN1a0raWt6a7MzszVTIVgdZHbie+mkVWDoMrgTQYX8tm/3yqTLQMXWhuV0hZtrUy",
    "dWlsRB8k0aTSaXFCzmmNgAqFh13uQ/rFW82zh5UCWXuaX43uc5JWebD4TzkN2b4vye3s/S3QCmZK5kT6",
    "jWPDaRyngOvaHgcBB9meMS6QT9Efb2SdV/a6QkrGm0nhMfJyZEY00FKEhxJfA4JlVDVhmmQdpCoXb++c",
    "qK/xo9DehmrivP1CL/dFPjy3wkbtHpb+uAatzBNtaqmEbYwtaw0rqxlkbKZT6baayf9klTXFah4bEzRD",
    "SJQrzM6HjhNYDiCBM9omNSowkyb9PVJqkRRvAgMBAAGjUzBRMB0GA1UdDgQWBBSyceRXYQd4wvXncCr1",
    "AcYneVlpWTAfBgNVHSMEGDAWgBSyceRXYQd4wvXncCr1AcYneVlpWTAPBgNVHRMBAf8EBTADAQH/MA0G",
    "CSqGSIb3DQEBCwUAA4IBAQBbcQVLwUMdKA5xj2woUkEe9kcTtS9YOMeCoBE48Fw8KfgkbKtKlte7yIBd",
    "gHdjjAke88g9Dh64OlcRQigu0fS025bXcw1g7AKc0fkBDro8j8GHqdi6APR5O9xnfdslBSX1cDN/530Q",
    "+vRpha/LxLfSG2UXovmb163110RD6ina9gTIvy9rplrbDIYpuR+SiI0uaQtcwCdbXPtHLlEUUp0ZbnW3",
    "i+RHmt9DnwQM1B/hAv9zdg9mls5Xirz7pTI39gHpSd86SfJWBbPPcJHabdmgRTJW8AbxMjS2xBDU3pxz",
    "Gw52MgfKKj4ozoiZRiNvvWvqUGOt1yKu7S7nbEPuW3rXMYIBXDCCAVgCAQEwMzAbMRkwFwYDVQQDDBBG",
    "b3JtdWxhIFZCQSBUZXN0AhRBkRrfKT0JZZzK1+e4PELj7mJo9TALBglghkgBZQMEAgEwDQYJKoZIhvcN",
    "AQEBBQAEggEAb4deZBs4wNKlhuzW4wZ9Ptljd+M4BgwM47rU3NZhrZgKG1+NDsSKkGyITZyAEx5PwWqn",
    "wGKOCzrqdipAK4BNkB6oGbhD61oioe7/9O5jFWyvYh4NqLkFmAfiHVS6ibmTaS5OnNSoo8BLncYLv+pR",
    "R1alojaQstnZj0LGG5qSwJjhtfIbkHgdVLx0BVb3fGeSb4xjPbEdAkwGfkdpzZTR95GSZF6c4mabyq9S",
    "YHUxTdVu92MSJWBzPmdR6M2/isqmgSqun0vE1kR/IbARZbtB6OsSzxE3rziwlHxoelDRsfyPnmi8TsNt",
    "hH5fWBntXXgwtszAsTVMK92tz4Fz0Q19pg==",
);

fn build_zip_with_vba_project_bin(vba_project_bin: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/vbaProject.bin", options)
        .expect("start vbaProject.bin");
    zip.write_all(vba_project_bin)
        .expect("write vbaProject.bin");

    zip.finish().expect("finish zip").into_inner()
}

fn build_zip_without_vba_project_bin() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options)
        .expect("start workbook.xml");
    zip.write_all(b"<workbook/>").expect("write workbook");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn verify_vba_digital_signature_reports_verified_on_native_targets() {
    let pkcs7_der = general_purpose::STANDARD
        .decode(TEST_PKCS7_DER_B64)
        .expect("base64 decode pkcs7");
    let vba_project_bin = build_vba_signature_ole(&pkcs7_der);
    let zip_bytes = build_zip_with_vba_project_bin(&vba_project_bin);

    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");
    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature inspection should succeed")
        .expect("expected signature to be present");

    #[cfg(not(target_arch = "wasm32"))]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);

    #[cfg(target_arch = "wasm32")]
    assert_eq!(sig.verification, VbaSignatureVerification::SignedButUnverified);
}

#[test]
fn verify_vba_digital_signature_returns_none_without_vba_project_bin() {
    let zip_bytes = build_zip_without_vba_project_bin();
    let pkg = XlsxPackage::from_bytes(&zip_bytes).expect("read package");

    let sig = pkg
        .verify_vba_digital_signature()
        .expect("signature inspection should succeed");
    assert!(sig.is_none(), "expected Ok(None) when vbaProject.bin is absent");
}
