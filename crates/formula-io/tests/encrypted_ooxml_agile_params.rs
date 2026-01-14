use std::io::Read as _;
use std::path::{Path, PathBuf};

/// Encrypted OOXML fixtures live at `fixtures/encrypted/ooxml/`.
fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn open_stream_case_tolerant<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<cfb::Stream<R>> {
    ole.open_stream(name)
        .or_else(|_| ole.open_stream(format!("/{name}")))
}

fn read_encryption_info_stream(path: &Path) -> Vec<u8> {
    let file = std::fs::File::open(path).expect("open fixture file");
    let mut ole = cfb::CompoundFile::open(file).expect("open cfb (OLE) container");
    let mut stream =
        open_stream_case_tolerant(&mut ole, "EncryptionInfo").expect("open EncryptionInfo stream");
    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .expect("read EncryptionInfo stream bytes");
    buf
}

#[derive(Debug, Clone, Copy)]
struct ExpectedAgileParams {
    spin_count: u32,
    hash_algorithm: &'static str,
    cipher_algorithm: &'static str,
    cipher_chaining: &'static str,
    key_bits: u32,
    block_size: u32,
}

fn get_required_attr<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> &'a str {
    node.attribute(name)
        .unwrap_or_else(|| panic!("missing `{}` attribute on <{}>", name, node.tag_name().name()))
}

fn assert_agile_fixture_params(fixture: &str, expected: ExpectedAgileParams) {
    const NS_ENCRYPTION: &str = "http://schemas.microsoft.com/office/2006/encryption";
    const NS_PASSWORD: &str = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

    let path = fixture_path(fixture);
    let encryption_info = read_encryption_info_stream(&path);

    let xml = formula_io::extract_agile_encryption_info_xml(&encryption_info)
        .unwrap_or_else(|err| panic!("extract Agile EncryptionInfo XML for {fixture}: {err}"));
    let doc = roxmltree::Document::parse(&xml)
        .unwrap_or_else(|err| panic!("parse Agile EncryptionInfo XML for {fixture}: {err}"));

    let key_data = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyData"
                && n.tag_name().namespace() == Some(NS_ENCRYPTION)
        })
        .unwrap_or_else(|| panic!("missing <keyData> element in {fixture}"));

    let encrypted_key = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "encryptedKey"
                && n.tag_name().namespace() == Some(NS_PASSWORD)
        })
        .unwrap_or_else(|| panic!("missing <p:encryptedKey> element in {fixture}"));

    let key_data_hash_algorithm = get_required_attr(key_data, "hashAlgorithm");
    let encrypted_key_hash_algorithm = get_required_attr(encrypted_key, "hashAlgorithm");

    assert_eq!(
        key_data_hash_algorithm, expected.hash_algorithm,
        "{fixture}: unexpected keyData@hashAlgorithm"
    );
    assert_eq!(
        encrypted_key_hash_algorithm, expected.hash_algorithm,
        "{fixture}: unexpected encryptedKey@hashAlgorithm"
    );
    assert_eq!(
        key_data_hash_algorithm, encrypted_key_hash_algorithm,
        "{fixture}: expected keyData@hashAlgorithm to match encryptedKey@hashAlgorithm"
    );

    let spin_count: u32 = get_required_attr(encrypted_key, "spinCount")
        .parse()
        .unwrap_or_else(|err| panic!("{fixture}: invalid encryptedKey@spinCount: {err}"));
    assert_eq!(
        spin_count, expected.spin_count,
        "{fixture}: unexpected encryptedKey@spinCount"
    );

    assert_eq!(
        get_required_attr(key_data, "cipherAlgorithm"),
        expected.cipher_algorithm,
        "{fixture}: unexpected keyData@cipherAlgorithm"
    );
    assert_eq!(
        get_required_attr(key_data, "cipherChaining"),
        expected.cipher_chaining,
        "{fixture}: unexpected keyData@cipherChaining"
    );

    let key_bits: u32 = get_required_attr(key_data, "keyBits")
        .parse()
        .unwrap_or_else(|err| panic!("{fixture}: invalid keyData@keyBits: {err}"));
    assert_eq!(key_bits, expected.key_bits, "{fixture}: unexpected keyData@keyBits");

    let block_size: u32 = get_required_attr(key_data, "blockSize")
        .parse()
        .unwrap_or_else(|err| panic!("{fixture}: invalid keyData@blockSize: {err}"));
    assert_eq!(
        block_size, expected.block_size,
        "{fixture}: unexpected keyData@blockSize"
    );
}

#[test]
fn agile_ooxml_fixtures_pin_expected_encryption_info_parameters() {
    // These assertions intentionally pin the Agile `EncryptionInfo` parameter choices for fixtures.
    // If the fixtures are regenerated, we want CI to fail loudly if algorithms/spinCount drift.
    //
    // In particular, a silent bump in `spinCount` can drastically slow down password verification
    // (iterated hashing).
    let cases = [
        (
            "agile.xlsx",
            ExpectedAgileParams {
                spin_count: 100_000,
                hash_algorithm: "SHA512",
                cipher_algorithm: "AES",
                cipher_chaining: "ChainingModeCBC",
                key_bits: 256,
                block_size: 16,
            },
        ),
        (
            "agile-large.xlsx",
            ExpectedAgileParams {
                spin_count: 100_000,
                hash_algorithm: "SHA512",
                cipher_algorithm: "AES",
                cipher_chaining: "ChainingModeCBC",
                key_bits: 256,
                block_size: 16,
            },
        ),
        (
            "agile-empty-password.xlsx",
            ExpectedAgileParams {
                // This fixture uses a reduced spinCount + SHA-256 + AES-128 to keep the empty-password
                // decrypt tests fast.
                spin_count: 1_000,
                hash_algorithm: "SHA256",
                cipher_algorithm: "AES",
                cipher_chaining: "ChainingModeCBC",
                key_bits: 128,
                block_size: 16,
            },
        ),
        (
            "agile-unicode.xlsx",
            ExpectedAgileParams {
                spin_count: 100_000,
                hash_algorithm: "SHA512",
                cipher_algorithm: "AES",
                cipher_chaining: "ChainingModeCBC",
                key_bits: 256,
                block_size: 16,
            },
        ),
        (
            "agile-unicode-excel.xlsx",
            ExpectedAgileParams {
                spin_count: 100_000,
                hash_algorithm: "SHA512",
                cipher_algorithm: "AES",
                cipher_chaining: "ChainingModeCBC",
                key_bits: 256,
                block_size: 16,
            },
        ),
    ];

    for (fixture, expected) in cases {
        assert_agile_fixture_params(fixture, expected);
    }
}
