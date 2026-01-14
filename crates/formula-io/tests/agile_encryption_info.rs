use std::io::Read;
use std::path::{Path, PathBuf};

use formula_io::extract_agile_encryption_info_xml;

/// Maximum password spin count allowed for committed fixtures.
///
/// Agile encryption uses the password spin count as the iteration count for key derivation.
/// Extremely large values can make CI decryption tests unnecessarily slow.
const MAX_CI_SPIN_COUNT: u32 = 100_000;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AgileEncryptionParams {
    spin_count: u32,
    cipher_algorithm: String,
    cipher_chaining: String,
    key_bits: u32,
    hash_algorithm: String,
    salt_size: Option<u32>,
}

impl AgileEncryptionParams {
    fn new(
        spin_count: u32,
        cipher_algorithm: &str,
        cipher_chaining: &str,
        key_bits: u32,
        hash_algorithm: &str,
        salt_size: Option<u32>,
    ) -> Self {
        Self {
            spin_count,
            cipher_algorithm: cipher_algorithm.to_string(),
            cipher_chaining: cipher_chaining.to_string(),
            key_bits,
            hash_algorithm: hash_algorithm.to_string(),
            salt_size,
        }
    }
}

fn read_encryption_info_xml(path: &Path) -> String {
    let file = std::fs::File::open(path).expect("open encrypted OOXML fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("open CFB container");
    let mut stream = ole
        .open_stream("EncryptionInfo")
        .or_else(|_| ole.open_stream("/EncryptionInfo"))
        .expect("open EncryptionInfo stream");
    let mut bytes = Vec::new();
    stream
        .read_to_end(&mut bytes)
        .expect("read EncryptionInfo stream bytes");

    extract_agile_encryption_info_xml(&bytes).unwrap_or_else(|err| {
        panic!("failed to extract Agile EncryptionInfo XML from {path:?}: {err}")
    })
}

fn parse_agile_params_from_xml(xml: &str) -> AgileEncryptionParams {
    // `spinCount` only appears on the password keyEncryptor element, so parse all parameters from
    // the same element for consistency.
    let doc = roxmltree::Document::parse(xml).expect("parse extracted `<encryption>` XML");
    let node = doc
        .descendants()
        .find(|node| node.is_element() && node.attribute("spinCount").is_some())
        .unwrap_or_else(|| {
            panic!("expected an element with `spinCount` attribute in extracted `<encryption>` XML")
        });

    let spin_count: u32 = node
        .attribute("spinCount")
        .unwrap()
        .parse()
        .unwrap_or_else(|_| {
            panic!(
                "expected `spinCount` to be an integer, got {:?}",
                node.attribute("spinCount")
            )
        });
    let key_bits: u32 = node
        .attribute("keyBits")
        .unwrap()
        .parse()
        .unwrap_or_else(|_| {
            panic!(
                "expected `keyBits` to be an integer, got {:?}",
                node.attribute("keyBits")
            )
        });

    AgileEncryptionParams::new(
        spin_count,
        node.attribute("cipherAlgorithm").unwrap(),
        node.attribute("cipherChaining").unwrap(),
        key_bits,
        node.attribute("hashAlgorithm").unwrap(),
        node.attribute("saltSize").map(|value| {
            value
                .parse::<u32>()
                .unwrap_or_else(|_| panic!("expected `saltSize` to be an integer, got {value:?}"))
        }),
    )
}

fn parse_agile_params_from_readme(readme: &str, fixture: &str) -> AgileEncryptionParams {
    let needle = format!("| {fixture} |");
    let line = readme
        .lines()
        .find(|line| line.trim_start().starts_with(&needle))
        .unwrap_or_else(|| {
            panic!(
                "missing README row for fixture {fixture:?} (expected a markdown table row starting with `{needle}`)"
            )
        });

    let cols: Vec<&str> = line
        .trim()
        .trim_matches('|')
        .split('|')
        .map(|c| c.trim())
        .collect();
    assert_eq!(
        cols.len(),
        7,
        "expected README table row for {fixture:?} to have 7 columns, got {cols:?}"
    );
    assert_eq!(
        cols[0], fixture,
        "expected first README column to be fixture name"
    );

    let salt_size = match cols[6] {
        "" | "-" => None,
        other => Some(other.parse::<u32>().unwrap_or_else(|_| {
            panic!(
                "expected README saltSize column for {fixture:?} to be an integer, got {other:?}"
            )
        })),
    };

    AgileEncryptionParams::new(
        cols[1].parse::<u32>().unwrap_or_else(|_| {
            panic!(
                "expected README spinCount to be an integer, got {:?}",
                cols[1]
            )
        }),
        cols[2],
        cols[3],
        cols[4].parse::<u32>().unwrap_or_else(|_| {
            panic!(
                "expected README keyBits to be an integer, got {:?}",
                cols[4]
            )
        }),
        cols[5],
        salt_size,
    )
}

#[test]
fn agile_encryption_info_params_match_docs_and_ci_bounds() {
    let readme = fixture_path("encrypted/ooxml/README.md");

    let readme_text =
        std::fs::read_to_string(&readme).expect("read fixtures/encrypted/ooxml/README.md");

    for (fixture_name, expected) in [
        (
            "agile.xlsx",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "agile-large.xlsx",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "agile-unicode.xlsx",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "agile-unicode-excel.xlsx",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "agile-basic.xlsm",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "basic-password.xlsm",
            AgileEncryptionParams::new(100_000, "AES", "ChainingModeCBC", 256, "SHA512", Some(16)),
        ),
        (
            "agile-empty-password.xlsx",
            AgileEncryptionParams::new(1_000, "AES", "ChainingModeCBC", 128, "SHA256", Some(16)),
        ),
    ] {
        let fixture_rel = format!("encrypted/ooxml/{fixture_name}");
        let fixture = fixture_path(&fixture_rel);

        let xml = read_encryption_info_xml(&fixture);
        let actual = parse_agile_params_from_xml(&xml);

        assert!(
            actual.spin_count <= MAX_CI_SPIN_COUNT,
            "{fixture_rel} has spinCount={} which exceeds MAX_CI_SPIN_COUNT={MAX_CI_SPIN_COUNT}. \
             Large spin counts can make CI decryption tests very slow; regenerate the fixture with a \
             smaller spinCount (or update the bound intentionally).",
            actual.spin_count
        );

        assert_eq!(
            actual,
            expected,
            "{fixture_rel} Agile EncryptionInfo parameters drifted.\n\
             Update `{}` and this test's `expected` parameters if the new values are intentional.\n\
             Extracted `<encryption>` XML:\n{xml}\n",
            readme.display()
        );

        let documented = parse_agile_params_from_readme(&readme_text, fixture_name);
        assert_eq!(
            documented,
            expected,
            "`{}` does not match the test's expected parameters for {fixture_rel}. \
             Keep the README in sync with the fixture + test expectations to prevent silent drift.",
            readme.display()
        );
    }
}
