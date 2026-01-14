use std::io::Read;
use std::path::{Path, PathBuf};

use formula_io::extract_agile_encryption_info_xml;

/// Maximum password spin count allowed for committed fixtures.
///
/// Agile encryption uses the password spin count as the iteration count for key derivation.
/// Extremely large values can make CI decryption tests unnecessarily slow.
const MAX_CI_SPIN_COUNT: u32 = 100_000;

const NS_ENCRYPTION: &str = "http://schemas.microsoft.com/office/2006/encryption";

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

fn get_required_attr<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> &'a str {
    node.attribute(name)
        .unwrap_or_else(|| panic!("missing `{name}` attribute on <{}>", node.tag_name().name()))
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

    let spin_count: u32 = get_required_attr(node, "spinCount")
        .parse()
        .unwrap_or_else(|err| panic!("invalid encryptedKey@spinCount: {err}"));
    let key_bits: u32 = get_required_attr(node, "keyBits")
        .parse()
        .unwrap_or_else(|err| panic!("invalid encryptedKey@keyBits: {err}"));

    AgileEncryptionParams::new(
        spin_count,
        get_required_attr(node, "cipherAlgorithm"),
        get_required_attr(node, "cipherChaining"),
        key_bits,
        get_required_attr(node, "hashAlgorithm"),
        node.attribute("saltSize").map(|value| {
            value
                .parse::<u32>()
                .unwrap_or_else(|_| panic!("expected `saltSize` to be an integer, got {value:?}"))
        }),
    )
}

fn parse_agile_params_table_from_readme(readme: &str) -> Vec<(String, AgileEncryptionParams)> {
    const HEADER: &str =
        "| fixture | spinCount | cipherAlgorithm | cipherChaining | keyBits | hashAlgorithm | saltSize |";

    let mut lines = readme.lines().peekable();
    while let Some(line) = lines.next() {
        if line.trim() == HEADER {
            break;
        }
    }
    assert!(
        readme.lines().any(|line| line.trim() == HEADER),
        "failed to locate Agile params table header in README (expected a line equal to `{HEADER}`)"
    );

    // Skip the markdown separator row (`| --- | ... |`).
    let sep = lines
        .next()
        .unwrap_or_else(|| panic!("missing markdown separator row after `{HEADER}`"));
    assert!(
        sep.trim_start().starts_with("| --- |"),
        "expected markdown separator row after `{HEADER}`, got {sep:?}"
    );

    let mut out = Vec::new();
    let mut seen = std::collections::HashSet::new();
    while let Some(&line) = lines.peek() {
        let line = line.trim();
        if !line.starts_with('|') {
            break;
        }
        lines.next();

        // Skip any accidental header re-occurrences.
        if line == HEADER || line.starts_with("| --- |") {
            continue;
        }

        let cols: Vec<&str> = line
            .trim_matches('|')
            .split('|')
            .map(|c| c.trim())
            .collect();
        assert_eq!(
            cols.len(),
            7,
            "expected README Agile params table row to have 7 columns, got {cols:?}"
        );

        let fixture = cols[0].to_string();
        assert!(
            !fixture.is_empty(),
            "expected README Agile params table row fixture name to be non-empty: {line:?}"
        );
        assert!(
            seen.insert(fixture.clone()),
            "duplicate fixture row in README Agile params table: {fixture:?}"
        );

        let salt_size = match cols[6] {
            "" | "-" => None,
            other => Some(other.parse::<u32>().unwrap_or_else(|_| {
                panic!(
                    "expected README saltSize column for {fixture:?} to be an integer, got {other:?}"
                )
            })),
        };

        let params = AgileEncryptionParams::new(
            cols[1].parse::<u32>().unwrap_or_else(|_| {
                panic!(
                    "expected README spinCount for {fixture:?} to be an integer, got {:?}",
                    cols[1]
                )
            }),
            cols[2],
            cols[3],
            cols[4].parse::<u32>().unwrap_or_else(|_| {
                panic!(
                    "expected README keyBits for {fixture:?} to be an integer, got {:?}",
                    cols[4]
                )
            }),
            cols[5],
            salt_size,
        );

        out.push((fixture, params));
    }

    assert!(
        !out.is_empty(),
        "README Agile params table had no data rows after `{HEADER}`"
    );
    out
}

fn assert_agile_key_data_params_match_expected(
    xml: &str,
    expected: &AgileEncryptionParams,
    fixture_rel: &str,
) {
    const EXPECTED_BLOCK_SIZE: u32 = 16;

    let doc = roxmltree::Document::parse(xml)
        .unwrap_or_else(|err| panic!("parse Agile EncryptionInfo XML for {fixture_rel}: {err}"));

    let key_data = doc
        .descendants()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "keyData"
                && node.tag_name().namespace() == Some(NS_ENCRYPTION)
        })
        .unwrap_or_else(|| panic!("missing <keyData> element in {fixture_rel}"));

    assert_eq!(
        get_required_attr(key_data, "cipherAlgorithm"),
        expected.cipher_algorithm.as_str(),
        "{fixture_rel}: unexpected keyData@cipherAlgorithm"
    );
    assert_eq!(
        get_required_attr(key_data, "cipherChaining"),
        expected.cipher_chaining.as_str(),
        "{fixture_rel}: unexpected keyData@cipherChaining"
    );

    let key_bits: u32 = get_required_attr(key_data, "keyBits")
        .parse()
        .unwrap_or_else(|err| panic!("{fixture_rel}: invalid keyData@keyBits: {err}"));
    assert_eq!(key_bits, expected.key_bits, "{fixture_rel}: unexpected keyData@keyBits");

    assert_eq!(
        get_required_attr(key_data, "hashAlgorithm"),
        expected.hash_algorithm.as_str(),
        "{fixture_rel}: unexpected keyData@hashAlgorithm"
    );

    let block_size: u32 = get_required_attr(key_data, "blockSize")
        .parse()
        .unwrap_or_else(|err| panic!("{fixture_rel}: invalid keyData@blockSize: {err}"));
    assert_eq!(
        block_size, EXPECTED_BLOCK_SIZE,
        "{fixture_rel}: unexpected keyData@blockSize"
    );
}

#[test]
fn agile_encryption_info_params_match_docs_and_ci_bounds() {
    let readme = fixture_path("encrypted/ooxml/README.md");

    let readme_text =
        std::fs::read_to_string(&readme).expect("read fixtures/encrypted/ooxml/README.md");

    let expected_by_fixture = parse_agile_params_table_from_readme(&readme_text);

    for required in [
        "agile.xlsx",
        "agile-large.xlsx",
        "agile-empty-password.xlsx",
        "agile-unicode.xlsx",
        "agile-unicode-excel.xlsx",
        "agile-basic.xlsm",
    ] {
        assert!(
            expected_by_fixture.iter().any(|(fixture, _)| fixture == required),
            "README Agile params table is missing required fixture row for {required:?}"
        );
    }

    for (fixture_name, expected) in expected_by_fixture {
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
             Update `{}` (and regenerate the fixture) if the new values are intentional.\n\
             Extracted `<encryption>` XML:\n{xml}\n",
            readme.display()
        );

        assert_agile_key_data_params_match_expected(&xml, &expected, &fixture_rel);
    }
}
