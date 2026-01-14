//! XLSX/XLSB round-trip diff tooling.
//!
//! This crate intentionally operates at the ZIP/Open Packaging Convention layer:
//! it compares workbook "parts" (files within the archive) rather than the ZIP
//! container bytes. This avoids false positives from differing compression or
//! timestamp metadata while still catching fidelity regressions.

pub mod cli;
mod part_kind;
mod presets;
mod rels;
mod xml;

use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::fs::File;
use std::io::Cursor;
use std::io::{Read, Seek, Write};
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use formula_office_crypto::{decrypt_encrypted_package_ole, is_encrypted_ooxml_ole};
use globset::{Glob, GlobSet, GlobSetBuilder};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub use part_kind::{classify_part, PartKind};
pub use presets::IgnorePreset;
pub use xml::{diff_xml, NormalizedXml};

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

#[cfg(test)]
static FAST_PATH_SKIPPED_PARTS: std::sync::atomic::AtomicUsize =
    std::sync::atomic::AtomicUsize::new(0);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum Severity {
    Critical,
    Warning,
    Info,
}

impl Severity {
    pub fn as_str(self) -> &'static str {
        match self {
            Severity::Critical => "CRITICAL",
            Severity::Warning => "WARN",
            Severity::Info => "INFO",
        }
    }
}

impl fmt::Display for Severity {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

#[derive(Debug, Clone)]
pub struct Difference {
    pub severity: Severity,
    pub part: String,
    pub path: String,
    pub kind: String,
    pub expected: Option<String>,
    pub actual: Option<String>,
}

impl Difference {
    pub fn new(
        severity: Severity,
        part: impl Into<String>,
        path: impl Into<String>,
        kind: impl Into<String>,
        expected: Option<String>,
        actual: Option<String>,
    ) -> Self {
        Self {
            severity,
            part: part.into(),
            path: path.into(),
            kind: kind.into(),
            expected,
            actual,
        }
    }
}

impl fmt::Display for Difference {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        writeln!(
            f,
            "[{}] {}{}{}",
            self.severity,
            self.part,
            if self.path.is_empty() { "" } else { ":" },
            self.path
        )?;
        writeln!(f, "  kind: {}", self.kind)?;
        if let Some(expected) = &self.expected {
            writeln!(f, "  expected: {}", expected)?;
        }
        if let Some(actual) = &self.actual {
            writeln!(f, "  actual:   {}", actual)?;
        }
        Ok(())
    }
}

#[derive(Debug, Default)]
pub struct DiffReport {
    pub differences: Vec<Difference>,
}

impl DiffReport {
    pub fn is_empty(&self) -> bool {
        self.differences.is_empty()
    }

    pub fn count(&self, severity: Severity) -> usize {
        self.differences
            .iter()
            .filter(|d| d.severity == severity)
            .count()
    }

    pub fn has_at_least(&self, threshold: Severity) -> bool {
        self.differences.iter().any(|d| d.severity <= threshold)
    }
}

pub struct WorkbookArchive {
    parts: BTreeMap<String, Vec<u8>>,
}

impl WorkbookArchive {
    pub fn open(path: &Path) -> Result<Self> {
        Self::open_with_password(path, None)
    }

    pub fn open_with_password(path: &Path, password: Option<&str>) -> Result<Self> {
        let mut file =
            File::open(path).with_context(|| format!("open workbook {}", path.display()))?;

        let mut header = [0u8; 8];
        let n = file.read(&mut header).context("read workbook header")?;
        file.rewind().context("rewind workbook stream")?;

        if n >= OLE_MAGIC.len() && header == OLE_MAGIC {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)
                .with_context(|| format!("read workbook {}", path.display()))?;

            if !is_encrypted_ooxml_ole(&bytes) {
                return Err(anyhow!(
                    "workbook {} is an OLE compound file but not an Office EncryptedPackage container",
                    path.display()
                ));
            }

            let password = password.ok_or_else(|| {
                anyhow!(
                    "workbook {} is encrypted; a password is required to decrypt it",
                    path.display()
                )
            })?;

            let decrypted = decrypt_encrypted_package_ole(&bytes, password)
                .map_err(|err| anyhow!("failed to decrypt workbook {}: {err}", path.display()))?;

            return WorkbookArchive::from_bytes(&decrypted);
        }

        let mut zip = ZipArchive::new(file).context("parse zip archive")?;
        Self::read_zip(&mut zip)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut zip = ZipArchive::new(cursor).context("parse zip archive")?;
        Self::read_zip(&mut zip)
    }

    fn read_zip<R: Read + std::io::Seek>(zip: &mut ZipArchive<R>) -> Result<Self> {
        let mut parts = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i).context("read zip entry")?;
            if file.is_dir() {
                continue;
            }
            let name = normalize_opc_part_name(file.name());
            // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
            // advertise enormous uncompressed sizes (zip-bomb style OOM).
            let mut buf = Vec::new();
            file.read_to_end(&mut buf)
                .with_context(|| format!("read part {name}"))?;
            if parts.insert(name.clone(), buf).is_some() {
                return Err(anyhow!(
                    "duplicate part name after normalization (possible invalid zip): {name}"
                ));
            }
        }

        Ok(Self { parts })
    }

    pub fn parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.parts.iter().map(|(k, v)| (k.as_str(), v.as_slice()))
    }

    pub fn part_names(&self) -> BTreeSet<&str> {
        self.parts.keys().map(|k| k.as_str()).collect()
    }

    pub fn get(&self, name: &str) -> Option<&[u8]> {
        self.parts.get(name).map(|v| v.as_slice())
    }
}

#[derive(Debug, Clone, Default)]
pub struct DiffOptions {
    /// Exact part names to ignore.
    ///
    /// These are matched against normalized OPC part names (forward slashes, no
    /// leading `/`, and `..` segments resolved). For convenience, callers may
    /// pass Windows-style separators (`\`) or leading slashes and they will be
    /// normalized before matching.
    pub ignore_parts: BTreeSet<String>,
    /// Glob patterns to ignore (see `globset` syntax).
    ///
    /// Patterns are matched against normalized OPC part names. Note that the
    /// library intentionally ignores invalid glob patterns; the `xlsx_diff`
    /// CLI validates globs up-front and will return an error instead.
    pub ignore_globs: Vec<String>,
    /// Fine-grained XML diff ignore rules.
    ///
    /// This allows callers to suppress known-noisy diffs (for example volatile
    /// attributes like `xr:uid` or `x14ac:dyDescent`) without ignoring an entire
    /// OPC part.
    ///
    /// Rules are evaluated against the `(part, XmlDiff.path, XmlDiff.kind)`
    /// tuple. When `ignore_paths` is empty, existing behavior is unchanged.
    pub ignore_paths: Vec<IgnorePathRule>,
    /// When enabled, calcChain-related diffs are treated as CRITICAL instead of being
    /// downgraded to WARNING.
    ///
    /// By default, the diff tool intentionally downgrades calcChain-related churn
    /// (missing `xl/calcChain.{xml,bin}` plus the corresponding plumbing diffs in
    /// `[Content_Types].xml` and `xl/_rels/workbook.*.rels`) to WARNING. This supports
    /// patch workflows where Excel invalidates the calculation chain as a side effect.
    ///
    /// For strict round-trip preservation scoring, callers can set this to `true` to
    /// keep calcChain diffs CRITICAL.
    pub strict_calc_chain: bool,
}

impl DiffOptions {
    /// Apply a built-in ignore preset (opt-in).
    pub fn apply_preset(&mut self, preset: IgnorePreset) {
        for rule in preset.owned_rules() {
            if !self.ignore_paths.contains(&rule) {
                self.ignore_paths.push(rule);
            }
        }
    }
}

/// Parameters for a single diff input (workbook path + optional password).
#[derive(Debug, Clone, Copy)]
pub struct DiffInput<'a> {
    pub path: &'a Path,
    pub password: Option<&'a str>,
}

impl<'a> DiffInput<'a> {
    pub fn new(path: &'a Path) -> Self {
        Self {
            path,
            password: None,
        }
    }
}

/// Fine-grained ignore rule for XML diffs (`XmlDiff`).
///
/// A rule matches (and thus suppresses) an XML diff when:
/// - `part` matches the normalized OPC part name (if provided), and
/// - `path_substring` is contained in the XML diff path, and
/// - `kind` matches the XML diff kind (if provided).
///
/// Part matching supports either exact matches or basic glob matching. If the
/// provided part pattern contains `*` or `?`, it is interpreted as a glob;
/// otherwise it is matched exactly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IgnorePathRule {
    /// Optional part selector (exact or glob).
    pub part: Option<String>,
    /// Substring to match against `XmlDiff.path`.
    pub path_substring: String,
    /// Optional `XmlDiff.kind` match (exact).
    pub kind: Option<String>,
}

pub fn diff_workbooks(expected: &Path, actual: &Path) -> Result<DiffReport> {
    let expected = WorkbookArchive::open(expected)?;
    let actual = WorkbookArchive::open(actual)?;
    Ok(diff_archives(&expected, &actual))
}

pub fn diff_workbooks_with_options(
    expected: &Path,
    actual: &Path,
    options: &DiffOptions,
) -> Result<DiffReport> {
    let expected = WorkbookArchive::open(expected)?;
    let actual = WorkbookArchive::open(actual)?;
    Ok(diff_archives_with_options(&expected, &actual, options))
}

/// Diff two workbooks, providing the same password for both (if either is encrypted).
///
/// For per-input passwords, use [`diff_workbooks_with_inputs`] / [`diff_workbooks_with_inputs_and_options`].
pub fn diff_workbooks_with_password(
    expected: &Path,
    actual: &Path,
    password: &str,
) -> Result<DiffReport> {
    diff_workbooks_with_password_and_options(expected, actual, password, &DiffOptions::default())
}

/// Diff two workbooks with custom diff options, providing the same password for both inputs.
pub fn diff_workbooks_with_password_and_options(
    expected: &Path,
    actual: &Path,
    password: &str,
    options: &DiffOptions,
) -> Result<DiffReport> {
    diff_workbooks_with_inputs_and_options(
        DiffInput {
            path: expected,
            password: Some(password),
        },
        DiffInput {
            path: actual,
            password: Some(password),
        },
        options,
    )
}

pub fn diff_workbooks_with_inputs(
    expected: DiffInput<'_>,
    actual: DiffInput<'_>,
) -> Result<DiffReport> {
    diff_workbooks_with_inputs_and_options(expected, actual, &DiffOptions::default())
}

pub fn diff_workbooks_with_inputs_and_options(
    expected: DiffInput<'_>,
    actual: DiffInput<'_>,
    options: &DiffOptions,
) -> Result<DiffReport> {
    let expected = WorkbookArchive::open_with_password(expected.path, expected.password)?;
    let actual = WorkbookArchive::open_with_password(actual.path, actual.password)?;
    Ok(diff_archives_with_options(&expected, &actual, options))
}

pub fn diff_archives(expected: &WorkbookArchive, actual: &WorkbookArchive) -> DiffReport {
    diff_archives_with_options(expected, actual, &DiffOptions::default())
}

pub fn diff_archives_with_options(
    expected: &WorkbookArchive,
    actual: &WorkbookArchive,
    options: &DiffOptions,
) -> DiffReport {
    let mut report = DiffReport::default();

    let ignore = IgnoreMatcher::new(options);
    let ignore_paths = IgnorePathMatcher::new(options);

    let expected_parts: BTreeSet<&str> = expected
        .part_names()
        .into_iter()
        .filter(|part| !ignore.matches(part))
        .collect();
    let actual_parts: BTreeSet<&str> = actual
        .part_names()
        .into_iter()
        .filter(|part| !ignore.matches(part))
        .collect();

    for part in expected_parts.difference(&actual_parts) {
        report.differences.push(Difference::new(
            severity_for_missing_part(part, options.strict_calc_chain),
            (*part).to_string(),
            "",
            "missing_part",
            None,
            None,
        ));
    }

    for part in actual_parts.difference(&expected_parts) {
        report.differences.push(Difference::new(
            severity_for_extra_part(part, options.strict_calc_chain),
            (*part).to_string(),
            "",
            "extra_part",
            None,
            None,
        ));
    }

    for part in expected_parts.intersection(&actual_parts) {
        let expected_bytes = expected.get(part).unwrap_or_default();
        let actual_bytes = actual.get(part).unwrap_or_default();

        // Fast path: for large corpora, most parts often round-trip byte-identically.
        // Avoid XML parsing/normalization work when there is nothing to diff.
        if expected_bytes == actual_bytes {
            #[cfg(test)]
            {
                FAST_PATH_SKIPPED_PARTS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            }
            continue;
        }

        let forced_xml = is_xml_extension(part);
        let xml_candidate =
            forced_xml || looks_like_xml(expected_bytes) || looks_like_xml(actual_bytes);

        if xml_candidate {
            match (
                xml::NormalizedXml::parse(part, expected_bytes),
                xml::NormalizedXml::parse(part, actual_bytes),
            ) {
                (Ok(expected_xml), Ok(actual_xml)) => {
                    let base = severity_for_part(part, options.strict_calc_chain);
                    let calc_chain_rel_ids = if part.ends_with(".rels") {
                        // Merge calcChain relationship ids from both sides so diffs can be
                        // downgraded even when the relationship is newly added (as opposed to
                        // removed).
                        let mut ids =
                            calc_chain_relationship_ids(part, expected_bytes).unwrap_or_default();
                        ids.extend(
                            calc_chain_relationship_ids(part, actual_bytes).unwrap_or_default(),
                        );
                        ids
                    } else {
                        BTreeSet::new()
                    };
                    let ignored_rel_ids = if part.ends_with(".rels") {
                        let expected =
                            relationship_target_ignore_map(part, expected_bytes, &ignore)
                                .unwrap_or_default();
                        let actual = relationship_target_ignore_map(part, actual_bytes, &ignore)
                            .unwrap_or_default();
                        RelationshipIgnoreMaps { expected, actual }
                    } else {
                        RelationshipIgnoreMaps::default()
                    };
                    let mut diffs = diff_xml(&expected_xml, &actual_xml, base);
                    if part.ends_with(".rels") {
                        diffs = postprocess_relationship_id_renumbering(
                            part,
                            expected_bytes,
                            actual_bytes,
                            diffs,
                            &ignore,
                        );
                    }
                    for diff in diffs {
                        if ignore_paths.matches(part, &diff.path, &diff.kind) {
                            continue;
                        }
                        if should_ignore_xml_diff(part, &diff.path, &ignored_rel_ids, &ignore) {
                            continue;
                        }
                        let severity = adjust_xml_diff_severity(
                            part,
                            diff.severity,
                            &diff.path,
                            diff.expected.as_deref(),
                            diff.actual.as_deref(),
                            &calc_chain_rel_ids,
                            options.strict_calc_chain,
                        );
                        report.differences.push(Difference::new(
                            severity,
                            part.to_string(),
                            diff.path,
                            diff.kind,
                            diff.expected,
                            diff.actual,
                        ));
                    }
                }
                (Err(err), Ok(_)) | (Ok(_), Err(err)) if forced_xml => {
                    report.differences.push(Difference::new(
                        Severity::Critical,
                        part.to_string(),
                        "",
                        "xml_parse_error",
                        None,
                        Some(err.to_string()),
                    ))
                }
                (Err(_), Ok(_)) | (Ok(_), Err(_)) => {
                    // For non-standard extensions, fall back to binary compare.
                    if expected_bytes != actual_bytes {
                        let (expected_summary, actual_summary) =
                            binary_diff_summary(expected_bytes, actual_bytes);
                        report.differences.push(Difference::new(
                            severity_for_part(part, options.strict_calc_chain),
                            part.to_string(),
                            "",
                            "binary_diff",
                            Some(expected_summary),
                            Some(actual_summary),
                        ));
                    }
                }
                (Err(err_a), Err(err_b)) if forced_xml => report.differences.push(Difference::new(
                    Severity::Critical,
                    part.to_string(),
                    "",
                    "xml_parse_error",
                    Some(err_a.to_string()),
                    Some(err_b.to_string()),
                )),
                (Err(_), Err(_)) => {
                    if expected_bytes != actual_bytes {
                        let (expected_summary, actual_summary) =
                            binary_diff_summary(expected_bytes, actual_bytes);
                        report.differences.push(Difference::new(
                            severity_for_part(part, options.strict_calc_chain),
                            part.to_string(),
                            "",
                            "binary_diff",
                            Some(expected_summary),
                            Some(actual_summary),
                        ));
                    }
                }
            }
        } else if expected_bytes != actual_bytes {
            let (expected_summary, actual_summary) =
                binary_diff_summary(expected_bytes, actual_bytes);
            report.differences.push(Difference::new(
                severity_for_part(part, options.strict_calc_chain),
                part.to_string(),
                "",
                "binary_diff",
                Some(expected_summary),
                Some(actual_summary),
            ));
        }
    }

    report
}

fn postprocess_relationship_id_renumbering(
    rels_part: &str,
    expected_bytes: &[u8],
    actual_bytes: &[u8],
    diffs: Vec<xml::XmlDiff>,
    ignore: &IgnoreMatcher,
) -> Vec<xml::XmlDiff> {
    let expected = match rels::relationship_semantic_id_map(rels_part, expected_bytes) {
        Ok(map) => map,
        Err(_) => return diffs,
    };
    let actual = match rels::relationship_semantic_id_map(rels_part, actual_bytes) {
        Ok(map) => map,
        Err(_) => return diffs,
    };

    let expected_map = expected.map;
    let actual_map = actual.map;
    let has_ambiguous_keys = expected.has_ambiguous_keys || actual.has_ambiguous_keys;

    let mut changes: Vec<rels::RelationshipIdChange> = Vec::new();
    for (key, expected_id) in &expected_map {
        let Some(actual_id) = actual_map.get(key) else {
            continue;
        };
        if expected_id != actual_id {
            changes.push(rels::RelationshipIdChange {
                key: key.clone(),
                expected_id: expected_id.clone(),
                actual_id: actual_id.clone(),
            });
        }
    }

    if changes.is_empty() {
        return diffs;
    }

    let pure_renumbering = !has_ambiguous_keys
        && expected_map.len() == actual_map.len()
        && expected_map
            .keys()
            .zip(actual_map.keys())
            .all(|(a, b)| a == b);

    if pure_renumbering {
        return changes
            .into_iter()
            .filter(|change| !ignore.matches(&change.key.resolved_target))
            .map(|change| xml::XmlDiff {
                severity: Severity::Critical,
                path: change.key.to_diff_path(),
                kind: "relationship_id_changed".to_string(),
                expected: Some(change.expected_id),
                actual: Some(change.actual_id),
            })
            .collect();
    }

    let mut suppress_missing: BTreeSet<String> = BTreeSet::new();
    let mut suppress_added: BTreeSet<String> = BTreeSet::new();
    let mut synthesized: Vec<xml::XmlDiff> = Vec::new();

    for change in changes {
        suppress_missing.insert(change.expected_id.clone());
        suppress_added.insert(change.actual_id.clone());

        // If the relationship target points to an ignored part, suppress the noisy
        // child_missing/child_added diffs but don't synthesize a new actionable diff.
        if ignore.matches(&change.key.resolved_target) {
            continue;
        }

        synthesized.push(xml::XmlDiff {
            severity: Severity::Critical,
            path: change.key.to_diff_path(),
            kind: "relationship_id_changed".to_string(),
            expected: Some(change.expected_id),
            actual: Some(change.actual_id),
        });
    }

    // If a relationship Id is reused due to a renumbering (e.g. removing one relationship and
    // shifting the remaining ones), the raw Id-keyed XML diff will report confusing attribute
    // changes for the reused Id. Synthesize explicit missing/added relationship diffs keyed by
    // relationship semantics so callers see the real change instead of attribute noise.
    for (key, expected_id) in &expected_map {
        if actual_map.contains_key(key) {
            continue;
        }
        // Only synthesize when the Id is reused on the other side by a relationship involved in
        // an Id renumbering change. Otherwise the base XML diff will already emit a
        // `child_missing` diff for this Id.
        if !suppress_added.contains(expected_id) {
            continue;
        }
        if ignore.matches(&key.resolved_target) {
            continue;
        }
        synthesized.push(xml::XmlDiff {
            severity: Severity::Critical,
            path: key.to_diff_path(),
            kind: "relationship_missing".to_string(),
            expected: Some(expected_id.clone()),
            actual: None,
        });
    }

    for (key, actual_id) in &actual_map {
        if expected_map.contains_key(key) {
            continue;
        }
        // Only synthesize when the Id is reused on the other side by a relationship involved in
        // an Id renumbering change. Otherwise the base XML diff will already emit a `child_added`
        // diff for this Id.
        if !suppress_missing.contains(actual_id) {
            continue;
        }
        if ignore.matches(&key.resolved_target) {
            continue;
        }
        synthesized.push(xml::XmlDiff {
            severity: Severity::Critical,
            path: key.to_diff_path(),
            kind: "relationship_added".to_string(),
            expected: None,
            actual: Some(actual_id.clone()),
        });
    }

    let mut out: Vec<xml::XmlDiff> = diffs
        .into_iter()
        .filter(|diff| match diff.kind.as_str() {
            "child_missing" => match relationship_id_from_path(&diff.path) {
                Some(id) => !suppress_missing.contains(id),
                None => true,
            },
            "child_added" => match relationship_id_from_path(&diff.path) {
                Some(id) => !suppress_added.contains(id),
                None => true,
            },
            "attribute_changed" | "attribute_missing" | "attribute_added" => {
                match relationship_id_from_path(&diff.path) {
                    Some(id) => !suppress_missing.contains(id) && !suppress_added.contains(id),
                    None => true,
                }
            }
            _ => true,
        })
        .collect();
    out.extend(synthesized);
    out
}

fn is_xml_extension(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels") || name.ends_with(".vml")
}

fn looks_like_xml(bytes: &[u8]) -> bool {
    if let Some(rest) = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return looks_like_xml_utf8(rest);
    }

    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        return looks_like_xml_utf16(rest, Utf16Endian::Little);
    }

    if let Some(rest) = bytes.strip_prefix(&[0xFE, 0xFF]) {
        return looks_like_xml_utf16(rest, Utf16Endian::Big);
    }

    if let Some(endian) = sniff_utf16_without_bom(bytes) {
        return looks_like_xml_utf16(bytes, endian);
    }

    looks_like_xml_utf8(bytes)
}

fn looks_like_xml_utf8(bytes: &[u8]) -> bool {
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b' ' | b'\t' | b'\n' | b'\r' => i += 1,
            b'<' => return true,
            _ => return false,
        }
    }
    false
}

fn looks_like_xml_utf16(bytes: &[u8], endian: Utf16Endian) -> bool {
    if bytes.len() % 2 != 0 {
        return false;
    }

    let mut i = 0usize;
    while i + 1 < bytes.len() {
        let word = match endian {
            Utf16Endian::Little => u16::from_le_bytes([bytes[i], bytes[i + 1]]),
            Utf16Endian::Big => u16::from_be_bytes([bytes[i], bytes[i + 1]]),
        };

        match word {
            // ASCII whitespace
            0x0009 | 0x000A | 0x000D | 0x0020 => i += 2,
            // '<'
            0x003C => return true,
            _ => return false,
        }
    }

    false
}

/// Decode raw bytes for an XML part into a Rust `&str`/`String`.
///
/// OOXML parts are commonly UTF-8, but the XML spec allows UTF-16 and Excel files
/// in the wild sometimes use it. `roxmltree` expects an already-decoded `&str`,
/// so we handle decoding here.
///
/// Supports:
/// - UTF-8, with or without a BOM (BOM is stripped)
/// - UTF-16LE/UTF-16BE via BOM (`FF FE` / `FE FF`)
/// - UTF-16LE/UTF-16BE via leading `<\0` / `\0<` patterns (optionally preceded by ASCII
///   whitespace when the XML declaration is omitted)
pub(crate) fn decode_xml_bytes(bytes: &[u8]) -> Result<Cow<'_, str>> {
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return Ok(Cow::Borrowed(std::str::from_utf8(&bytes[3..])?));
    }

    if bytes.starts_with(&[0xFF, 0xFE]) {
        return decode_utf16(&bytes[2..], Utf16Endian::Little).map(Cow::Owned);
    }
    if bytes.starts_with(&[0xFE, 0xFF]) {
        return decode_utf16(&bytes[2..], Utf16Endian::Big).map(Cow::Owned);
    }

    if let Some(endian) = sniff_utf16_without_bom(bytes) {
        return decode_utf16(bytes, endian).map(Cow::Owned);
    }

    // Fall back to UTF-8 (the overwhelmingly common encoding for OOXML parts).
    Ok(Cow::Borrowed(std::str::from_utf8(bytes)?))
}

#[derive(Debug, Clone, Copy)]
enum Utf16Endian {
    Little,
    Big,
}

fn decode_utf16(bytes: &[u8], endian: Utf16Endian) -> Result<String> {
    if bytes.len() % 2 != 0 {
        return Err(anyhow!("invalid UTF-16 byte length: {}", bytes.len()));
    }

    let mut words = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        let word = match endian {
            Utf16Endian::Little => u16::from_le_bytes([chunk[0], chunk[1]]),
            Utf16Endian::Big => u16::from_be_bytes([chunk[0], chunk[1]]),
        };
        words.push(word);
    }

    Ok(String::from_utf16(&words)?)
}

fn sniff_utf16_without_bom(bytes: &[u8]) -> Option<Utf16Endian> {
    if bytes.len() < 2 {
        return None;
    }

    // Fast path: document starts with `<`.
    if bytes[0] == b'<' && bytes[1] == 0x00 {
        return Some(Utf16Endian::Little);
    }
    if bytes[0] == 0x00 && bytes[1] == b'<' {
        return Some(Utf16Endian::Big);
    }

    // Tolerate (rare) UTF-16 documents that start with ASCII whitespace before `<`.
    //
    // Note: leading whitespace is only valid if the XML declaration is omitted,
    // but we accept it as a pragmatic heuristic for inputs seen in the wild.
    if bytes[1] == 0x00 && bytes[0].is_ascii_whitespace() {
        let mut i = 0usize;
        while i + 1 < bytes.len() {
            let a = bytes[i];
            let b = bytes[i + 1];
            if b == 0x00 && a.is_ascii_whitespace() {
                i += 2;
                continue;
            }
            if a == b'<' && b == 0x00 {
                return Some(Utf16Endian::Little);
            }
            break;
        }
    }

    if bytes[0] == 0x00 && bytes[1].is_ascii_whitespace() {
        let mut i = 0usize;
        while i + 1 < bytes.len() {
            let a = bytes[i];
            let b = bytes[i + 1];
            if a == 0x00 && b.is_ascii_whitespace() {
                i += 2;
                continue;
            }
            if a == 0x00 && b == b'<' {
                return Some(Utf16Endian::Big);
            }
            break;
        }
    }

    None
}

fn severity_for_part(part: &str, strict_calc_chain: bool) -> Severity {
    if part == "[Content_Types].xml" {
        return Severity::Critical;
    }

    if part.starts_with("docProps/") {
        return Severity::Info;
    }

    if part.ends_with(".rels") {
        return Severity::Critical;
    }

    if part.starts_with("xl/theme/") {
        return Severity::Warning;
    }

    if is_calc_chain_part(part) {
        return if strict_calc_chain {
            Severity::Critical
        } else {
            Severity::Warning
        };
    }

    Severity::Critical
}

fn severity_for_missing_part(part: &str, strict_calc_chain: bool) -> Severity {
    severity_for_part(part, strict_calc_chain)
}

fn severity_for_extra_part(part: &str, strict_calc_chain: bool) -> Severity {
    if strict_calc_chain && is_calc_chain_part(part) {
        return Severity::Critical;
    }

    if part == "[Content_Types].xml" || part.ends_with(".rels") {
        return Severity::Critical;
    }

    if part.starts_with("docProps/") {
        return Severity::Info;
    }

    Severity::Warning
}

fn is_calc_chain_part(part: &str) -> bool {
    let part = part.trim_start_matches('/');
    part.eq_ignore_ascii_case("xl/calcChain.xml") || part.eq_ignore_ascii_case("xl/calcChain.bin")
}

struct IgnoreMatcher {
    exact: BTreeSet<String>,
    globs: GlobSet,
}

impl IgnoreMatcher {
    fn new(options: &DiffOptions) -> Self {
        let exact = options
            .ignore_parts
            .iter()
            .map(|part| normalize_opc_part_name(part))
            .collect();

        let mut builder = GlobSetBuilder::new();
        for pattern in &options.ignore_globs {
            let pattern = pattern.trim();
            if pattern.is_empty() {
                continue;
            }
            let pattern = pattern.replace('\\', "/");
            let pattern = pattern.trim_start_matches('/');
            if pattern.is_empty() {
                continue;
            }
            if let Ok(glob) = Glob::new(pattern) {
                builder.add(glob);
            }
        }
        let globs = builder.build().unwrap_or_else(|_| GlobSet::empty());

        Self { exact, globs }
    }

    fn matches(&self, part: &str) -> bool {
        if self.exact.contains(part) || self.globs.is_match(part) {
            return true;
        }

        if part.starts_with('/') || part.contains('\\') {
            let normalized = normalize_opc_part_name(part);
            return self.exact.contains(&normalized) || self.globs.is_match(&normalized);
        }

        false
    }
}

struct IgnorePathMatcher {
    rules: Vec<IgnorePathRuleMatcher>,
}

#[derive(Debug, Clone)]
struct IgnorePathRuleMatcher {
    part: Option<IgnorePathPartMatcher>,
    path_substring: String,
    kind: Option<String>,
}

#[derive(Debug, Clone)]
enum IgnorePathPartMatcher {
    Exact(String),
    Glob(globset::GlobMatcher),
}

fn normalize_ignore_path_substring(raw: &str) -> String {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return String::new();
    }
    let normalized = trimmed.replace('\\', "/");
    let normalized = normalized.trim();
    if normalized.is_empty() {
        return String::new();
    }

    // `XmlDiff.path` uses resolved namespace URIs (e.g. `@{uri}uid`), not prefixes (e.g. `xr:uid`).
    // For convenience, accept simple `prefix:local` patterns by mapping them to `}local`, which
    // matches the `QName` string form without requiring callers to spell out the full URI.
    //
    // Keep this intentionally conservative so we don't rewrite URL substrings or bracketed XPath-ish
    // snippets.
    let candidate = normalized.trim_start_matches('@');
    if candidate.contains("://")
        || candidate.contains('{')
        || candidate.contains('}')
        || candidate.contains('/')
        || candidate.contains('[')
        || candidate.contains(']')
    {
        return normalized.to_string();
    }

    let mut iter = candidate.splitn(3, ':');
    let prefix = iter.next().unwrap_or_default();
    let local = iter.next();
    if iter.next().is_some() {
        return normalized.to_string();
    }
    let Some(local) = local else {
        return normalized.to_string();
    };
    if prefix.is_empty() || local.is_empty() {
        return normalized.to_string();
    }
    if !is_simple_xml_name_segment(prefix) || !is_simple_xml_name_segment(local) {
        return normalized.to_string();
    }

    format!("}}{local}")
}

fn is_simple_xml_name_segment(s: &str) -> bool {
    let mut chars = s.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }
    chars.all(|c| c.is_ascii_alphanumeric() || matches!(c, '_' | '-' | '.'))
}

impl IgnorePathMatcher {
    fn new(options: &DiffOptions) -> Self {
        let mut rules = Vec::new();
        for rule in &options.ignore_paths {
            let path_substring = normalize_ignore_path_substring(&rule.path_substring);
            if path_substring.is_empty() {
                continue;
            }

            let kind = rule
                .kind
                .as_ref()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());

            let mut part = None;
            if let Some(pattern) = rule.part.as_ref() {
                let normalized = pattern.trim().replace('\\', "/");
                let normalized = normalized.trim_start_matches('/');
                if normalized.is_empty() {
                    part = None;
                } else if !normalized.contains('*') && !normalized.contains('?') {
                    // Treat patterns without glob metacharacters as exact matches. This allows
                    // exact part names like `[Content_Types].xml` without requiring escaping glob
                    // syntax.
                    part = Some(IgnorePathPartMatcher::Exact(normalize_opc_part_name(
                        normalized,
                    )));
                } else if let Ok(glob) = Glob::new(normalized) {
                    part = Some(IgnorePathPartMatcher::Glob(glob.compile_matcher()));
                } else {
                    // Like `ignore_globs`, invalid glob patterns are ignored. Crucially, do not
                    // fall back to "match any part" when the part selector is invalid, as that
                    // could unexpectedly suppress diffs across the entire workbook.
                    continue;
                }
            }

            rules.push(IgnorePathRuleMatcher {
                part,
                path_substring,
                kind,
            });
        }

        Self { rules }
    }

    fn matches(&self, part: &str, path: &str, kind: &str) -> bool {
        self.rules.iter().any(|rule| {
            if let Some(rule_kind) = &rule.kind {
                if rule_kind != kind {
                    return false;
                }
            }

            if let Some(matcher) = &rule.part {
                let ok = match matcher {
                    IgnorePathPartMatcher::Exact(exact) => exact == part,
                    IgnorePathPartMatcher::Glob(glob) => glob.is_match(part),
                };
                if !ok {
                    return false;
                }
            }

            path.contains(&rule.path_substring)
        })
    }
}

fn should_ignore_xml_diff(
    part: &str,
    path: &str,
    ignored_rels: &RelationshipIgnoreMaps,
    ignore: &IgnoreMatcher,
) -> bool {
    if part == "[Content_Types].xml" {
        if let Some(part_name) = content_type_override_part_name_from_path(path) {
            let normalized = normalize_opc_part_name(part_name);
            if ignore.matches(&normalized) {
                return true;
            }
        }
    }

    if part.ends_with(".rels") {
        if let Some(id) = relationship_id_from_path(path) {
            let expected = ignored_rels.expected.get(id).copied();
            let actual = ignored_rels.actual.get(id).copied();
            match (expected, actual) {
                (Some(true), Some(true)) | (Some(true), None) | (None, Some(true)) => return true,
                _ => {}
            }
        }
    }

    false
}

fn content_type_override_part_name_from_path(path: &str) -> Option<&str> {
    let needle = "Override[@PartName=\"";
    let start = path.find(needle)? + needle.len();
    let rest = &path[start..];
    let end = rest.find('"')?;
    Some(&rest[..end])
}

fn normalize_opc_part_name(part_name: &str) -> String {
    normalize_opc_path(part_name.trim_start_matches('/'))
}

fn normalize_opc_path(path: &str) -> String {
    let normalized = path.replace('\\', "/");
    let mut out: Vec<&str> = Vec::new();
    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            _ => out.push(segment),
        }
    }
    out.join("/")
}

#[derive(Debug, Default)]
struct RelationshipIgnoreMaps {
    expected: BTreeMap<String, bool>,
    actual: BTreeMap<String, bool>,
}

fn relationship_target_ignore_map(
    rels_part: &str,
    bytes: &[u8],
    ignore: &IgnoreMatcher,
) -> Result<BTreeMap<String, bool>> {
    let text =
        decode_xml_bytes(bytes).with_context(|| format!("decode xml bytes for {rels_part}"))?;
    let doc =
        Document::parse(text.as_ref()).with_context(|| format!("parse xml for {rels_part}"))?;

    let mut ids = BTreeMap::new();
    for node in doc
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        let Some(id) = node.attribute("Id") else {
            continue;
        };
        let target = node.attribute("Target").unwrap_or_default();
        let resolved = match node.attribute("TargetMode") {
            Some(mode) if mode.eq_ignore_ascii_case("External") => target.replace('\\', "/"),
            _ => resolve_relationship_target(rels_part, target),
        };
        ids.insert(id.to_string(), ignore.matches(&resolved));
    }

    Ok(ids)
}

fn resolve_relationship_target(rels_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    // Relationship targets are URIs; internal targets may include a fragment (e.g. `foo.xml#bar`).
    // OPC part names do not include fragments, so strip them before resolving.
    let target = target.split_once('#').map(|(t, _)| t).unwrap_or(&target);
    if target.is_empty() {
        // A target of just `#fragment` refers to the relationship source part itself.
        return source_part_from_rels_part(rels_part);
    }
    if let Some(rest) = target.strip_prefix('/') {
        return normalize_opc_path(rest);
    }

    // Non-standard but common: some producers emit targets like `xl/media/image1.png` (no leading
    // `/`) inside `xl/_rels/workbook.xml.rels`. Per OPC rules this would normally be resolved
    // relative to the source part directory (and therefore become `xl/xl/media/...`), but the
    // intent is clearly to refer to the package-root `xl/*` tree. Treat these as absolute package
    // paths so ignore rules like `xl/media/*` still work.
    let normalized_target = normalize_opc_path(target);
    if normalized_target
        .get(..3)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("xl/"))
    {
        return normalized_target;
    }

    let base_dir = rels_base_dir(rels_part);
    normalize_opc_path(&format!("{base_dir}{target}"))
}

fn source_part_from_rels_part(rels_part: &str) -> String {
    if rels_part == "_rels/.rels" {
        return String::new();
    }

    if let Some(rels_file) = rels_part.strip_prefix("_rels/") {
        return normalize_opc_path(rels_file.strip_suffix(".rels").unwrap_or(rels_file));
    }

    if let Some((dir, rels_file)) = rels_part.rsplit_once("/_rels/") {
        let rels_file = rels_file.strip_suffix(".rels").unwrap_or(rels_file);
        if dir.is_empty() {
            return normalize_opc_path(rels_file);
        }
        return normalize_opc_path(&format!("{dir}/{rels_file}"));
    }

    normalize_opc_path(rels_part.strip_suffix(".rels").unwrap_or(rels_part))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_relationship_target_strips_uri_fragments() {
        assert_eq!(
            resolve_relationship_target("xl/_rels/workbook.xml.rels", "worksheets/sheet1.xml#frag"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_relationship_target("xl/_rels/workbook.xml.rels", "/xl/media/image1.png#frag"),
            "xl/media/image1.png"
        );
        assert_eq!(
            resolve_relationship_target("xl/_rels/workbook.xml.rels", "#frag"),
            "xl/workbook.xml"
        );
        assert_eq!(
            resolve_relationship_target("xl/worksheets/_rels/sheet1.xml.rels", "#frag"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(resolve_relationship_target("_rels/.rels", "#frag"), "");
    }

    #[test]
    fn diff_archives_fast_path_skips_byte_identical_parts() {
        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert("xl/workbook.xml".to_string(), b"<workbook/>".to_vec());
        parts.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#
                .to_vec(),
        );

        let expected = WorkbookArchive {
            parts: parts.clone(),
        };
        let actual = WorkbookArchive { parts };

        let before = FAST_PATH_SKIPPED_PARTS.load(std::sync::atomic::Ordering::Relaxed);
        let report = diff_archives(&expected, &actual);
        let after = FAST_PATH_SKIPPED_PARTS.load(std::sync::atomic::Ordering::Relaxed);

        assert!(report.is_empty());
        assert!(
            after > before,
            "expected fast path to skip at least one part (before={before}, after={after})"
        );
    }

    #[test]
    fn resolve_relationship_target_tolerates_xl_prefixed_targets() {
        assert_eq!(
            resolve_relationship_target("xl/_rels/workbook.xml.rels", "xl/media/image1.png"),
            "xl/media/image1.png"
        );
    }
}

fn rels_base_dir(rels_part: &str) -> String {
    if rels_part.starts_with("_rels/") {
        return String::new();
    }

    if let Some(pos) = rels_part.rfind("/_rels/") {
        return rels_part[..pos + 1].to_string();
    }

    String::new()
}

fn adjust_xml_diff_severity(
    part: &str,
    severity: Severity,
    path: &str,
    expected: Option<&str>,
    actual: Option<&str>,
    calc_chain_rel_ids: &BTreeSet<String>,
    strict_calc_chain: bool,
) -> Severity {
    if severity != Severity::Critical {
        return severity;
    }

    if strict_calc_chain {
        return severity;
    }

    // CalcChain invalidation intentionally rewrites workbook relationships and content types.
    // Downgrade diffs that are clearly about calcChain so tests can assert “no unexpected
    // critical diffs” for patch/streaming flows.
    if part == "[Content_Types].xml"
        && (mentions_calc_chain(path)
            || expected.is_some_and(mentions_calc_chain)
            || actual.is_some_and(mentions_calc_chain))
    {
        return Severity::Warning;
    }

    if part.ends_with(".rels") {
        if mentions_calc_chain(path)
            || expected.is_some_and(mentions_calc_chain)
            || actual.is_some_and(mentions_calc_chain)
        {
            return Severity::Warning;
        }

        if let Some(id) = relationship_id_from_path(path) {
            if calc_chain_rel_ids.contains(id) {
                return Severity::Warning;
            }
        }
    }

    severity
}

fn mentions_calc_chain(value: &str) -> bool {
    value.to_ascii_lowercase().contains("calcchain")
}

fn relationship_id_from_path(path: &str) -> Option<&str> {
    let needle = "[@Id=\"";
    let start = path.find(needle)? + needle.len();
    let rest = &path[start..];
    let end = rest.find('"')?;
    Some(&rest[..end])
}

fn calc_chain_relationship_ids(part_name: &str, bytes: &[u8]) -> Result<BTreeSet<String>> {
    let text =
        decode_xml_bytes(bytes).with_context(|| format!("decode xml bytes for {part_name}"))?;
    let doc =
        Document::parse(text.as_ref()).with_context(|| format!("parse xml for {part_name}"))?;

    let mut ids = BTreeSet::new();
    for node in doc
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        let Some(id) = node.attribute("Id") else {
            continue;
        };
        let ty = node.attribute("Type").unwrap_or_default();
        let target = node.attribute("Target").unwrap_or_default();
        if is_calc_chain_relationship(ty, target) {
            ids.insert(id.to_string());
        }
    }

    Ok(ids)
}

fn is_calc_chain_relationship(ty: &str, target: &str) -> bool {
    let target = target.replace('\\', "/").to_ascii_lowercase();
    if target.ends_with("calcchain.xml") || target.ends_with("calcchain.bin") {
        return true;
    }

    ty.to_ascii_lowercase().contains("relationships/calcchain")
}

fn binary_diff_summary(expected: &[u8], actual: &[u8]) -> (String, String) {
    let expected_len = expected.len();
    let actual_len = actual.len();

    let min_len = expected_len.min(actual_len);
    let mut first_diff = None;
    for idx in 0..min_len {
        if expected[idx] != actual[idx] {
            first_diff = Some(idx);
            break;
        }
    }
    if first_diff.is_none() && expected_len != actual_len {
        first_diff = Some(min_len);
    }

    let mut expected_summary = format!("{expected_len} bytes");
    let mut actual_summary = format!("{actual_len} bytes");
    if let Some(offset) = first_diff {
        expected_summary.push_str(&format!(" (first diff at offset {offset})"));
        actual_summary.push_str(&format!(" (first diff at offset {offset})"));
    }

    (expected_summary, actual_summary)
}

/// A minimal “load → save” round-trip that preserves each part byte-for-byte.
///
/// This is a convenience helper for situations where you want to validate the
/// diff tooling itself without invoking a higher-level XLSX writer. The CI
/// harness uses `formula-xlsx` for its round-trip path.
pub fn roundtrip_zip_copy(original: &Path, out_path: &Path) -> Result<()> {
    let src_file =
        File::open(original).with_context(|| format!("open workbook {}", original.display()))?;
    let mut archive = ZipArchive::new(src_file).context("parse zip archive")?;

    let dst_file = File::create(out_path)
        .with_context(|| format!("create workbook {}", out_path.display()))?;
    let mut writer = ZipWriter::new(dst_file);

    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).context("read zip entry")?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
        // advertise enormous uncompressed sizes (zip-bomb style OOM).
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)
            .with_context(|| format!("read part {name}"))?;

        writer
            .start_file(name, options)
            .context("write zip entry header")?;
        writer.write_all(&buf).context("write zip entry")?;
    }

    writer.finish().context("finalize zip archive")?;
    Ok(())
}

/// Collect `.xlsx` / `.xlsm` / `.xlsb` fixture paths under `root`.
///
/// Note: this helper is used by round-trip test harnesses that treat each file as a ZIP/OPC
/// package. Password-protected/encrypted OOXML workbooks that use an OLE/CFB wrapper
/// (`EncryptionInfo` + `EncryptedPackage`) are **not ZIP archives** and must be excluded from this
/// corpus (they otherwise fail ZIP-based round-trip tests).
///
/// Convention: keep them under `fixtures/xlsx/encrypted/` (or `fixtures/encrypted/ooxml/`) and this
/// helper will skip the `encrypted/` subtree.
pub fn collect_fixture_paths(root: &Path) -> Result<Vec<PathBuf>> {
    if !root.exists() {
        return Err(anyhow!("fixtures root {} does not exist", root.display()));
    }

    let encrypted_dir = root.join("encrypted");
    let mut files = Vec::new();
    for entry in walkdir::WalkDir::new(root)
        .follow_links(false)
        .into_iter()
        .filter_map(|e| e.ok())
    {
        if !entry.file_type().is_file() {
            continue;
        }
        let path = entry.path();
        if path.starts_with(&encrypted_dir) {
            continue;
        }
        match path.extension().and_then(|s| s.to_str()) {
            Some("xlsx") | Some("xlsm") | Some("xlsb") => files.push(path.to_path_buf()),
            _ => {}
        }
    }

    files.sort();
    Ok(files)
}
