//! XLSB round-trip diff tooling.
//!
//! This crate intentionally operates at the ZIP/Open Packaging Convention layer:
//! it compares workbook "parts" (files within the archive) rather than the ZIP
//! container bytes. This avoids false positives from differing compression or
//! timestamp metadata while still catching fidelity regressions.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::fs::File;
use std::io::Cursor;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use anyhow::{anyhow, Context, Result};
use roxmltree::Document;
use xlsx_diff::NormalizedXml;
use zip::ZipArchive;

pub use xlsx_diff::Severity;

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

/// In-memory representation of an XLSB ZIP package.
///
/// This is intentionally cloneable without duplicating payloads (the part bytes
/// are reference-counted).
#[derive(Debug, Clone)]
pub struct WorkbookArchive {
    parts: BTreeMap<String, Arc<Vec<u8>>>,
}

impl WorkbookArchive {
    pub fn open(path: &Path) -> Result<Self> {
        let file = File::open(path).with_context(|| format!("open workbook {}", path.display()))?;
        let mut zip = ZipArchive::new(file).context("parse zip archive")?;
        Self::read_zip(&mut zip)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let mut zip = ZipArchive::new(cursor).context("parse zip archive")?;
        Self::read_zip(&mut zip)
    }

    fn read_zip<R: Read + std::io::Seek>(zip: &mut ZipArchive<R>) -> Result<Self> {
        let mut parts: BTreeMap<String, Arc<Vec<u8>>> = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i).context("read zip entry")?;
            if file.is_dir() {
                continue;
            }

            // ZIP entry names in valid XLSB packages should not start with `/`, but some producers
            // emit them. Normalize entry names so diffing doesn't report spurious missing/extra
            // parts when one side uses a leading slash.
            //
            // Keep the original name as a fallback key if normalization would collide with an
            // existing part (e.g. both `xl/workbook.bin` and `/xl/workbook.bin` exist).
            let original_name = file.name();
            let name = normalize_zip_entry_name(original_name);
            let name = if parts.contains_key(&name) && name != original_name {
                original_name.to_string()
            } else {
                name
            };
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)
                .with_context(|| format!("read part {name}"))?;
            parts.insert(name, Arc::new(buf));
        }

        Ok(Self { parts })
    }

    pub fn parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.parts
            .iter()
            .map(|(k, v)| (k.as_str(), v.as_slice()))
    }

    pub fn part_names(&self) -> BTreeSet<&str> {
        self.parts.keys().map(|k| k.as_str()).collect()
    }

    pub fn get(&self, name: &str) -> Option<&[u8]> {
        self.parts.get(name).map(|v| v.as_slice())
    }
}

fn normalize_zip_entry_name(name: &str) -> String {
    let mut normalized = name.trim_start_matches('/');
    let replaced;
    if normalized.contains('\\') {
        replaced = normalized.replace('\\', "/");
        normalized = &replaced;
    }
    normalized.to_string()
}

#[derive(Debug, Clone, Default)]
pub struct DiffOptions {
    pub ignore_parts: BTreeSet<String>,
}

pub fn diff_workbooks(expected: &Path, actual: &Path) -> Result<DiffReport> {
    let expected = WorkbookArchive::open(expected)?;
    let actual = WorkbookArchive::open(actual)?;
    Ok(diff_archives(&expected, &actual))
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

    let expected_parts: BTreeSet<&str> = expected
        .part_names()
        .into_iter()
        .filter(|part| !options.ignore_parts.contains(*part))
        .collect();
    let actual_parts: BTreeSet<&str> = actual
        .part_names()
        .into_iter()
        .filter(|part| !options.ignore_parts.contains(*part))
        .collect();

    for part in expected_parts.difference(&actual_parts) {
        report.differences.push(Difference::new(
            severity_for_missing_part(part),
            (*part).to_string(),
            "",
            "missing_part",
            None,
            None,
        ));
    }

    for part in actual_parts.difference(&expected_parts) {
        report.differences.push(Difference::new(
            Severity::Warning,
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

        if is_xmlish_part(part) {
            let calc_chain_rel_ids = if part.ends_with(".rels") {
                // Merge calcChain relationship ids from both sides so diffs can be downgraded even
                // when the relationship is newly added (as opposed to removed).
                let mut ids = calc_chain_relationship_ids(part, expected_bytes).unwrap_or_default();
                ids.extend(calc_chain_relationship_ids(part, actual_bytes).unwrap_or_default());
                ids
            } else {
                BTreeSet::new()
            };

            match (
                NormalizedXml::parse(part, expected_bytes),
                NormalizedXml::parse(part, actual_bytes),
            ) {
                (Ok(expected_xml), Ok(actual_xml)) => {
                    let base = base_severity_for_xml_part(part);
                    for diff in xlsx_diff::diff_xml(&expected_xml, &actual_xml, base) {
                        let severity = adjust_xml_diff_severity(
                            part,
                            diff.severity,
                            &diff.path,
                            diff.expected.as_deref(),
                            diff.actual.as_deref(),
                            &calc_chain_rel_ids,
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
                (Err(err), Ok(_)) | (Ok(_), Err(err)) => report.differences.push(Difference::new(
                    Severity::Critical,
                    part.to_string(),
                    "",
                    "xml_parse_error",
                    None,
                    Some(err.to_string()),
                )),
                (Err(err_a), Err(err_b)) => report.differences.push(Difference::new(
                    Severity::Critical,
                    part.to_string(),
                    "",
                    "xml_parse_error",
                    Some(err_a.to_string()),
                    Some(err_b.to_string()),
                )),
            }
        } else if expected_bytes != actual_bytes {
            let severity = severity_for_binary_part(part);
            let (expected_summary, actual_summary) = binary_diff_summary(expected_bytes, actual_bytes);
            report.differences.push(Difference::new(
                severity,
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

fn is_xmlish_part(part: &str) -> bool {
    part.ends_with(".xml") || part.ends_with(".rels")
}

fn is_calc_chain_part(part: &str) -> bool {
    part.trim_start_matches('/')
        .eq_ignore_ascii_case("xl/calcChain.bin")
}

fn severity_for_missing_part(part: &str) -> Severity {
    if is_calc_chain_part(part) {
        // `formula-xlsb` intentionally drops calcChain on edits (and therefore the
        // part disappears). Treat this as expected, warning-level churn.
        Severity::Warning
    } else {
        Severity::Critical
    }
}

fn severity_for_binary_part(part: &str) -> Severity {
    if is_calc_chain_part(part) {
        Severity::Warning
    } else {
        Severity::Critical
    }
}

fn base_severity_for_xml_part(part: &str) -> Severity {
    if part.starts_with("docProps/") {
        return Severity::Info;
    }

    // `.rels` and `[Content_Types].xml` are core package plumbing for XLSB.
    Severity::Critical
}

fn adjust_xml_diff_severity(
    part: &str,
    severity: Severity,
    path: &str,
    expected: Option<&str>,
    actual: Option<&str>,
    calc_chain_rel_ids: &BTreeSet<String>,
) -> Severity {
    if severity != Severity::Critical {
        return severity;
    }

    // CalcChain invalidation intentionally rewrites workbook relationships and
    // content types. Downgrade diffs that are clearly about calcChain so tests
    // can assert “no unexpected critical diffs” for patch flows.
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
    let text = std::str::from_utf8(bytes)
        .with_context(|| format!("part {part_name} is not valid UTF-8"))?;
    let doc = Document::parse(text).with_context(|| format!("parse xml for {part_name}"))?;

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
    if target.ends_with("calcchain.bin") {
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

pub fn collect_fixture_paths(root: &Path) -> Result<Vec<PathBuf>> {
    if !root.exists() {
        return Err(anyhow!("fixtures root {} does not exist", root.display()));
    }

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
        match path.extension().and_then(|s| s.to_str()) {
            Some("xlsb") => files.push(path.to_path_buf()),
            _ => {}
        }
    }

    files.sort();
    Ok(files)
}
