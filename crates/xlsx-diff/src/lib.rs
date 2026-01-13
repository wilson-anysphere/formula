//! XLSX/XLSB round-trip diff tooling.
//!
//! This crate intentionally operates at the ZIP/Open Packaging Convention layer:
//! it compares workbook "parts" (files within the archive) rather than the ZIP
//! container bytes. This avoids false positives from differing compression or
//! timestamp metadata while still catching fidelity regressions.

mod part_kind;
mod xml;
mod rels;

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::fs::File;
use std::io::Cursor;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use globset::{Glob, GlobSet, GlobSetBuilder};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub use part_kind::{classify_part, PartKind};
pub use xml::{diff_xml, NormalizedXml};

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
        let mut parts = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i).context("read zip entry")?;
            if file.is_dir() {
                continue;
            }
            let name = normalize_opc_part_name(file.name());
            let mut buf = Vec::with_capacity(file.size() as usize);
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
            severity_for_extra_part(part),
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
                    let base = severity_for_part(part);
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
                        if should_ignore_xml_diff(part, &diff.path, &ignored_rel_ids, &ignore) {
                            continue;
                        }
                        let severity = if diff.kind == "relationship_id_changed" {
                            Severity::Critical
                        } else {
                            adjust_xml_diff_severity(
                                part,
                                diff.severity,
                                &diff.path,
                                diff.expected.as_deref(),
                                diff.actual.as_deref(),
                                &calc_chain_rel_ids,
                            )
                        };
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
                            severity_for_part(part),
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
                            severity_for_part(part),
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
                severity_for_part(part),
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
    let changes = match rels::detect_relationship_id_changes(rels_part, expected_bytes, actual_bytes)
    {
        Ok(changes) => changes,
        Err(_) => return diffs,
    };

    if changes.is_empty() {
        return diffs;
    }

    let mut missing_ids: BTreeSet<String> = BTreeSet::new();
    let mut added_ids: BTreeSet<String> = BTreeSet::new();
    for diff in &diffs {
        match diff.kind.as_str() {
            "child_missing" => {
                if let Some(id) = relationship_id_from_path(&diff.path) {
                    missing_ids.insert(id.to_string());
                }
            }
            "child_added" => {
                if let Some(id) = relationship_id_from_path(&diff.path) {
                    added_ids.insert(id.to_string());
                }
            }
            _ => {}
        }
    }

    let mut suppress_missing: BTreeSet<String> = BTreeSet::new();
    let mut suppress_added: BTreeSet<String> = BTreeSet::new();
    let mut synthesized: Vec<xml::XmlDiff> = Vec::new();

    for change in changes {
        // Only treat it as a "pure" Id renumbering if the XML diff algorithm already
        // classified it as a missing + added relationship. This avoids double-reporting
        // cases like Id reuse/swaps where the key-by-Id algorithm will instead surface
        // attribute changes.
        if !missing_ids.contains(&change.expected_id) || !added_ids.contains(&change.actual_id) {
            continue;
        }

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

    if suppress_missing.is_empty() && suppress_added.is_empty() {
        return diffs;
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
    let mut i = 0usize;
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        i = 3;
    }
    while i < bytes.len() {
        match bytes[i] {
            b' ' | b'\t' | b'\n' | b'\r' => i += 1,
            b'<' => return true,
            _ => return false,
        }
    }
    false
}

fn severity_for_part(part: &str) -> Severity {
    if part == "[Content_Types].xml" {
        return Severity::Critical;
    }

    if part.starts_with("docProps/") {
        return Severity::Info;
    }

    if part.ends_with(".rels") {
        return Severity::Critical;
    }

    if part.starts_with("xl/theme/") || is_calc_chain_part(part) {
        return Severity::Warning;
    }

    Severity::Critical
}

fn severity_for_missing_part(part: &str) -> Severity {
    severity_for_part(part)
}

fn severity_for_extra_part(part: &str) -> Severity {
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
    let text = std::str::from_utf8(bytes)
        .with_context(|| format!("part {rels_part} is not valid UTF-8"))?;
    let doc = Document::parse(text).with_context(|| format!("parse xml for {rels_part}"))?;

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
) -> Severity {
    if severity != Severity::Critical {
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
        let mut buf = Vec::with_capacity(file.size() as usize);
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
            Some("xlsx") | Some("xlsm") | Some("xlsb") => files.push(path.to_path_buf()),
            _ => {}
        }
    }

    files.sort();
    Ok(files)
}
