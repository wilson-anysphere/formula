//! XLSX round-trip diff tooling.
//!
//! This crate intentionally operates at the ZIP/Open Packaging Convention layer:
//! it compares workbook "parts" (files within the archive) rather than the ZIP
//! container bytes. This avoids false positives from differing compression or
//! timestamp metadata while still catching fidelity regressions.

mod xml;

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::fs::File;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub use xml::{diff_xml, NormalizedXml};

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

        let mut parts = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i).context("read zip entry")?;
            if file.is_dir() {
                continue;
            }
            let name = file.name().to_string();
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)
                .with_context(|| format!("read part {name}"))?;
            parts.insert(name, buf);
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

pub fn diff_workbooks(expected: &Path, actual: &Path) -> Result<DiffReport> {
    let expected = WorkbookArchive::open(expected)?;
    let actual = WorkbookArchive::open(actual)?;
    Ok(diff_archives(&expected, &actual))
}

pub fn diff_archives(expected: &WorkbookArchive, actual: &WorkbookArchive) -> DiffReport {
    let mut report = DiffReport::default();

    let expected_parts = expected.part_names();
    let actual_parts = actual.part_names();

    for part in expected_parts.difference(&actual_parts) {
        report.differences.push(Difference::new(
            Severity::Critical,
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
                    for diff in diff_xml(&expected_xml, &actual_xml, base) {
                        report.differences.push(Difference::new(
                            diff.severity,
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
                        report.differences.push(Difference::new(
                            Severity::Critical,
                            part.to_string(),
                            "",
                            "binary_diff",
                            Some(format!("{} bytes", expected_bytes.len())),
                            Some(format!("{} bytes", actual_bytes.len())),
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
                        report.differences.push(Difference::new(
                            Severity::Critical,
                            part.to_string(),
                            "",
                            "binary_diff",
                            Some(format!("{} bytes", expected_bytes.len())),
                            Some(format!("{} bytes", actual_bytes.len())),
                        ));
                    }
                }
            }
        } else if expected_bytes != actual_bytes {
            report.differences.push(Difference::new(
                Severity::Critical,
                part.to_string(),
                "",
                "binary_diff",
                Some(format!("{} bytes", expected_bytes.len())),
                Some(format!("{} bytes", actual_bytes.len())),
            ));
        }
    }

    report
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

    if part.starts_with("xl/theme/") || part == "xl/calcChain.xml" {
        return Severity::Warning;
    }

    Severity::Critical
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
            Some("xlsx") | Some("xlsm") => files.push(path.to_path_buf()),
            _ => {}
        }
    }

    files.sort();
    Ok(files)
}
