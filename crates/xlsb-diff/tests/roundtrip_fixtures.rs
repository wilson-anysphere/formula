use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::path::Path;

use anyhow::{Context, Result};
use formula_xlsb::XlsbWorkbook;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn format_report(report: &xlsb_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn roundtrip_fixtures_no_critical_diffs() -> Result<()> {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../formula-xlsb/tests/fixtures");
    let fixtures = xlsb_diff::collect_fixture_paths(&fixtures_root)?;
    assert!(
        !fixtures.is_empty(),
        "no fixtures found under {}",
        fixtures_root.display()
    );

    for fixture in fixtures {
        let tmpdir = tempfile::tempdir()?;
        let roundtripped = tmpdir.path().join("roundtripped.xlsb");

        let wb =
            XlsbWorkbook::open(&fixture).with_context(|| format!("open fixture {}", fixture.display()))?;
        wb.save_as(&roundtripped)
            .with_context(|| format!("save_as {}", fixture.display()))?;

        let report = xlsb_diff::diff_workbooks(&fixture, &roundtripped)?;
        if report.has_at_least(xlsb_diff::Severity::Critical) {
            eprintln!("Critical diffs detected for fixture {}", fixture.display());
            for diff in report
                .differences
                .iter()
                .filter(|d| d.severity == xlsb_diff::Severity::Critical)
            {
                eprintln!("{diff}");
            }
            panic!(
                "fixture {} did not round-trip cleanly (critical diffs present)",
                fixture.display()
            );
        }
    }

    Ok(())
}

fn insert_before_closing_tag(mut xml: String, closing_tag: &str, insert: &str) -> String {
    let idx = xml
        .rfind(closing_tag)
        .unwrap_or_else(|| panic!("missing closing tag {closing_tag}"));
    xml.insert_str(idx, insert);
    xml
}

fn max_rid_suffix(xml: &str) -> u32 {
    let mut max = 0u32;
    for chunk in xml.split("Id=\"rId").skip(1) {
        let digits: String = chunk.chars().take_while(|c| c.is_ascii_digit()).collect();
        if let Ok(n) = digits.parse::<u32>() {
            max = max.max(n);
        }
    }
    max
}

fn build_fixture_with_calc_chain(base_bytes: &[u8]) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(base_bytes)).expect("open base xlsb zip");
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();

    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("read zip entry");
        if !file.is_file() {
            continue;
        }
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read part bytes");
        parts.insert(file.name().to_string(), buf);
    }

    parts.insert("xl/calcChain.bin".to_string(), b"dummy".to_vec());

    let content_types =
        String::from_utf8(parts["[Content_Types].xml"].clone()).expect("utf8 content types");
    let content_types = insert_before_closing_tag(
        content_types,
        "</Types>",
        "  <Override PartName=\"/xl/calcChain.bin\" ContentType=\"application/vnd.ms-excel.calcChain\"/>\n",
    );
    parts.insert("[Content_Types].xml".to_string(), content_types.into_bytes());

    let workbook_rels =
        String::from_utf8(parts["xl/_rels/workbook.bin.rels"].clone()).expect("utf8 workbook rels");
    let max_rid = max_rid_suffix(&workbook_rels);
    let calc_chain_rid = format!("rId{}", max_rid + 1);
    let workbook_rels = insert_before_closing_tag(
        workbook_rels,
        "</Relationships>",
        &format!(
            "  <Relationship Id=\"{calc_chain_rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.bin\"/>\n"
        ),
    );
    parts.insert("xl/_rels/workbook.bin.rels".to_string(), workbook_rels.into_bytes());

    let cursor = Cursor::new(Vec::new());
    let mut zip_out = ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip_out.start_file(name, options).expect("write part header");
        zip_out.write_all(&bytes).expect("write part bytes");
    }

    zip_out.finish().expect("finish zip").into_inner()
}

#[test]
fn patch_roundtrip_only_changes_expected_parts() -> Result<()> {
    let fixtures_root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../formula-xlsb/tests/fixtures");
    let base_fixture = fixtures_root.join("simple.xlsb");
    let base_bytes = std::fs::read(&base_fixture)
        .with_context(|| format!("read fixture {}", base_fixture.display()))?;
    let variant_bytes = build_fixture_with_calc_chain(&base_bytes);

    let tmpdir = tempfile::tempdir()?;
    let input_path = tmpdir.path().join("with_calc_chain.xlsb");
    let patched_path = tmpdir.path().join("patched.xlsb");
    std::fs::write(&input_path, variant_bytes)
        .with_context(|| format!("write temp fixture {}", input_path.display()))?;

    let wb = XlsbWorkbook::open(&input_path).context("open patched fixture")?;
    let sheet_part = wb
        .sheet_metas()
        .get(0)
        .context("fixture workbook has no sheets")?
        .part_path
        .clone();

    // Mutate B1 on the first sheet.
    wb.save_with_edits(&patched_path, 0, 0, 1, 123.0)
        .context("save_with_edits")?;

    let report = xlsb_diff::diff_workbooks(&input_path, &patched_path)?;
    let report_text = format_report(&report);

    // The worksheet payload must change and remain a critical binary diff.
    let sheet_diffs: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.part == sheet_part)
        .collect();
    assert!(
        !sheet_diffs.is_empty(),
        "expected patched worksheet part {sheet_part} to differ; report:\n{report_text}"
    );
    assert!(
        sheet_diffs
            .iter()
            .all(|d| d.severity == xlsb_diff::Severity::Critical),
        "patched worksheet diffs should be CRITICAL; report:\n{report_text}"
    );

    // CalcChain invalidation may cause warning-level churn in plumbing parts.
    for diff in report.differences.iter().filter(|d| d.part != sheet_part) {
        assert!(
            diff.severity != xlsb_diff::Severity::Critical,
            "unexpected CRITICAL diff outside patched sheet:\n{diff}\nfull report:\n{report_text}"
        );
    }

    let allowed_parts: BTreeSet<String> = [
        sheet_part.clone(),
        "xl/calcChain.bin".to_string(),
        "[Content_Types].xml".to_string(),
        "xl/_rels/workbook.bin.rels".to_string(),
    ]
    .into_iter()
    .collect();
    let diff_parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = diff_parts.difference(&allowed_parts).cloned().collect();
    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{report_text}"
    );

    let extra_parts: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "extra_part")
        .map(|d| d.part.clone())
        .collect();
    assert!(
        extra_parts.is_empty(),
        "unexpected extra parts in diff: {extra_parts:?}\n{report_text}"
    );

    let missing_parts: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part")
        .map(|d| (d.part.clone(), d.severity))
        .collect();
    assert_eq!(
        missing_parts,
        vec![("xl/calcChain.bin".to_string(), xlsb_diff::Severity::Warning)],
        "expected only xl/calcChain.bin to be missing (warning-level); report:\n{report_text}"
    );

    Ok(())
}
