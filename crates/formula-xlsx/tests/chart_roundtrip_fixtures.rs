use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use formula_xlsx::load_from_bytes;
use rust_xlsxwriter::{Chart, ChartType as XlsxChartType, Workbook};
use tempfile::tempdir;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

const CHART_PART_PREFIXES: &[&str] = &["xl/charts/", "xl/drawings/", "xl/media/"];

#[test]
fn chart_roundtrip_fixtures_preserve_chart_opc_parts() -> Result<(), Box<dyn Error>> {
    let fixtures_root = chart_fixtures_root();
    let mut fixtures = xlsx_diff::collect_fixture_paths(&fixtures_root)?;
    fixtures.retain(|path| path.extension().and_then(|ext| ext.to_str()) == Some("xlsx"));

    assert!(
        !fixtures.is_empty(),
        "no .xlsx fixtures found under {}",
        fixtures_root.display()
    );

    for fixture in fixtures {
        assert_fixture_roundtrip_preserves_charts(&fixture)?;
    }

    Ok(())
}

fn chart_fixtures_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let preferred = manifest_dir.join("../../fixtures/charts/xlsx");
    if preferred.exists() {
        return preferred;
    }
    manifest_dir.join("../../fixtures/xlsx/charts")
}

fn assert_fixture_roundtrip_preserves_charts(fixture: &Path) -> Result<(), Box<dyn Error>> {
    let original_bytes = std::fs::read(fixture)?;
    let doc = load_from_bytes(&original_bytes)?;
    let roundtripped_bytes = doc.save_to_vec()?;

    let tmpdir = tempdir()?;
    let roundtripped_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&roundtripped_path, &roundtripped_bytes)?;

    let report = xlsx_diff::diff_workbooks(fixture, &roundtripped_path)?;
    let expected = xlsx_diff::WorkbookArchive::open(fixture)?;
    let actual = xlsx_diff::WorkbookArchive::open(&roundtripped_path)?;

    let expected_parts: BTreeSet<String> = expected.part_names().into_iter().map(str::to_string).collect();
    let actual_parts: BTreeSet<String> = actual.part_names().into_iter().map(str::to_string).collect();

    let mut issues: BTreeMap<String, Vec<String>> = BTreeMap::new();

    for part in expected_parts.difference(&actual_parts) {
        if part.starts_with("xl/charts/") || part.starts_with("xl/drawings/") {
            issues
                .entry(part.to_string())
                .or_default()
                .push("missing_part".to_string());
        }
    }

    for part in expected_parts.iter() {
        if !part.ends_with(".rels") {
            continue;
        }

        let expected_bytes = expected
            .get(part)
            .ok_or_else(|| format!("expected archive missing part {part}"))?;
        match actual.get(part) {
            Some(actual_bytes) if actual_bytes == expected_bytes => {}
            Some(actual_bytes) => issues
                .entry(part.to_string())
                .or_default()
                .push(format!(
                    "rels_bytes_changed (expected {}, actual {}, first_diff {})",
                    expected_bytes.len(),
                    actual_bytes.len(),
                    first_diff_offset(expected_bytes, actual_bytes)
                        .map(|idx| idx.to_string())
                        .unwrap_or_else(|| "unknown".to_string())
                )),
            None => issues
                .entry(part.to_string())
                .or_default()
                .push("rels_missing".to_string()),
        }
    }

    for part in actual_parts.difference(&expected_parts) {
        if part.ends_with(".rels") {
            issues
                .entry(part.to_string())
                .or_default()
                .push("extra_rels_part".to_string());
        }
    }

    for part in expected_parts.iter() {
        if !is_chart_related_part(part) {
            continue;
        }

        let expected_bytes = expected
            .get(part)
            .ok_or_else(|| format!("expected archive missing part {part}"))?;
        match actual.get(part) {
            Some(actual_bytes) if actual_bytes == expected_bytes => {}
            Some(actual_bytes) => {
                let mut msg = format!(
                    "chart_part_bytes_changed (expected {}, actual {}, first_diff {})",
                    expected_bytes.len(),
                    actual_bytes.len(),
                    first_diff_offset(expected_bytes, actual_bytes)
                        .map(|idx| idx.to_string())
                        .unwrap_or_else(|| "unknown".to_string())
                );
                if part.starts_with("xl/charts/") && part.ends_with(".xml") {
                    if let (Ok(expected_xml), Ok(actual_xml)) = (
                        xlsx_diff::NormalizedXml::parse(part, expected_bytes),
                        xlsx_diff::NormalizedXml::parse(part, actual_bytes),
                    ) {
                        if expected_xml == actual_xml {
                            msg.push_str(" (normalized_xml_equal)");
                        }
                    }
                }
                issues.entry(part.to_string()).or_default().push(msg);
            }
            None => {
                issues
                    .entry(part.to_string())
                    .or_default()
                    .push("chart_part_missing".to_string());
            }
        }
    }

    if !issues.is_empty() {
        eprintln!(
            "Chart OPC parts changed during no-op round-trip for fixture {}",
            fixture.display()
        );
        print_grouped_report(&issues, &report);
        panic!("fixture {} did not preserve chart parts", fixture.display());
    }

    Ok(())
}

fn is_chart_related_part(part: &str) -> bool {
    CHART_PART_PREFIXES
        .iter()
        .any(|prefix| part.starts_with(prefix))
}

fn first_diff_offset(a: &[u8], b: &[u8]) -> Option<usize> {
    let shared = a.len().min(b.len());
    for idx in 0..shared {
        if a[idx] != b[idx] {
            return Some(idx);
        }
    }
    (a.len() != b.len()).then_some(shared)
}

fn print_grouped_report(issues: &BTreeMap<String, Vec<String>>, report: &xlsx_diff::DiffReport) {
    let mut diffs_by_part: BTreeMap<&str, Vec<&xlsx_diff::Difference>> = BTreeMap::new();
    for diff in &report.differences {
        if is_chart_related_part(&diff.part) || diff.part.ends_with(".rels") {
            diffs_by_part.entry(&diff.part).or_default().push(diff);
        }
    }

    let mut parts: BTreeSet<&str> = BTreeSet::new();
    parts.extend(issues.keys().map(String::as_str));
    parts.extend(diffs_by_part.keys().copied());

    for part in parts {
        eprintln!("- {part}");
        if let Some(entries) = issues.get(part) {
            for entry in entries {
                eprintln!("  issue: {entry}");
            }
        }
        if let Some(diffs) = diffs_by_part.get(part) {
            for diff in diffs {
                eprintln!(
                    "  diff: [{}] {}{}{}",
                    diff.severity,
                    diff.kind,
                    if diff.path.is_empty() { "" } else { " " },
                    diff.path
                );
                if diff.expected.is_some() || diff.actual.is_some() {
                    eprintln!(
                        "    expected: {}",
                        diff.expected.as_deref().unwrap_or("<none>")
                    );
                    eprintln!(
                        "    actual:   {}",
                        diff.actual.as_deref().unwrap_or("<none>")
                    );
                }
            }
        }
    }
}

#[test]
fn chart_roundtrip_preserves_chart_xml_with_mc_alternate_content_and_extlst(
) -> Result<(), Box<dyn Error>> {
    let base_xlsx = build_simple_chart_xlsx();
    let patched_xlsx = patch_xlsx_chart1_xml(&base_xlsx)?;

    let expected_chart_xml = zip_part(&patched_xlsx, "xl/charts/chart1.xml")?;
    let chart_xml_str =
        std::str::from_utf8(&expected_chart_xml).expect("chart1.xml should be UTF-8");
    assert!(
        chart_xml_str.contains("mc:AlternateContent"),
        "fixture chart1.xml should include mc:AlternateContent"
    );
    assert!(
        chart_xml_str.contains("<c:extLst"),
        "fixture chart1.xml should include a c:extLst"
    );

    let doc = load_from_bytes(&patched_xlsx)?;
    let roundtripped_bytes = doc.save_to_vec()?;
    let actual_chart_xml = zip_part(&roundtripped_bytes, "xl/charts/chart1.xml")?;

    assert_eq!(
        actual_chart_xml, expected_chart_xml,
        "xl/charts/chart1.xml bytes changed during no-op round-trip"
    );

    Ok(())
}

fn build_simple_chart_xlsx() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Category").unwrap();
    worksheet.write_string(0, 1, "Value").unwrap();

    let categories = ["A", "B", "C", "D"];
    let values = [2.0, 4.0, 3.0, 5.0];

    for (i, (cat, val)) in categories.iter().zip(values).enumerate() {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, *cat).unwrap();
        worksheet.write_number(row, 1, val).unwrap();
    }

    let mut chart = Chart::new(XlsxChartType::Column);
    chart.title().set_name("Example Chart");

    let series = chart.add_series();
    series
        .set_categories("Sheet1!$A$2:$A$5")
        .set_values("Sheet1!$B$2:$B$5");

    worksheet.insert_chart(1, 3, &chart).unwrap();

    workbook.save_to_buffer().unwrap()
}

fn patch_xlsx_chart1_xml(xlsx_bytes: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    let cursor = Cursor::new(xlsx_bytes);
    let mut archive = ZipArchive::new(cursor)?;

    let out_cursor = Cursor::new(Vec::new());
    let mut out_zip = ZipWriter::new(out_cursor);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();
        let options = FileOptions::<()>::default().compression_method(file.compression());

        if file.is_dir() {
            out_zip.add_directory(name, options)?;
            continue;
        }

        out_zip.start_file(name.clone(), options)?;
        if name.trim_start_matches('/') == "xl/charts/chart1.xml" {
            let mut buf = Vec::new();
            file.read_to_end(&mut buf)?;
            let patched = patch_chart_xml(&buf)?;
            out_zip.write_all(&patched)?;
        } else {
            std::io::copy(&mut file, &mut out_zip)?;
        }
    }

    Ok(out_zip.finish()?.into_inner())
}

fn patch_chart_xml(chart_xml: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    const BAR_CHART_OPEN: &str = "<c:barChart";
    const BAR_CHART_CLOSE: &str = "</c:barChart>";

    let xml = std::str::from_utf8(chart_xml)?;
    let start = xml
        .find(BAR_CHART_OPEN)
        .ok_or("chart1.xml missing <c:barChart>")?;
    let end_rel = xml[start..]
        .find(BAR_CHART_CLOSE)
        .ok_or("chart1.xml missing </c:barChart>")?;
    let end = start + end_rel + BAR_CHART_CLOSE.len();

    let bar_chart = &xml[start..end];

    let bar_chart_with_ext = if bar_chart.contains("<c:extLst") {
        bar_chart.to_string()
    } else {
        let insert_at = bar_chart
            .rfind(BAR_CHART_CLOSE)
            .ok_or("chart1.xml missing </c:barChart>")?;
        let mut out = String::with_capacity(bar_chart.len() + 128);
        out.push_str(&bar_chart[..insert_at]);
        out.push_str(
            r#"<c:extLst><c:ext uri="{77B8C3E4-5F7E-4BCE-9C65-FF0F0F0F0F0F}"><fx:dummy xmlns:fx="urn:formula-xlsx:test">1</fx:dummy></c:ext></c:extLst>"#,
        );
        out.push_str(BAR_CHART_CLOSE);
        out
    };

    let wrapped = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14">{bar}</mc:Choice><mc:Fallback>{bar}</mc:Fallback></mc:AlternateContent>"#,
        bar = bar_chart_with_ext
    );

    let mut out = String::with_capacity(xml.len() + wrapped.len() + 32);
    out.push_str(&xml[..start]);
    out.push_str(&wrapped);
    out.push_str(&xml[end..]);
    Ok(out.into_bytes())
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor)?;

    let mut buf = Vec::new();
    if let Ok(mut file) = archive.by_name(name) {
        file.read_to_end(&mut buf)?;
        return Ok(buf);
    }

    let mut file = archive.by_name(&format!("/{name}"))?;
    file.read_to_end(&mut buf)?;
    Ok(buf)
}
