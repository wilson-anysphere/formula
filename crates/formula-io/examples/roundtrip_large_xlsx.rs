//! Perf/memory regression harness for the lazy/streaming XLSX package pipeline.
//!
//! This example generates a synthetic `.xlsx` with:
//! - A *very large* `xl/worksheets/sheet1.xml` (highly compressible, but large when inflated).
//! - A large incompressible binary part (`xl/media/image1.png`) so the ZIP on disk is ~100MB.
//!
//! It then runs `formula_io::open_workbook` followed by `formula_io::save_workbook` and prints
//! wall time and RSS deltas (Linux only).
//!
//! Run (recommended with `--release`):
//!   cargo run -p formula-io --example roundtrip_large_xlsx --release -- [rows] [image-mb] [cell-bytes]
//!
//! Example (≈100MB `.xlsx` on disk, hundreds of MB inflated worksheet XML):
//!   cargo run -p formula-io --example roundtrip_large_xlsx --release -- 100000 96 2048
//!
//! Notes:
//! - RSS is reported via `/proc/self/status` and is only available on Linux.
//! - This is intentionally "synthetic" but is useful to validate the <500MB memory target.
//! - The temp directory is deleted on exit (paths are printed for convenience).
use std::fs;
use std::io::Write;
use std::path::Path;
use std::time::Instant;

use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let rows: u32 = args
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or(100_000);
    let image_mb: usize = args
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or(96);
    let cell_bytes: usize = args
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or(2048);

    let tmpdir = tempfile::tempdir()?;
    let input_path = tmpdir.path().join("synthetic_large.xlsx");
    let output_path = tmpdir.path().join("roundtrip.xlsx");

    let image_bytes = image_mb * 1024 * 1024;
    let payload = "A".repeat(cell_bytes);

    println!("temp dir: {}", tmpdir.path().display());
    println!("config:");
    println!("  rows:              {rows}");
    println!("  cell payload:      {cell_bytes} bytes (repeated 'A')");
    println!("  media/image1.png:  {image_mb} MiB pseudorandom bytes");

    let start_gen = Instant::now();
    write_synthetic_workbook(&input_path, rows, &payload, image_bytes)?;
    let gen_elapsed = start_gen.elapsed();
    let input_size = fs::metadata(&input_path)?.len();
    println!(
        "\nsynthetic workbook written: {} ({:.2} MiB) in {:.2?}",
        input_path.display(),
        input_size as f64 / (1024.0 * 1024.0),
        gen_elapsed
    );

    println!("\n== formula_io::open_workbook ==");
    let rss_before_open = rss_kb();
    let start_open = Instant::now();
    let workbook = formula_io::open_workbook(&input_path)?;
    let open_elapsed = start_open.elapsed();
    let rss_after_open = rss_kb();
    println!("wall time: {:.2?}", open_elapsed);
    print_rss_delta(rss_before_open, rss_after_open);

    println!("\n== formula_io::save_workbook ==");
    let rss_before_save = rss_kb();
    let start_save = Instant::now();
    formula_io::save_workbook(&workbook, &output_path)?;
    let save_elapsed = start_save.elapsed();
    let rss_after_save = rss_kb();
    println!("wall time: {:.2?}", save_elapsed);
    print_rss_delta(rss_before_save, rss_after_save);

    let output_size = fs::metadata(&output_path)?.len();
    println!(
        "\noutput written: {} ({:.2} MiB)",
        output_path.display(),
        output_size as f64 / (1024.0 * 1024.0)
    );

    println!("\n== Total RSS delta (before open -> after save) ==");
    print_rss_delta(rss_before_open, rss_after_save);

    Ok(())
}

fn write_synthetic_workbook(
    path: &Path,
    rows: u32,
    cell_payload: &str,
    image_bytes: usize,
) -> Result<(), Box<dyn std::error::Error>> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);

    let xml_options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
    let binary_options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip.start_file("[Content_Types].xml", xml_options)?;
    zip.write_all(content_types_xml().as_bytes())?;

    zip.start_file("_rels/.rels", xml_options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    zip.start_file("xl/workbook.xml", xml_options)?;
    zip.write_all(workbook_xml().as_bytes())?;

    zip.start_file("xl/_rels/workbook.xml.rels", xml_options)?;
    zip.write_all(workbook_rels_xml().as_bytes())?;

    zip.start_file("xl/styles.xml", xml_options)?;
    zip.write_all(minimal_styles_xml().as_bytes())?;

    zip.start_file("xl/worksheets/sheet1.xml", xml_options)?;
    write_sheet1_xml(&mut zip, rows, cell_payload)?;

    zip.start_file("xl/media/image1.png", binary_options)?;
    write_pseudorandom_bytes(&mut zip, image_bytes)?;

    zip.finish()?;
    Ok(())
}

fn content_types_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"#
}

fn root_rels_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
}

fn workbook_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
}

fn workbook_rels_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"#
}

fn minimal_styles_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"#
}

fn write_sheet1_xml<W: Write>(
    mut w: W,
    rows: u32,
    cell_payload: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    write!(
        w,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A{rows}"/>
  <sheetData>"#
    )?;

    for r in 1..=rows {
        write!(
            w,
            r#"<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>"#
        )?;
        w.write_all(cell_payload.as_bytes())?;
        write!(w, r#"</t></is></c></row>"#)?;
    }

    write!(
        w,
        r#"</sheetData>
</worksheet>
"#
    )?;

    Ok(())
}

fn splitmix64(state: &mut u64) -> u64 {
    // https://prng.di.unimi.it/splitmix64.c
    *state = state.wrapping_add(0x9E3779B97F4A7C15);
    let mut z = *state;
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58476D1CE4E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D049BB133111EB);
    z ^ (z >> 31)
}

fn write_pseudorandom_bytes<W: Write>(mut w: W, bytes: usize) -> Result<(), std::io::Error> {
    let mut remaining = bytes;
    let mut state = 0x1234_5678_9ABC_DEF0u64;
    let mut buf = vec![0u8; 64 * 1024];

    while remaining > 0 {
        let chunk_len = remaining.min(buf.len());
        let mut i = 0usize;
        while i + 8 <= chunk_len {
            let v = splitmix64(&mut state).to_le_bytes();
            buf[i..i + 8].copy_from_slice(&v);
            i += 8;
        }
        if i < chunk_len {
            let v = splitmix64(&mut state).to_le_bytes();
            buf[i..chunk_len].copy_from_slice(&v[..chunk_len - i]);
        }
        w.write_all(&buf[..chunk_len])?;
        remaining -= chunk_len;
    }

    Ok(())
}

#[cfg(target_os = "linux")]
fn rss_kb() -> Option<u64> {
    let status = fs::read_to_string("/proc/self/status").ok()?;
    for line in status.lines() {
        if let Some(rest) = line.strip_prefix("VmRSS:") {
            let kb = rest
                .split_whitespace()
                .next()
                .and_then(|v| v.parse::<u64>().ok())?;
            return Some(kb);
        }
    }
    None
}

#[cfg(not(target_os = "linux"))]
fn rss_kb() -> Option<u64> {
    None
}

fn format_rss(rss: Option<u64>) -> String {
    match rss {
        Some(kb) => format!("{kb} kB"),
        None => "n/a".to_string(),
    }
}

fn print_rss_delta(before: Option<u64>, after: Option<u64>) {
    println!("rss before: {}", format_rss(before));
    println!("rss after:  {}", format_rss(after));
    match (before, after) {
        (Some(b), Some(a)) if a >= b => println!("rss delta:  {} kB", a - b),
        (Some(b), Some(a)) => println!("rss delta:  -{} kB", b - a),
        _ => {}
    }
}

