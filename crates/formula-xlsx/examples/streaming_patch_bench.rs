use std::fs;
use std::io::{Write};
use std::path::Path;
use std::time::Instant;

use formula_model::{Cell, CellRef, CellValue};
use formula_xlsx::{load_from_bytes, patch_xlsx_streaming, WorksheetCellPatch};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let rows: u32 = std::env::args()
        .nth(1)
        .and_then(|s| s.parse().ok())
        .unwrap_or(100_000);

    let tmpdir = tempfile::tempdir()?;
    let input_path = tmpdir.path().join("synthetic.xlsx");
    let streaming_out = tmpdir.path().join("streaming.xlsx");
    let inmem_out = tmpdir.path().join("in_memory.xlsx");

    write_synthetic_workbook(&input_path, rows)?;
    println!("synthetic workbook written to {}", input_path.display());

    let patch_cell = CellRef::new(rows / 2, 0); // A{rows/2 + 1}
    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        patch_cell,
        CellValue::Number(123.0),
        Some("=1+2".to_string()),
    );

    println!("\n== Streaming patch ==");
    let rss_before = rss_kb();
    let start = Instant::now();
    patch_xlsx_streaming(
        fs::File::open(&input_path)?,
        fs::File::create(&streaming_out)?,
        &[patch.clone()],
    )?;
    let elapsed = start.elapsed();
    let rss_after = rss_kb();
    println!("wall time: {:.2?}", elapsed);
    print_rss_delta(rss_before, rss_after);

    println!("\n== In-memory load + rewrite ==");
    let rss_before = rss_kb();
    let start = Instant::now();
    let bytes = fs::read(&input_path)?;
    let rss_after_read = rss_kb();
    let mut doc = load_from_bytes(&bytes)?;
    let rss_after_load = rss_kb();

    // Apply the same patch by mutating the in-memory workbook model.
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc
        .workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists in synthetic workbook");
    let mut cell = sheet.cell(patch_cell).cloned().unwrap_or_else(Cell::default);
    cell.value = CellValue::Number(123.0);
    cell.formula = Some("1+2".to_string());
    sheet.set_cell(patch_cell, cell);

    let out_bytes = doc.save_to_vec()?;
    let rss_after_save = rss_kb();
    fs::write(&inmem_out, &out_bytes)?;
    let elapsed = start.elapsed();
    println!("wall time: {:.2?}", elapsed);
    println!("after read:  {}", format_rss(rss_after_read));
    println!("after load:  {}", format_rss(rss_after_load));
    println!("after save:  {}", format_rss(rss_after_save));
    print_rss_delta(rss_before, rss_after_save);

    println!("\noutputs:");
    println!("  streaming: {}", streaming_out.display());
    println!("  in-memory: {}", inmem_out.display());

    Ok(())
}

fn write_synthetic_workbook(path: &Path, rows: u32) -> Result<(), Box<dyn std::error::Error>> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(content_types_xml().as_bytes())?;

    zip.start_file("_rels/.rels", options)?;
    zip.write_all(root_rels_xml().as_bytes())?;

    zip.start_file("xl/workbook.xml", options)?;
    zip.write_all(workbook_xml().as_bytes())?;

    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    zip.write_all(workbook_rels_xml().as_bytes())?;

    zip.start_file("xl/styles.xml", options)?;
    zip.write_all(minimal_styles_xml().as_bytes())?;

    zip.start_file("xl/worksheets/sheet1.xml", options)?;
    write_sheet1_xml(&mut zip, rows)?;

    zip.finish()?;
    Ok(())
}

fn content_types_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
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

fn write_sheet1_xml<W: Write>(mut w: W, rows: u32) -> Result<(), Box<dyn std::error::Error>> {
    write!(
        w,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A{rows}"/>
  <sheetData>"#
    )?;

    for r in 1..=rows {
        write!(w, r#"<row r="{r}"><c r="A{r}"><v>{r}</v></c></row>"#)?;
    }

    write!(
        w,
        r#"</sheetData>
</worksheet>
"#
    )?;

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

