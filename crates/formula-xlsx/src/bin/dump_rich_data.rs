//! Developer CLI to inspect Excel "rich data" (images-in-cell / rich values) wiring.
//!
//! Usage:
//!   cargo run -p formula-xlsx --bin dump_rich_data -- <path-to-xlsx>
//!
//! Output (one line per cell with `vm="..."`):
//!   <sheet>!<cell> vm=<vm> -> rv=<rich_value_index> -> <xl/media/* path> rel=<rel_index>
//!
//! Any field that cannot be resolved is printed as `-`.

#[cfg(not(target_arch = "wasm32"))]
use std::collections::{BTreeMap, HashMap};
#[cfg(not(target_arch = "wasm32"))]
use std::error::Error;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use formula_model::{CellRef, HyperlinkTarget};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::rich_data::RichDataVmIndex;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::XlsxPackage;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::events::Event;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::Reader;

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn usage() -> &'static str {
    "dump_rich_data <path.xlsx> [--print-parts] [--extract-cell-images] [--extract-cell-images-out <dir>]\n\
\n\
Print a best-effort mapping from worksheet cells with `vm` attributes to:\n\
  - xl/metadata.xml valueMetadata indices (vm)\n\
  - xl/richData/richValue*.xml indices (rv)\n\
  - xl/richData/richValueRel*.xml relationship-table indices (rel)\n\
  - resolved xl/media/* targets\n\
\n\
Output format (one line per cell):\n\
  <sheet>!<cell> vm=<vm> -> rv=<rich_value_index> -> <xl/media/* path> rel=<rel_index>\n\
\n\
Options:\n\
  --print-parts\n\
      Print a list of richData-related ZIP parts (to stderr).\n\
  --extract-cell-images\n\
      Extract rich-data in-cell images by cell and print a summary (to stderr).\n\
  --extract-cell-images-out <dir>\n\
      Write extracted in-cell images to <dir> as files.\n"
}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let mut xlsx_path: Option<PathBuf> = None;
    let mut print_parts = false;
    let mut extract_cell_images = false;
    let mut extract_cell_images_out: Option<PathBuf> = None;

    let mut args = std::env::args().skip(1).peekable();
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--help" | "-h" => {
                eprintln!("{}", usage());
                return Ok(());
            }
            "--print-parts" => {
                print_parts = true;
            }
            "--extract-cell-images" => {
                extract_cell_images = true;
            }
            "--extract-cell-images-out" => {
                let Some(dir) = args.next() else {
                    return Err(format!(
                        "missing <dir> for --extract-cell-images-out\n\n{}",
                        usage()
                    )
                    .into());
                };
                if dir.is_empty() {
                    return Err(format!(
                        "missing <dir> for --extract-cell-images-out\n\n{}",
                        usage()
                    )
                    .into());
                }
                extract_cell_images = true;
                extract_cell_images_out = Some(PathBuf::from(dir));
            }
            flag if flag.starts_with("--extract-cell-images-out=") => {
                let Some((_, dir)) = flag.split_once('=') else {
                    unreachable!();
                };
                if dir.is_empty() {
                    return Err(format!(
                        "missing <dir> for --extract-cell-images-out\n\n{}",
                        usage()
                    )
                    .into());
                }
                extract_cell_images = true;
                extract_cell_images_out = Some(PathBuf::from(dir));
            }
            flag if flag.starts_with('-') => {
                return Err(format!("unknown flag: {flag}\n\n{}", usage()).into());
            }
            value => {
                if xlsx_path.is_some() {
                    return Err(format!("unexpected extra argument: {value}\n\n{}", usage()).into());
                }
                xlsx_path = Some(PathBuf::from(value));
            }
        }
    }

    let xlsx_path = xlsx_path.ok_or_else(|| format!("missing <path.xlsx>\n\n{}", usage()))?;

    let bytes = fs::read(&xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    let found_vm_cells = dump_vm_mappings(&pkg);

    // Optional debug output.
    if print_parts {
        dump_rich_data_parts(&pkg);
    }
    if extract_cell_images {
        dump_rich_cell_images_by_cell(&pkg, extract_cell_images_out.as_deref());
    }

    // For the common "no richData" case, print a single line and exit 0.
    if !found_vm_cells {
        // `dump_vm_mappings` already printed the message.
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone)]
struct SheetInfo {
    sheet_index: usize,
    sheet_name: String,
    worksheet_part: String,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone)]
struct VmCellEntry {
    sheet_index: usize,
    sheet_name: String,
    cell_ref: String,
    cell_ref_parsed: Option<CellRef>,
    vm: String,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone)]
struct MappingRow {
    sheet_index: usize,
    sheet_name: String,
    cell_ref: String,
    cell_ref_parsed: Option<CellRef>,
    vm: String,
    rich_value_index: Option<u32>,
    rel_index: Option<u32>,
    media_part: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_vm_mappings(pkg: &XlsxPackage) -> bool {
    let sheets = list_sheets_best_effort(pkg);
    let vm_cells = collect_vm_cells(pkg, &sheets);

    if vm_cells.is_empty() {
        println!("no richData found");
        return false;
    }

    let index = match RichDataVmIndex::build(pkg) {
        Ok(index) => index,
        Err(err) => {
            eprintln!("warning: failed to build richData index: {err}");
            RichDataVmIndex::default()
        }
    };

    let mut rows: Vec<MappingRow> = Vec::with_capacity(vm_cells.len());
    for cell in vm_cells {
        let resolution = cell
            .vm
            .parse::<u32>()
            .ok()
            .map(|vm| index.resolve_vm(vm))
            .unwrap_or_default();

        rows.push(MappingRow {
            sheet_index: cell.sheet_index,
            sheet_name: cell.sheet_name,
            cell_ref: cell.cell_ref,
            cell_ref_parsed: cell.cell_ref_parsed,
            vm: cell.vm,
            rich_value_index: resolution.rich_value_index,
            rel_index: resolution.rel_index,
            media_part: resolution.target_part,
        });
    }

    // Stable output: sheet order (workbook order), then cell (row/col), then raw cell string.
    rows.sort_by(|a, b| {
        let sheet_cmp = a.sheet_index.cmp(&b.sheet_index);
        if sheet_cmp != std::cmp::Ordering::Equal {
            return sheet_cmp;
        }
        match (&a.cell_ref_parsed, &b.cell_ref_parsed) {
            (Some(a_ref), Some(b_ref)) => {
                let rc_cmp = (a_ref.row, a_ref.col).cmp(&(b_ref.row, b_ref.col));
                if rc_cmp != std::cmp::Ordering::Equal {
                    return rc_cmp;
                }
            }
            (Some(_), None) => return std::cmp::Ordering::Less,
            (None, Some(_)) => return std::cmp::Ordering::Greater,
            (None, None) => {}
        }
        a.cell_ref.cmp(&b.cell_ref)
    });

    for row in rows {
        let rich_value_index = row
            .rich_value_index
            .map(|n| n.to_string())
            .unwrap_or_else(|| "-".to_string());
        let rel_index = row
            .rel_index
            .map(|n| n.to_string())
            .unwrap_or_else(|| "-".to_string());
        let media_part = row.media_part.unwrap_or_else(|| "-".to_string());

        // Keep the `... -> <target>` segment stable for substring checks in tests/log-scraping.
        println!(
            "{}!{} vm={} -> rv={} -> {} rel={}",
            row.sheet_name, row.cell_ref, row.vm, rich_value_index, media_part, rel_index
        );
    }

    true
}

#[cfg(not(target_arch = "wasm32"))]
fn list_sheets_best_effort(pkg: &XlsxPackage) -> Vec<SheetInfo> {
    match pkg.worksheet_parts() {
        Ok(parts) => parts
            .into_iter()
            .enumerate()
            .map(|(idx, part)| SheetInfo {
                sheet_index: idx,
                sheet_name: part.name,
                worksheet_part: part.worksheet_part,
            })
            .collect(),
        Err(err) => {
            eprintln!(
                "warning: failed to read workbook sheet list ({err}); falling back to xl/worksheets/*"
            );
            let mut worksheet_parts: Vec<String> = pkg
                .part_names()
                .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
                .map(|s| s.to_string())
                .collect();
            worksheet_parts.sort();
            worksheet_parts
                .into_iter()
                .enumerate()
                .map(|(idx, part)| SheetInfo {
                    sheet_index: idx,
                    sheet_name: part.clone(),
                    worksheet_part: part,
                })
                .collect()
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn collect_vm_cells(pkg: &XlsxPackage, sheets: &[SheetInfo]) -> Vec<VmCellEntry> {
    let mut out = Vec::new();
    for sheet in sheets {
        let Some(xml) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };
        for (cell_ref, vm) in parse_vm_cells_best_effort(xml) {
            let cell_ref_parsed = CellRef::from_a1(&cell_ref).ok();
            out.push(VmCellEntry {
                sheet_index: sheet.sheet_index,
                sheet_name: sheet.sheet_name.clone(),
                cell_ref,
                cell_ref_parsed,
                vm,
            });
        }
    }
    out
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_vm_cells_best_effort(worksheet_xml: &[u8]) -> Vec<(String, String)> {
    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut in_sheet_data = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true;
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if in_sheet_data && local_name(e.name().as_ref()) == b"c" =>
            {
                let mut r: Option<String> = None;
                let mut vm: Option<String> = None;
                for attr in e.attributes().with_checks(false) {
                    let Ok(attr) = attr else {
                        continue;
                    };
                    let Ok(v) = attr.unescape_value() else {
                        continue;
                    };
                    match local_name(attr.key.as_ref()) {
                        b"r" => r = Some(v.into_owned()),
                        b"vm" => vm = Some(v.into_owned()),
                        _ => {}
                    }
                }

                if let (Some(r), Some(vm)) = (r, vm) {
                    out.push((r, vm));
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }

        buf.clear();
    }

    out
}

#[cfg(not(target_arch = "wasm32"))]
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(idx) => &name[(idx + 1)..],
        None => name,
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_rich_data_parts(pkg: &XlsxPackage) {
    let mut parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| {
            *name == "xl/metadata.xml"
                || name.starts_with("xl/richData/")
                || name.starts_with("xl/richData/_rels/")
        })
        .collect();
    parts.sort();

    eprintln!("richData parts ({}):", parts.len());
    for part in parts {
        eprintln!("  {part}");
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_rich_cell_images_by_cell(pkg: &XlsxPackage, out_dir: Option<&Path>) {
    let images_by_cell = match pkg.extract_embedded_cell_images() {
        Ok(v) => v,
        Err(err) => {
            eprintln!("rich-data in-cell images (by cell): failed to extract: {err}");
            return;
        }
    };

    if images_by_cell.is_empty() {
        eprintln!("rich-data in-cell images (by cell): (none)");
        return;
    }

    let sheet_name_by_part: HashMap<String, String> = list_sheets_best_effort(pkg)
        .into_iter()
        .map(|sheet| (sheet.worksheet_part, sheet.sheet_name))
        .collect();

    eprintln!(
        "rich-data in-cell images (by cell): {} cell(s)",
        images_by_cell.len()
    );

    let mut counts_by_sheet: BTreeMap<String, usize> = BTreeMap::new();
    for ((worksheet_part, _cell), _image) in &images_by_cell {
        let sheet = sheet_name_by_part
            .get(worksheet_part)
            .unwrap_or(worksheet_part);
        *counts_by_sheet.entry(sheet.clone()).or_insert(0) += 1;
    }
    if !counts_by_sheet.is_empty() {
        eprintln!("  sheets:");
        for (sheet, count) in counts_by_sheet {
            eprintln!("    {sheet}: {count}");
        }
    }

    if let Some(out_dir) = out_dir {
        if let Err(err) = fs::create_dir_all(out_dir) {
            eprintln!(
                "rich-data in-cell images (by cell): failed to create output dir {}: {err}",
                out_dir.display()
            );
            return;
        }

        let mut manifest: Vec<String> = Vec::with_capacity(images_by_cell.len() + 1);
        manifest.push(
            "sheet\tcell\tbytes\tfile\timage_part\tcalc_origin\talt_text\thyperlink".to_string(),
        );

        let mut written = 0usize;
        let mut failed = 0usize;
        let mut printed_failures = 0usize;
        let max_printed_failures = 10usize;

        for ((worksheet_part, cell), image) in &images_by_cell {
            let sheet = sheet_name_by_part
                .get(worksheet_part)
                .unwrap_or(worksheet_part);
            let sheet_sanitized = sanitize_filename_component(sheet);
            let cell_a1 = cell.to_string();
            let bytes = &image.image_bytes;
            let ext = guess_image_extension(bytes).unwrap_or("bin");

            let mut file_name = format!("{sheet_sanitized}_{cell_a1}.{ext}");
            let mut path = out_dir.join(&file_name);
            if path.exists() {
                for suffix in 1u32.. {
                    file_name = format!("{sheet_sanitized}_{cell_a1}_{suffix}.{ext}");
                    path = out_dir.join(&file_name);
                    if !path.exists() {
                        break;
                    }
                }
            }

            match fs::write(&path, bytes) {
                Ok(()) => {
                    written += 1;
                    let hyperlink = image
                        .hyperlink_target
                        .as_ref()
                        .map(format_hyperlink_target)
                        .unwrap_or_else(|| "-".to_string());
                    manifest.push(format!(
                        "{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
                        tsv_escape(sheet),
                        cell_a1,
                        bytes.len(),
                        file_name,
                        tsv_escape(&image.image_part),
                        image.calc_origin,
                        tsv_escape(image.alt_text.as_deref().unwrap_or("-")),
                        tsv_escape(&hyperlink),
                    ));
                }
                Err(err) => {
                    failed += 1;
                    if printed_failures < max_printed_failures {
                        eprintln!("warning: failed to write {}: {err}", path.display());
                        printed_failures += 1;
                    }
                }
            }
        }

        let manifest_path = out_dir.join("manifest.tsv");
        if let Err(err) = fs::write(&manifest_path, manifest.join("\n") + "\n") {
            eprintln!(
                "warning: failed to write manifest {}: {err}",
                manifest_path.display()
            );
        }

        eprintln!(
            "rich-data in-cell images (by cell): wrote {written} file(s) to {} (failed: {failed}; manifest: {})",
            out_dir.display(),
            manifest_path.display()
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn sanitize_filename_component(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | '.') {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    let trimmed = out.trim_matches('_');
    if trimmed.is_empty() {
        "sheet".to_string()
    } else {
        trimmed.to_string()
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn guess_image_extension(bytes: &[u8]) -> Option<&'static str> {
    const PNG_MAGIC: &[u8] = b"\x89PNG\r\n\x1a\n";
    if bytes.starts_with(PNG_MAGIC) {
        return Some("png");
    }
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        return Some("jpg");
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        return Some("gif");
    }
    if bytes.starts_with(b"BM") {
        return Some("bmp");
    }
    if bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*") {
        return Some("tiff");
    }
    // Windows Metafile (placeable WMF) signature.
    if bytes.starts_with(b"\xD7\xCD\xC6\x9A") {
        return Some("wmf");
    }
    // Enhanced Metafile: ENHMETAHEADER signature at offset 0x28 (40 bytes): 0x464D4520 (" EMF").
    if bytes.len() >= 44 && &bytes[40..44] == b" EMF" {
        return Some("emf");
    }
    // SVG (common for Office 365 / modern sources).
    let prefix = bytes.get(..512).unwrap_or(bytes);
    if let Ok(s) = std::str::from_utf8(prefix) {
        let lower = s.to_ascii_lowercase();
        if lower.contains("<svg") {
            return Some("svg");
        }
    }
    None
}

#[cfg(not(target_arch = "wasm32"))]
fn format_hyperlink_target(target: &HyperlinkTarget) -> String {
    match target {
        HyperlinkTarget::ExternalUrl { uri } => uri.clone(),
        HyperlinkTarget::Email { uri } => uri.clone(),
        HyperlinkTarget::Internal { sheet, cell } => format!("{sheet}!{cell}"),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn tsv_escape(value: &str) -> String {
    // Sheet names are user-controlled; keep the manifest format robust.
    value.replace(['\t', '\r', '\n'], " ")
}
