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
use formula_xlsx::openxml;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::rich_data::RichDataVmIndex;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::XlsxPackage;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::events::Event;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::Reader;

#[cfg(not(target_arch = "wasm32"))]
mod workbook_open;

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn usage() -> &'static str {
    "dump_rich_data [--password <pw>] <path.xlsx> [--print-parts] [--extract-cell-images] [--extract-cell-images-out <dir>]\n\
\n\
Print a best-effort overview of rich-data related parts:\n\
  - xl/metadata.xml presence/size\n\
  - xl/richData/* part list + sizes\n\
  - likely in-cell image parts (xl/cellImages*, xl/media/*, etc)\n\
  - workbook.xml.rels and metadata.xml.rels entries involving metadata/richData\n\
  - per-worksheet vm/cm usage + a few sample cells\n\
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
  --password <pw>\n\
      Password for Office-encrypted workbooks (OLE `EncryptedPackage`; use --password '' for empty password).\n\
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
    let mut password: Option<String> = None;
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
            "--password" => {
                let Some(pw) = args.next() else {
                    return Err(format!("missing <pw> for --password\n\n{}", usage()).into());
                };
                if pw.starts_with('-') {
                    return Err(format!("missing <pw> for --password\n\n{}", usage()).into());
                }
                password = Some(pw);
            }
            flag if flag.starts_with("--password=") => {
                let Some((_, pw)) = flag.split_once('=') else {
                    unreachable!();
                };
                password = Some(pw.to_string());
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

    let pkg = workbook_open::open_xlsx_package(&xlsx_path, password.as_deref())?;

    eprintln!("workbook: {}", xlsx_path.display());
    dump_required_part_presence(&pkg);
    dump_workbook_relationships(&pkg);
    dump_metadata_relationships(&pkg);
    dump_worksheet_vm_cm_usage(&pkg);

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
fn dump_required_part_presence(pkg: &XlsxPackage) {
    eprintln!();
    eprintln!("parts:");

    match find_part_bytes_case_insensitive(pkg, "xl/metadata.xml") {
        Some((name, bytes)) => eprintln!("  {name}: present ({} bytes)", bytes.len()),
        None => eprintln!("  xl/metadata.xml: missing"),
    };

    let mut rich_data_parts: Vec<(&str, usize)> = pkg
        .part_names()
        .filter_map(|name| {
            let canonical = name.strip_prefix('/').unwrap_or(name);
            let lower = canonical.to_ascii_lowercase();
            if lower.starts_with("xl/richdata/") {
                Some((canonical, pkg.part(name).map(|b| b.len()).unwrap_or(0)))
            } else {
                None
            }
        })
        .collect();
    rich_data_parts.sort_by(|a, b| a.0.cmp(&b.0));
    if rich_data_parts.is_empty() {
        eprintln!("  xl/richData/: (none)");
    } else {
        eprintln!("  xl/richData/: {} part(s)", rich_data_parts.len());
        for (name, len) in rich_data_parts {
            eprintln!("    {name} ({len} bytes)");
        }
    }

    let mut image_parts: Vec<(&str, usize)> = pkg
        .part_names()
        .filter_map(|name| {
            let canonical = name.strip_prefix('/').unwrap_or(name);
            if is_in_cell_image_candidate_part(canonical) {
                Some((canonical, pkg.part(name).map(|b| b.len()).unwrap_or(0)))
            } else {
                None
            }
        })
        .collect();
    image_parts.sort_by(|a, b| a.0.cmp(&b.0));
    if image_parts.is_empty() {
        eprintln!("  in-cell image candidates: (none)");
    } else {
        eprintln!("  in-cell image candidates: {} part(s)", image_parts.len());
        for (name, len) in image_parts {
            eprintln!("    {name} ({len} bytes)");
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn is_in_cell_image_candidate_part(part_name: &str) -> bool {
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
    let lower = part_name.to_ascii_lowercase();

    // Common/known in-cell image parts.
    if lower.starts_with("xl/cellimages") || lower.starts_with("xl/_rels/cellimages") {
        return true;
    }

    // Image payloads.
    if lower.starts_with("xl/media/") {
        return true;
    }

    // Fallback heuristic: any part whose name suggests cell images.
    lower.contains("cellimage")
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_workbook_relationships(pkg: &XlsxPackage) {
    eprintln!();
    eprintln!("workbook.xml.rels (filtered metadata/richData):");

    let Some((rels_part_name, rels_bytes)) =
        find_part_bytes_case_insensitive(pkg, "xl/_rels/workbook.xml.rels")
    else {
        eprintln!("  (missing)");
        return;
    };
    eprintln!("  part: {rels_part_name} ({} bytes)", rels_bytes.len());

    let relationships = match openxml::parse_relationships(rels_bytes) {
        Ok(r) => r,
        Err(err) => {
            eprintln!("  warning: failed to parse workbook.xml.rels: {err}");
            return;
        }
    };

    let mut matched = 0usize;
    for rel in relationships {
        let target_lower = rel.target.to_ascii_lowercase();
        let type_lower = rel.type_uri.to_ascii_lowercase();
        // Excel's RichData-related relationship types use several different prefixes
        // (`metadata`, `richValue*`, `rdRichValue*`, etc). Filter with a best-effort
        // substring match that catches the common cases.
        if target_lower.contains("metadata")
            || target_lower.contains("richdata")
            || type_lower.contains("metadata")
            || type_lower.contains("richdata")
            || type_lower.contains("richvalue")
        {
            matched += 1;
            let resolved = openxml::resolve_target("xl/workbook.xml", &rel.target);
            let resolved_size =
                find_part_bytes_case_insensitive(pkg, &resolved).map(|(_, b)| b.len());
            if let Some(mode) = rel.target_mode.as_deref() {
                eprintln!(
                    "  Id={} Type={} Target={} TargetMode={} (resolved: {}, {})",
                    rel.id,
                    rel.type_uri,
                    rel.target,
                    mode,
                    resolved,
                    resolved_size
                        .map(|n| format!("{n} bytes"))
                        .unwrap_or_else(|| "missing".to_string())
                );
            } else {
                eprintln!(
                    "  Id={} Type={} Target={} (resolved: {}, {})",
                    rel.id,
                    rel.type_uri,
                    rel.target,
                    resolved,
                    resolved_size
                        .map(|n| format!("{n} bytes"))
                        .unwrap_or_else(|| "missing".to_string())
                );
            }
        }
    }

    if matched == 0 {
        eprintln!("  (none)");
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_metadata_relationships(pkg: &XlsxPackage) {
    eprintln!();
    eprintln!("metadata.xml.rels (filtered richData):");

    let Some((metadata_part_name, _metadata_bytes)) =
        find_part_bytes_case_insensitive(pkg, "xl/metadata.xml")
    else {
        eprintln!("  (metadata missing)");
        return;
    };

    let rels_part_name = openxml::rels_part_name(metadata_part_name);
    let Some((rels_part_name, rels_bytes)) = find_part_bytes_case_insensitive(pkg, &rels_part_name)
    else {
        eprintln!("  (missing relationships part: {rels_part_name})");
        return;
    };
    eprintln!("  part: {rels_part_name} ({} bytes)", rels_bytes.len());

    let relationships = match openxml::parse_relationships(rels_bytes) {
        Ok(r) => r,
        Err(err) => {
            eprintln!("  warning: failed to parse {rels_part_name}: {err}");
            return;
        }
    };

    let mut matched = 0usize;
    for rel in relationships {
        let target_lower = rel.target.to_ascii_lowercase();
        let type_lower = rel.type_uri.to_ascii_lowercase();
        if target_lower.contains("richdata")
            || type_lower.contains("richdata")
            || type_lower.contains("richvalue")
        {
            matched += 1;
            let resolved = openxml::resolve_target(metadata_part_name, &rel.target);
            let resolved_size =
                find_part_bytes_case_insensitive(pkg, &resolved).map(|(_, b)| b.len());
            if let Some(mode) = rel.target_mode.as_deref() {
                eprintln!(
                    "  Id={} Type={} Target={} TargetMode={} (resolved: {}, {})",
                    rel.id,
                    rel.type_uri,
                    rel.target,
                    mode,
                    resolved,
                    resolved_size
                        .map(|n| format!("{n} bytes"))
                        .unwrap_or_else(|| "missing".to_string())
                );
            } else {
                eprintln!(
                    "  Id={} Type={} Target={} (resolved: {}, {})",
                    rel.id,
                    rel.type_uri,
                    rel.target,
                    resolved,
                    resolved_size
                        .map(|n| format!("{n} bytes"))
                        .unwrap_or_else(|| "missing".to_string())
                );
            }
        }
    }

    if matched == 0 {
        eprintln!("  (none)");
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_worksheet_vm_cm_usage(pkg: &XlsxPackage) {
    let shared_strings = parse_shared_strings_best_effort(pkg);
    let sheets = list_sheets_best_effort(pkg);
    if sheets.is_empty() {
        return;
    }

    eprintln!();
    eprintln!("worksheet vm/cm usage:");

    for sheet in &sheets {
        let Some(bytes) = pkg.part(&sheet.worksheet_part) else {
            eprintln!(
                "  {} ({}): (missing part)",
                sheet.sheet_name, sheet.worksheet_part
            );
            continue;
        };

        let scan = match scan_sheet_vm_cm_best_effort(bytes, shared_strings.as_ref()) {
            Ok(scan) => scan,
            Err(err) => {
                eprintln!(
                    "  warning: failed to scan {} ({}): {err}",
                    sheet.sheet_name, sheet.worksheet_part
                );
                VmCmScan::default()
            }
        };
        eprintln!(
            "  {} ({}): vm-cells={} cm-cells={}",
            sheet.sheet_name, sheet.worksheet_part, scan.vm_cells, scan.cm_cells
        );
        if !scan.samples.is_empty() {
            eprintln!("    samples:");
            for sample in &scan.samples {
                let value = sample.value.as_deref().unwrap_or("<no value>");
                eprintln!(
                    "      {}: vm={:?} cm={:?} value={}",
                    sample.cell_ref, sample.vm, sample.cm, value
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone)]
struct VmCmSample {
    cell_ref: String,
    vm: Option<u32>,
    cm: Option<u32>,
    value: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone, Default)]
struct VmCmScan {
    vm_cells: u64,
    cm_cells: u64,
    samples: Vec<VmCmSample>,
}

#[cfg(not(target_arch = "wasm32"))]
fn scan_sheet_vm_cm_best_effort(
    worksheet_xml: &[u8],
    shared_strings: Option<&Vec<String>>,
) -> Result<VmCmScan, quick_xml::Error> {
    const MAX_SAMPLES: usize = 5;

    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut out = VmCmScan::default();
    #[derive(Debug)]
    struct CurrentCell {
        r: Option<String>,
        t: Option<String>,
        vm: Option<u32>,
        cm: Option<u32>,
        v_text: String,
        inline_text: String,
        in_v: bool,
        in_is: bool,
        in_is_t: bool,
    }
    let mut current: Option<CurrentCell> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name.eq_ignore_ascii_case(b"c") {
                    let mut r: Option<String> = None;
                    let mut t: Option<String> = None;
                    let mut vm: Option<u32> = None;
                    let mut cm: Option<u32> = None;

                    for attr in e.attributes().with_checks(false) {
                        let Ok(attr) = attr else {
                            continue;
                        };
                        let Ok(v) = attr.unescape_value() else {
                            continue;
                        };
                        match local_name(attr.key.as_ref()) {
                            b"r" => r = Some(v.into_owned()),
                            b"t" => t = Some(v.into_owned()),
                            b"vm" => vm = v.parse::<u32>().ok(),
                            b"cm" => cm = v.parse::<u32>().ok(),
                            _ => {}
                        }
                    }

                    if vm.is_some() {
                        out.vm_cells += 1;
                    }
                    if cm.is_some() {
                        out.cm_cells += 1;
                    }

                    if out.samples.len() < MAX_SAMPLES && (vm.is_some() || cm.is_some()) {
                        current = Some(CurrentCell {
                            r,
                            t,
                            vm,
                            cm,
                            v_text: String::new(),
                            inline_text: String::new(),
                            in_v: false,
                            in_is: false,
                            in_is_t: false,
                        });
                    } else {
                        current = None;
                    }
                } else if let Some(cell) = current.as_mut() {
                    if name.eq_ignore_ascii_case(b"v") {
                        cell.in_v = true;
                    } else if name.eq_ignore_ascii_case(b"is") {
                        cell.in_is = true;
                    } else if cell.in_is && name.eq_ignore_ascii_case(b"t") {
                        cell.in_is_t = true;
                    }
                }
            }
            Ok(Event::Empty(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name.eq_ignore_ascii_case(b"c") {
                    let mut r: Option<String> = None;
                    let mut vm: Option<u32> = None;
                    let mut cm: Option<u32> = None;

                    for attr in e.attributes().with_checks(false) {
                        let Ok(attr) = attr else {
                            continue;
                        };
                        let Ok(v) = attr.unescape_value() else {
                            continue;
                        };
                        match local_name(attr.key.as_ref()) {
                            b"r" => r = Some(v.into_owned()),
                            b"vm" => vm = v.parse::<u32>().ok(),
                            b"cm" => cm = v.parse::<u32>().ok(),
                            _ => {}
                        }
                    }

                    if vm.is_some() {
                        out.vm_cells += 1;
                    }
                    if cm.is_some() {
                        out.cm_cells += 1;
                    }

                    if out.samples.len() < MAX_SAMPLES && (vm.is_some() || cm.is_some()) {
                        out.samples.push(VmCmSample {
                            cell_ref: r.unwrap_or_else(|| "<missing r>".to_string()),
                            vm,
                            cm,
                            value: None,
                        });
                    }
                }
            }
            Ok(Event::Text(t)) => {
                if let Some(cell) = current.as_mut() {
                    let Ok(text) = t.unescape() else {
                        buf.clear();
                        continue;
                    };
                    if cell.in_v {
                        cell.v_text.push_str(&text);
                    } else if cell.in_is_t {
                        cell.inline_text.push_str(&text);
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if let Some(cell) = current.as_mut() {
                    let text = String::from_utf8_lossy(t.as_ref());
                    if cell.in_v {
                        cell.v_text.push_str(&text);
                    } else if cell.in_is_t {
                        cell.inline_text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if let Some(cell) = current.as_mut() {
                    if name.eq_ignore_ascii_case(b"v") {
                        cell.in_v = false;
                    } else if name.eq_ignore_ascii_case(b"t") {
                        cell.in_is_t = false;
                    } else if name.eq_ignore_ascii_case(b"is") {
                        cell.in_is = false;
                    }
                }
                if name.eq_ignore_ascii_case(b"c") {
                    if let Some(cell) = current.take() {
                        if out.samples.len() < MAX_SAMPLES
                            && (cell.vm.is_some() || cell.cm.is_some())
                        {
                            let raw = if !cell.inline_text.is_empty() {
                                Some(cell.inline_text)
                            } else if !cell.v_text.is_empty() {
                                Some(cell.v_text)
                            } else {
                                None
                            };

                            let mut value = raw;
                            if cell.t.as_deref() == Some("s") {
                                if let (Some(shared), Some(raw)) =
                                    (shared_strings, value.as_deref())
                                {
                                    if let Ok(idx) = raw.parse::<usize>() {
                                        if let Some(s) = shared.get(idx) {
                                            value = Some(s.clone());
                                        }
                                    }
                                }
                            }

                            out.samples.push(VmCmSample {
                                cell_ref: cell.r.unwrap_or_else(|| "<missing r>".to_string()),
                                vm: cell.vm,
                                cm: cell.cm,
                                value,
                            });
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => return Err(err),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_shared_strings_best_effort(pkg: &XlsxPackage) -> Option<Vec<String>> {
    let (_part_name, bytes) = find_part_bytes_case_insensitive(pkg, "xl/sharedStrings.xml")?;
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut out: Vec<String> = Vec::new();
    let mut current_si: Option<String> = None;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name.eq_ignore_ascii_case(b"si") {
                    current_si = Some(String::new());
                } else if name.eq_ignore_ascii_case(b"t") {
                    in_t = true;
                }
            }
            Ok(Event::End(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name.eq_ignore_ascii_case(b"t") {
                    in_t = false;
                } else if name.eq_ignore_ascii_case(b"si") {
                    out.push(current_si.take().unwrap_or_default());
                }
            }
            Ok(Event::Text(t)) => {
                if in_t {
                    if let Some(si) = current_si.as_mut() {
                        if let Ok(text) = t.unescape() {
                            si.push_str(&text);
                        }
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if in_t {
                    if let Some(si) = current_si.as_mut() {
                        si.push_str(&String::from_utf8_lossy(t.as_ref()));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(err) => {
                eprintln!("warning: failed to parse xl/sharedStrings.xml: {err}");
                return None;
            }
            _ => {}
        }
        buf.clear();
    }

    Some(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn find_part_bytes_case_insensitive<'a>(
    pkg: &'a XlsxPackage,
    desired: &str,
) -> Option<(&'a str, &'a [u8])> {
    let desired = desired.strip_prefix('/').unwrap_or(desired);
    for name in pkg.part_names() {
        let canonical = name.strip_prefix('/').unwrap_or(name);
        if canonical.eq_ignore_ascii_case(desired) {
            let bytes = pkg.part(name)?;
            return Some((name, bytes));
        }
    }
    None
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
    use std::io::Write;

    let sheets = list_sheets_best_effort(pkg);
    let vm_cells = collect_vm_cells(pkg, &sheets);

    if vm_cells.is_empty() {
        // Ignore broken pipes so `... | head` doesn't panic.
        let _ = writeln!(std::io::stdout(), "no richData found");
        return false;
    }

    let index = match RichDataVmIndex::build(pkg) {
        Ok(index) => index,
        Err(err) => {
            eprintln!("warning: failed to build richData index: {err}");
            RichDataVmIndex::default()
        }
    };

    let mut rows: Vec<MappingRow> = Vec::new();
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
        // Ignore broken pipes so `... | head` doesn't panic.
        let _ = writeln!(
            std::io::stdout(),
            "{}!{} vm={} -> rv={} -> {} rel={}",
            row.sheet_name,
            row.cell_ref,
            row.vm,
            rich_value_index,
            media_part,
            rel_index
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
                .filter_map(|name| {
                    let canonical = name.strip_prefix('/').unwrap_or(name);
                    (canonical.starts_with("xl/worksheets/") && canonical.ends_with(".xml"))
                        .then_some(canonical.to_string())
                })
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
            let canonical = name.strip_prefix('/').unwrap_or(name);
            let lower = canonical.to_ascii_lowercase();
            lower == "xl/metadata.xml"
                || lower.starts_with("xl/richdata/")
                || lower.starts_with("xl/richdata/_rels/")
        })
        .map(|name| name.strip_prefix('/').unwrap_or(name))
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

        let mut manifest: Vec<String> = Vec::new();
        manifest.push(
            "sheet\tcell\tbytes\tfile\timage_part\tcalc_origin\talt_text\thyperlink".to_string(),
        );

        let mut written = 0usize;
        let mut failed = 0usize;
        let mut printed_failures = 0usize;
        let max_printed_failures = 10usize;

        // Keep output deterministic for repeated runs / log comparisons.
        let mut entries: Vec<(&str, CellRef, &formula_xlsx::EmbeddedCellImage)> =
            Vec::new();
        for ((worksheet_part, cell), image) in &images_by_cell {
            let sheet = sheet_name_by_part
                .get(worksheet_part)
                .map(String::as_str)
                .unwrap_or(worksheet_part.as_str());
            entries.push((sheet, *cell, image));
        }
        entries.sort_by(|a, b| {
            let sheet_cmp = a.0.cmp(b.0);
            if sheet_cmp != std::cmp::Ordering::Equal {
                return sheet_cmp;
            }
            (a.1.row, a.1.col).cmp(&(b.1.row, b.1.col))
        });

        for (sheet, cell, image) in entries {
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
    let mut out = String::new();
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
