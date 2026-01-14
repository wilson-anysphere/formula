#[cfg(not(target_arch = "wasm32"))]
use std::error::Error;
#[cfg(not(target_arch = "wasm32"))]
use std::io;
#[cfg(not(target_arch = "wasm32"))]
use std::io::Cursor;
#[cfg(not(target_arch = "wasm32"))]
use std::io::Write;
#[cfg(not(target_arch = "wasm32"))]
use std::path::PathBuf;

#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::metadata::parse_metadata_xml;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::openxml;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::XlsxPackage;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::events::Event;
#[cfg(not(target_arch = "wasm32"))]
use quick_xml::Reader;
#[cfg(not(target_arch = "wasm32"))]
use std::cmp::Ordering;

#[cfg(not(target_arch = "wasm32"))]
const MAX_CELL_REFS: usize = 10;

#[cfg(not(target_arch = "wasm32"))]
mod workbook_open;

#[cfg(not(target_arch = "wasm32"))]
fn usage() -> &'static str {
    "dump_metadata [--password <pw>] <path.xlsx|path.xlsm>\n\
\n\
Debug CLI to inspect SpreadsheetML linked/rich-data metadata:\n\
  - (use --password '' for empty-password encrypted workbooks)\n\
  - xl/metadata.xml\n\
  - xl/_rels/metadata.xml.rels\n\
  - xl/richData/*\n\
  - worksheet <c> vm/cm attributes\n"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let result = main_inner();
    if let Err(err) = result {
        if let Some(io_err) = err.downcast_ref::<io::Error>() {
            if io_err.kind() == io::ErrorKind::BrokenPipe {
                // Allow piping output to tools like `head` without panicking.
                return Ok(());
            }
        }
        return Err(err);
    }
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn main_inner() -> Result<(), Box<dyn Error>> {
    let mut xlsx_path: Option<PathBuf> = None;
    let mut password: Option<String> = None;

    let mut args = std::env::args().skip(1).peekable();
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--help" | "-h" => {
                eprintln!("{}", usage());
                return Ok(());
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

    let path = xlsx_path.ok_or_else(|| format!("missing <path.xlsx|path.xlsm>\n\n{}", usage()))?;

    let pkg = workbook_open::open_xlsx_package(&path, password.as_deref())?;

    let mut out = io::BufWriter::new(io::stdout());

    writeln!(out, "workbook: {}", path.display())?;
    writeln!(out)?;

    dump_metadata_xml(&pkg, &mut out)?;
    dump_metadata_relationships(&pkg, &mut out)?;
    dump_richdata_parts(&pkg, &mut out)?;
    dump_worksheet_vm_cm_summary(&pkg, &mut out)?;

    out.flush()?;

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_metadata_xml(pkg: &XlsxPackage, out: &mut dyn Write) -> io::Result<()> {
    writeln!(out, "[xl/metadata.xml]")?;

    let Some(bytes) = pkg.part("xl/metadata.xml") else {
        writeln!(out, "  (missing)")?;
        writeln!(out)?;
        return Ok(());
    };

    writeln!(out, "  size: {} bytes", bytes.len())?;

    match parse_metadata_xml(bytes) {
        Ok(doc) => {
            writeln!(out, "  metadataTypes: {}", doc.metadata_types.len())?;
            for (idx, ty) in doc.metadata_types.iter().enumerate() {
                let attrs: Vec<String> = ty
                    .attributes
                    .iter()
                    .filter(|(k, _)| k.as_str() != "name")
                    .map(|(k, v)| format!("{k}={v}"))
                    .collect();
                if attrs.is_empty() {
                    writeln!(out, "    [{idx}] {}", ty.name)?;
                } else {
                    writeln!(out, "    [{idx}] {} ({})", ty.name, attrs.join(", "))?;
                }
            }

            let cell_metadata_blocks = doc.cell_metadata.len();
            let cell_metadata_rc_total: usize =
                doc.cell_metadata.iter().map(|b| b.block.records.len()).sum();
            writeln!(
                out,
                "  cellMetadata: blocks={} rc_records={}",
                cell_metadata_blocks, cell_metadata_rc_total
            )?;

            let value_metadata_blocks = doc.value_metadata.len();
            let value_metadata_rc_total: usize =
                doc.value_metadata.iter().map(|b| b.block.records.len()).sum();
            writeln!(
                out,
                "  valueMetadata: blocks={} rc_records={}",
                value_metadata_blocks, value_metadata_rc_total
            )?;

            writeln!(
                out,
                "  futureMetadata blocks: {}",
                doc.future_metadata_blocks.len()
            )?;
            for (idx, block) in doc.future_metadata_blocks.iter().enumerate() {
                writeln!(
                    out,
                    "    [{idx}] count={} inner_xml_chars={}",
                    block.count,
                    block.block.inner_xml.len()
                )?;
            }
        }
        Err(err) => {
            writeln!(out, "  warning: failed to parse metadata.xml: {err}")?;
        }
    }

    writeln!(out)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_metadata_relationships(pkg: &XlsxPackage, out: &mut dyn Write) -> io::Result<()> {
    writeln!(out, "[xl/_rels/metadata.xml.rels]")?;

    let Some(bytes) = pkg.part("xl/_rels/metadata.xml.rels") else {
        writeln!(out, "  (missing)")?;
        writeln!(out)?;
        return Ok(());
    };

    writeln!(out, "  size: {} bytes", bytes.len())?;

    let mut relationships = match openxml::parse_relationships(bytes) {
        Ok(rels) => rels,
        Err(err) => {
            writeln!(out, "  warning: failed to parse relationships: {err}")?;
            writeln!(out)?;
            return Ok(());
        }
    };

    relationships.sort_by(|a, b| relationship_id_sort_cmp(&a.id, &b.id));

    writeln!(out, "  relationships: {}", relationships.len())?;
    for rel in relationships {
        let mode = rel
            .target_mode
            .as_deref()
            .map(|s| s.trim())
            .filter(|s| !s.is_empty())
            .unwrap_or("Internal");
        let resolved = if mode.eq_ignore_ascii_case("External") {
            "<external>".to_string()
        } else {
            openxml::resolve_target("xl/metadata.xml", &rel.target)
        };
        writeln!(
            out,
            "    - id={} type={} target={} mode={} resolved={}",
            rel.id, rel.type_uri, rel.target, mode, resolved
        )?;
    }

    writeln!(out)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_richdata_parts(pkg: &XlsxPackage, out: &mut dyn Write) -> io::Result<()> {
    writeln!(out, "[xl/richData/*]")?;

    let mut parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| name.starts_with("xl/richData/"))
        .collect();
    parts.sort();

    writeln!(out, "  parts: {}", parts.len())?;
    for name in parts {
        writeln!(out, "    - {name}")?;
    }

    writeln!(out)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Default)]
struct WorksheetCellMetadataSummary {
    cm_cells: usize,
    vm_cells: usize,
    cm_refs: Vec<String>,
    vm_refs: Vec<String>,
    warning: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_worksheet_vm_cm_summary(pkg: &XlsxPackage, out: &mut dyn Write) -> io::Result<()> {
    writeln!(out, "[xl/worksheets/sheet*.xml vm/cm]")?;

    let mut sheets: Vec<&str> = pkg
        .part_names()
        .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .collect();
    sheets.sort_by(|a, b| worksheet_part_sort_cmp(a, b));

    if sheets.is_empty() {
        writeln!(
            out,
            "  (no worksheet parts matched xl/worksheets/sheet*.xml)"
        )?;
        writeln!(out)?;
        return Ok(());
    }

    for sheet_part in sheets {
        writeln!(out, "  {sheet_part}")?;
        let Some(bytes) = pkg.part(sheet_part) else {
            writeln!(out, "    warning: missing part bytes")?;
            continue;
        };

        let summary = scan_worksheet_cells_for_vm_cm(bytes, MAX_CELL_REFS);
        writeln!(
            out,
            "    cells with cm: {}{}",
            summary.cm_cells,
            format_first_refs(summary.cm_cells, &summary.cm_refs)
        )?;
        writeln!(
            out,
            "    cells with vm: {}{}",
            summary.vm_cells,
            format_first_refs(summary.vm_cells, &summary.vm_refs)
        )?;
        if let Some(warn) = summary.warning {
            writeln!(out, "    warning: {warn}")?;
        }
    }

    writeln!(out)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn relationship_id_sort_cmp(a: &str, b: &str) -> Ordering {
    match (parse_rid_index(a), parse_rid_index(b)) {
        (Some(ai), Some(bi)) => ai.cmp(&bi).then_with(|| a.cmp(b)),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => a.cmp(b),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn worksheet_part_sort_cmp(a: &str, b: &str) -> Ordering {
    match (parse_sheet_index(a), parse_sheet_index(b)) {
        (Some(ai), Some(bi)) => ai.cmp(&bi).then_with(|| a.cmp(b)),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => a.cmp(b),
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_rid_index(id: &str) -> Option<u32> {
    let prefix = id.get(0..3)?;
    if !prefix.eq_ignore_ascii_case("rid") {
        return None;
    }
    id[3..].parse().ok()
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_sheet_index(part_name: &str) -> Option<u32> {
    let file_name = part_name.rsplit('/').next()?;
    let rest = file_name.strip_prefix("sheet")?;
    let num = rest.strip_suffix(".xml")?;
    num.parse().ok()
}

#[cfg(not(target_arch = "wasm32"))]
fn format_first_refs(total: usize, refs: &[String]) -> String {
    if refs.is_empty() {
        return String::new();
    }
    format!(" (first {} of {}: {})", refs.len(), total, refs.join(", "))
}

#[cfg(not(target_arch = "wasm32"))]
fn scan_worksheet_cells_for_vm_cm(bytes: &[u8], max_refs: usize) -> WorksheetCellMetadataSummary {
    let mut summary = WorksheetCellMetadataSummary::default();

    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(ev) => ev,
            Err(err) => {
                summary.warning.get_or_insert_with(|| format!("xml parse error: {err}"));
                break;
            }
        };

        match event {
            Event::Start(e) | Event::Empty(e) => {
                if openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"c") {
                    let mut cell_ref: Option<String> = None;
                    let mut has_cm = false;
                    let mut has_vm = false;

                    for attr in e.attributes().with_checks(false) {
                        match attr {
                            Ok(attr) => {
                                let key = openxml::local_name(attr.key.as_ref());
                                if key.eq_ignore_ascii_case(b"r") {
                                    match attr.unescape_value() {
                                        Ok(v) => {
                                            cell_ref = Some(v.into_owned());
                                        }
                                        Err(err) => {
                                            summary
                                                .warning
                                                .get_or_insert_with(|| format!("bad cell ref value: {err}"));
                                        }
                                    }
                                } else if key.eq_ignore_ascii_case(b"cm") {
                                    has_cm = true;
                                } else if key.eq_ignore_ascii_case(b"vm") {
                                    has_vm = true;
                                }
                            }
                            Err(err) => {
                                summary
                                    .warning
                                    .get_or_insert_with(|| format!("xml attribute error: {err}"));
                            }
                        }
                    }

                    if has_cm {
                        summary.cm_cells += 1;
                        if summary.cm_refs.len() < max_refs {
                            if let Some(r) = cell_ref.as_ref() {
                                summary.cm_refs.push(r.clone());
                            }
                        }
                    }
                    if has_vm {
                        summary.vm_cells += 1;
                        if summary.vm_refs.len() < max_refs {
                            if let Some(r) = cell_ref.as_ref() {
                                summary.vm_refs.push(r.clone());
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    summary
}
