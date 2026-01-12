//! Developer CLI to inspect Excel "rich data" / images-in-cell wiring.
//!
//! Usage:
//!   cargo run -p formula-xlsx --bin dump_rich_data -- <path-to-xlsx>
//!
//! Output (one line per cell with `vm="..."`):
//!   <sheet>\t<cell>\t<vm>\t<rich_value_index>\t<rel_index>\t<xl/media/* path>
//!
//! Any field that cannot be resolved is printed as `-`.

use std::collections::HashMap;
use std::error::Error;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
use std::path::PathBuf;

use formula_model::CellRef;
use formula_xlsx::rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map;
use formula_xlsx::XlsxPackage;
use quick_xml::events::Event;
use quick_xml::Reader;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn usage() -> &'static str {
    "dump_rich_data <path.xlsx>\n\
\n\
Print a best-effort mapping from worksheet cells with `vm` attributes to:\n\
  - xl/metadata.xml valueMetadata indices (vm)\n\
  - xl/richData/richValue.xml indices\n\
  - xl/richData/richValueRel.xml relationship-table indices\n\
  - resolved xl/media/* targets\n\
\n\
Output format (TSV, one line per cell):\n\
  sheet\\tcell\\tvm\\trich_value_index\\trel_index\\txl_media_path\n"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let mut args = std::env::args().skip(1);
    let Some(arg) = args.next() else {
        eprintln!("{}", usage());
        return Ok(());
    };
    if arg == "--help" || arg == "-h" {
        eprintln!("{}", usage());
        return Ok(());
    }
    if args.next().is_some() {
        return Err(format!("unexpected extra arguments\n\n{}", usage()).into());
    }

    let xlsx_path = PathBuf::from(arg);
    let bytes = fs::read(&xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    dump(&pkg);
    Ok(())
}

#[derive(Debug, Clone)]
struct SheetInfo {
    sheet_index: usize,
    sheet_name: String,
    worksheet_part: String,
}

#[derive(Debug, Clone)]
struct VmCellEntry {
    sheet_index: usize,
    sheet_name: String,
    cell_ref: String,
    cell_ref_parsed: Option<CellRef>,
    vm: String,
}

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

fn dump(pkg: &XlsxPackage) {
    let sheets = list_sheets_best_effort(pkg);
    let vm_cells = collect_vm_cells(pkg, &sheets);

    if vm_cells.is_empty() {
        println!("no richData found");
        return;
    }

    let vm_to_rv = parse_vm_to_rich_value_map_best_effort(pkg);

    let rich_value_xml_part = find_part_case_insensitive(pkg, "xl/richData/richValue.xml");
    let rich_value_rel_xml_part = find_part_case_insensitive(pkg, "xl/richData/richValueRel.xml");

    let rel_index_to_rid: Vec<Option<String>> = rich_value_rel_xml_part
        .as_deref()
        .and_then(|part| pkg.part(part))
        .and_then(|bytes| parse_rich_value_rel_table_best_effort(bytes).ok())
        .unwrap_or_default();

    let rid_to_target: HashMap<String, String> = rich_value_rel_xml_part
        .as_deref()
        .and_then(|part| parse_rich_value_rel_targets_best_effort(pkg, part).ok())
        .unwrap_or_default();

    let rich_value_index_to_rel_index: Vec<Option<u32>> = rich_value_xml_part
        .as_deref()
        .and_then(|part| pkg.part(part))
        .and_then(|bytes| {
            parse_rich_value_rel_indices_best_effort(bytes, &rel_index_to_rid, &rid_to_target).ok()
        })
        .unwrap_or_default();

    let mut rows: Vec<MappingRow> = Vec::with_capacity(vm_cells.len());
    for cell in vm_cells {
        let vm_u32 = cell.vm.parse::<u32>().ok();
        let rich_value_index = vm_u32.and_then(|vm| vm_to_rv.get(&vm).copied());

        let rel_index = rich_value_index.and_then(|rv_idx| {
            rich_value_index_to_rel_index
                .get(rv_idx as usize)
                .copied()
                .flatten()
        });

        let media_part = rel_index
            .and_then(|rel_idx| rel_index_to_rid.get(rel_idx as usize).cloned().flatten())
            .and_then(|rid| rid_to_target.get(&rid).cloned());

        rows.push(MappingRow {
            sheet_index: cell.sheet_index,
            sheet_name: cell.sheet_name,
            cell_ref: cell.cell_ref,
            cell_ref_parsed: cell.cell_ref_parsed,
            vm: cell.vm,
            rich_value_index,
            rel_index,
            media_part,
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
        println!(
            "{}\t{}\t{}\t{}\t{}\t{}",
            row.sheet_name, row.cell_ref, row.vm, rich_value_index, rel_index, media_part
        );
    }
}

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
            eprintln!("warning: failed to read workbook sheet list ({err}); falling back to xl/worksheets/*");
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
                    let key = local_name(attr.key.as_ref());
                    let Ok(value) = attr.unescape_value() else {
                        continue;
                    };
                    if key == b"r" {
                        r = Some(value.into_owned());
                    } else if key == b"vm" {
                        vm = Some(value.into_owned());
                    }
                }
                if let (Some(r), Some(vm)) = (r, vm) {
                    out.push((r, vm));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }

    out
}

fn parse_vm_to_rich_value_map_best_effort(pkg: &XlsxPackage) -> HashMap<u32, u32> {
    let Some(metadata_part) = find_part_case_insensitive(pkg, "xl/metadata.xml") else {
        return HashMap::new();
    };
    let Some(bytes) = pkg.part(&metadata_part) else {
        return HashMap::new();
    };
    match parse_value_metadata_vm_to_rich_value_index_map(bytes) {
        Ok(map) => map,
        Err(err) => {
            eprintln!("warning: failed to parse {metadata_part}: {err}");
            HashMap::new()
        }
    }
}

fn parse_rich_value_rel_table_best_effort(
    xml: &[u8],
) -> Result<Vec<Option<String>>, Box<dyn Error>> {
    let xml = std::str::from_utf8(xml)?;
    let doc = roxmltree::Document::parse(xml)?;
    let mut rels = Vec::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rel")
    {
        rels.push(
            node.attribute((REL_NS, "id"))
                .or_else(|| node.attribute("r:id"))
                .or_else(|| node.attribute("id"))
                .or_else(|| {
                    let clark = format!("{{{REL_NS}}}id");
                    node.attribute(clark.as_str())
                })
                .map(|s| s.to_string()),
        );
    }
    Ok(rels)
}

fn parse_rich_value_rel_targets_best_effort(
    pkg: &XlsxPackage,
    rich_value_rel_part: &str,
) -> Result<HashMap<String, String>, Box<dyn Error>> {
    let rels_part = rels_part_name(rich_value_rel_part);
    let Some(xml_bytes) = pkg.part(&rels_part) else {
        return Ok(HashMap::new());
    };

    let xml = std::str::from_utf8(xml_bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

    let mut out = HashMap::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        let Some(id) = node.attribute("Id") else {
            continue;
        };
        let Some(target) = node.attribute("Target") else {
            continue;
        };
        if node
            .attribute("TargetMode")
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        out.insert(id.to_string(), resolve_target(rich_value_rel_part, target));
    }

    Ok(out)
}

fn parse_rich_value_rel_indices_best_effort(
    rich_value_xml: &[u8],
    rel_index_to_rid: &[Option<String>],
    rid_to_target: &HashMap<String, String>,
) -> Result<Vec<Option<u32>>, Box<dyn Error>> {
    let xml = std::str::from_utf8(rich_value_xml)?;
    let doc = roxmltree::Document::parse(xml)?;

    let rv_nodes: Vec<roxmltree::Node<'_, '_>> = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "values")
        .map(|values| {
            values
                .children()
                .filter(|n| n.is_element() && n.tag_name().name() == "rv")
                .collect()
        })
        .unwrap_or_else(|| {
            doc.descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rv")
                .collect()
        });

    let mut out: Vec<Option<u32>> = Vec::with_capacity(rv_nodes.len());
    for rv in rv_nodes {
        out.push(resolve_rel_index_for_rv(
            rv,
            rel_index_to_rid,
            rid_to_target,
        ));
    }
    Ok(out)
}

fn resolve_rel_index_for_rv(
    rv: roxmltree::Node<'_, '_>,
    rel_index_to_rid: &[Option<String>],
    rid_to_target: &HashMap<String, String>,
) -> Option<u32> {
    let (strong, weak) = collect_rich_value_candidates(rv);

    // Prefer candidates that resolve to a concrete target.
    choose_rel_index(&strong, rel_index_to_rid, rid_to_target)
        .or_else(|| choose_rel_index(&weak, rel_index_to_rid, rid_to_target))
        .or_else(|| strong.first().copied())
        .or_else(|| weak.first().copied())
}

fn collect_rich_value_candidates(rv: roxmltree::Node<'_, '_>) -> (Vec<u32>, Vec<u32>) {
    let mut strong = Vec::new();
    let mut weak = Vec::new();
    let mut seen = std::collections::HashSet::new();

    for node in rv.descendants().filter(|n| n.is_element()) {
        let Some(text) = node.text() else {
            continue;
        };
        let text = text.trim();
        if text.is_empty() {
            continue;
        }
        let Ok(n) = text.parse::<u32>() else {
            continue;
        };
        if !seen.insert(n) {
            continue;
        }

        let name = node.tag_name().name();
        let is_strong = name.to_ascii_lowercase().contains("rel")
            || node
                .attribute("kind")
                .is_some_and(|v| v.eq_ignore_ascii_case("rel"))
            || node
                .attribute("t")
                .is_some_and(|v| v.eq_ignore_ascii_case("rel"));

        if is_strong {
            strong.push(n);
        } else {
            weak.push(n);
        }
    }

    (strong, weak)
}

fn choose_rel_index(
    candidates: &[u32],
    rel_index_to_rid: &[Option<String>],
    rid_to_target: &HashMap<String, String>,
) -> Option<u32> {
    for &idx in candidates {
        let Some(Some(rid)) = rel_index_to_rid.get(idx as usize) else {
            continue;
        };
        if rid_to_target.contains_key(rid) {
            return Some(idx);
        }
    }
    for &idx in candidates {
        if (idx as usize) < rel_index_to_rid.len() {
            return Some(idx);
        }
    }
    None
}

fn find_part_case_insensitive(pkg: &XlsxPackage, wanted: &str) -> Option<String> {
    if pkg.part(wanted).is_some() {
        return Some(wanted.to_string());
    }
    let wanted_lower = wanted.to_ascii_lowercase();
    pkg.part_names()
        .find(|name| name.to_ascii_lowercase() == wanted_lower)
        .map(|s| s.to_string())
}

fn rels_part_name(part_name: &str) -> String {
    let (dir, file) = part_name.rsplit_once('/').unwrap_or(("", part_name));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn resolve_target(base_part: &str, target: &str) -> String {
    let (target, is_absolute) = match target.strip_prefix('/') {
        Some(target) => (target, true),
        None => (target, false),
    };
    let base_dir = if is_absolute {
        ""
    } else {
        base_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("")
    };

    let mut components: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };

    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                components.pop();
            }
            _ => components.push(segment),
        }
    }

    components.join("/")
}
