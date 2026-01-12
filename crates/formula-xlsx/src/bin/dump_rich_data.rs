use std::error::Error;

#[cfg(not(target_arch = "wasm32"))]
use std::collections::{BTreeMap, HashMap, HashSet};
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
#[cfg(not(target_arch = "wasm32"))]
use std::io::Cursor;
#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::cell_images::CellImages;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::metadata::parse_metadata_xml;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::{parse_value_metadata_vm_to_rich_value_index_map, XlsxPackage};

fn usage() -> &'static str {
    "dump_rich_data <path.xlsx> [--print-parts]\n\
\n\
Debug helper for inspecting Excel rich data (linked entities, images-in-cell).\n\
\n\
Options:\n\
  --print-parts   List rich-data related ZIP parts found in the workbook\n\
"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let mut xlsx_path: Option<PathBuf> = None;
    let mut print_parts = false;

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
    dump_one(&xlsx_path, print_parts)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_one(xlsx_path: &Path, print_parts: bool) -> Result<(), Box<dyn Error>> {
    let bytes = fs::read(xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    println!("workbook: {}", xlsx_path.display());

    if print_parts {
        print_interesting_parts(&pkg);
    }

    let vm_to_rich_value_index = dump_metadata(&pkg);
    let cell_images = dump_cell_images(&pkg);
    let usage = scan_worksheet_vm_cm_usage(&pkg);

    if let Some(vm_to_rich_value_index) = vm_to_rich_value_index.as_ref() {
        dump_vm_to_rv_usage(vm_to_rich_value_index, &usage);
    }

    dump_rich_data_graph(&pkg)?;

    let rv_to_targets = resolve_rich_value_targets(&pkg)?;

    dump_rich_value_image_targets(vm_to_rich_value_index.as_ref(), cell_images.as_ref(), &rv_to_targets)?;

    dump_vm_cell_mappings(&pkg, vm_to_rich_value_index.as_ref(), &rv_to_targets);

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn print_interesting_parts(pkg: &XlsxPackage) {
    let mut parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| {
            name.starts_with("xl/richData/")
                || name.starts_with("xl/metadata.xml")
                || name.starts_with("xl/cellimages")
        })
        .collect();
    parts.sort();

    println!();
    println!("parts (rich-data related):");
    if parts.is_empty() {
        println!("  (none)");
        return;
    }
    for name in parts {
        let len = pkg.part(name).map(|b| b.len()).unwrap_or(0);
        println!("  {name} ({len} bytes)");
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_metadata(pkg: &XlsxPackage) -> Option<HashMap<u32, u32>> {
    let Some(bytes) = pkg.part("xl/metadata.xml") else {
        println!();
        println!("xl/metadata.xml: (missing)");
        return None;
    };

    println!();
    println!("xl/metadata.xml:");

    let xml = match std::str::from_utf8(bytes) {
        Ok(xml) => xml,
        Err(err) => {
            println!("  (not utf-8: {err})");
            return None;
        }
    };

    match parse_metadata_xml(xml) {
        Ok(doc) => {
            println!("  metadataTypes ({}):", doc.metadata_types.len());
            for (idx, ty) in doc.metadata_types.iter().enumerate() {
                let idx_1b = idx + 1;
                if ty.name.is_empty() {
                    println!("    {idx_1b}: (missing name)");
                } else {
                    println!("    {idx_1b}: {}", ty.name);
                }
            }
            println!("  cellMetadata blocks: {}", doc.cell_metadata.len());
            println!("  valueMetadata blocks: {}", doc.value_metadata.len());
            println!("  futureMetadata blocks: {}", doc.future_metadata_blocks.len());
        }
        Err(err) => {
            println!("  (failed to parse metadata.xml: {err})");
        }
    }

    match parse_value_metadata_vm_to_rich_value_index_map(bytes) {
        Ok(mut map) => {
            if map.is_empty() {
                println!("  vm -> rich_value_index: (none resolved)");
                return Some(map);
            }

            println!("  vm -> rich_value_index ({}):", map.len());
            let mut entries: Vec<(u32, u32)> = map.drain().collect();
            entries.sort_by_key(|(vm, _)| *vm);

            let max = 50usize;
            for (idx, (vm, rv)) in entries.iter().enumerate() {
                if idx >= max {
                    println!("    ... ({} more)", entries.len().saturating_sub(max));
                    break;
                }
                println!("    vm {vm} -> rv {rv}");
            }

            Some(entries.into_iter().collect())
        }
        Err(err) => {
            println!("  vm -> rich_value_index: (failed to resolve: {err})");
            None
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_cell_images(pkg: &XlsxPackage) -> Option<CellImages> {
    let has_cellimages = pkg.part_names().any(is_cell_images_part_name);
    if !has_cellimages {
        return None;
    }

    println!();
    println!("cell images (xl/cellimages*.xml):");

    let mut workbook = formula_model::Workbook::default();
    match CellImages::parse_from_parts(pkg.parts_map(), &mut workbook) {
        Ok(cell_images) => {
            if cell_images.parts.is_empty() {
                println!("  (no cell images parts parsed)");
                return Some(cell_images);
            }
            for part in &cell_images.parts {
                println!("  part: {}", part.path);
                println!("    rels: {}", part.rels_path);
                if part.images.is_empty() {
                    println!("    images: (none)");
                    continue;
                }
                println!("    images ({}):", part.images.len());
                for (idx, img) in part.images.iter().enumerate() {
                    println!("      [{idx}] {} -> {}", img.embed_rel_id, img.target_path);
                }
            }
            Some(cell_images)
        }
        Err(err) => {
            println!("  (failed to parse cell images: {err})");
            None
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone, Default)]
struct VmCmUsage {
    total_cells: u64,
    vm_cells: u64,
    cm_cells: u64,
    vm_counts: HashMap<u32, u64>,
    cm_counts: HashMap<u32, u64>,
}

#[cfg(not(target_arch = "wasm32"))]
fn scan_worksheet_vm_cm_usage(pkg: &XlsxPackage) -> VmCmUsage {
    let mut usage = VmCmUsage::default();

    let sheets = match pkg.worksheet_parts() {
        Ok(parts) => parts,
        Err(err) => {
            println!();
            println!("worksheets: (failed to resolve sheet parts: {err})");
            return usage;
        }
    };

    if sheets.is_empty() {
        return usage;
    }

    println!();
    println!("worksheet vm/cm usage:");

    for sheet in sheets {
        let Some(bytes) = pkg.part(&sheet.worksheet_part) else {
            println!("  {} ({}): (missing part)", sheet.name, sheet.worksheet_part);
            continue;
        };

        let mut sheet_total_cells: u64 = 0;
        let mut sheet_vm_cells: u64 = 0;
        let mut sheet_cm_cells: u64 = 0;

        let mut sheet_vm_counts: HashMap<u32, u64> = HashMap::new();
        let mut sheet_cm_counts: HashMap<u32, u64> = HashMap::new();

        let mut reader = quick_xml::Reader::from_reader(Cursor::new(bytes));
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e))
                | Ok(quick_xml::events::Event::Empty(e)) => {
                    if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"c") {
                        sheet_total_cells += 1;
                        let mut has_vm = false;
                        let mut has_cm = false;
                        for attr in e.attributes().with_checks(false) {
                            let attr = match attr {
                                Ok(attr) => attr,
                                Err(_) => continue,
                            };
                            let key = local_name(attr.key.as_ref());
                            if !key.eq_ignore_ascii_case(b"vm") && !key.eq_ignore_ascii_case(b"cm") {
                                continue;
                            }
                            let val = match attr.unescape_value() {
                                Ok(v) => v,
                                Err(_) => continue,
                            };
                            let Ok(parsed) = val.parse::<u32>() else {
                                continue;
                            };
                            if key.eq_ignore_ascii_case(b"vm") {
                                has_vm = true;
                                *sheet_vm_counts.entry(parsed).or_insert(0) += 1;
                                *usage.vm_counts.entry(parsed).or_insert(0) += 1;
                            } else {
                                has_cm = true;
                                *sheet_cm_counts.entry(parsed).or_insert(0) += 1;
                                *usage.cm_counts.entry(parsed).or_insert(0) += 1;
                            }
                        }
                        if has_vm {
                            sheet_vm_cells += 1;
                        }
                        if has_cm {
                            sheet_cm_cells += 1;
                        }
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }

        usage.total_cells += sheet_total_cells;
        usage.vm_cells += sheet_vm_cells;
        usage.cm_cells += sheet_cm_cells;

        println!(
            "  {} ({}): cells={} vm-cells={} cm-cells={}",
            sheet.name, sheet.worksheet_part, sheet_total_cells, sheet_vm_cells, sheet_cm_cells
        );
        print_top_counts("vm", &sheet_vm_counts, 10);
        print_top_counts("cm", &sheet_cm_counts, 10);
    }

    println!(
        "  TOTAL: cells={} vm-cells={} cm-cells={}",
        usage.total_cells, usage.vm_cells, usage.cm_cells
    );
    print_top_counts("vm", &usage.vm_counts, 10);
    print_top_counts("cm", &usage.cm_counts, 10);

    usage
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_vm_to_rv_usage(vm_to_rv: &HashMap<u32, u32>, usage: &VmCmUsage) {
    if vm_to_rv.is_empty() || usage.vm_counts.is_empty() {
        return;
    }

    let mut rv_counts: HashMap<u32, u64> = HashMap::new();
    for (vm, count) in &usage.vm_counts {
        let Some(rv) = vm_to_rv.get(vm) else {
            continue;
        };
        *rv_counts.entry(*rv).or_insert(0) += *count;
    }

    if rv_counts.is_empty() {
        return;
    }

    println!();
    println!("derived rich value usage (from worksheet vm + metadata.xml):");
    print_top_counts("rv", &rv_counts, 20);
}

#[cfg(not(target_arch = "wasm32"))]
fn print_top_counts(label: &str, counts: &HashMap<u32, u64>, max: usize) {
    if counts.is_empty() {
        return;
    }
    let mut entries: Vec<(u32, u64)> = counts.iter().map(|(k, v)| (*k, *v)).collect();
    entries.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| a.0.cmp(&b.0)));
    println!("    top {label}:");
    for (idx, (k, v)) in entries.into_iter().enumerate() {
        if idx >= max {
            break;
        }
        println!("      {label} {k}: {v}");
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_rich_data_graph(pkg: &XlsxPackage) -> Result<(), Box<dyn Error>> {
    let mut rich_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| name.starts_with("xl/richData/") && name.ends_with(".xml"))
        .collect();
    rich_parts.sort();
    if rich_parts.is_empty() {
        return Ok(());
    }

    println!();
    println!("richData part graph (relationships):");

    for part in rich_parts {
        let rels_part = rels_part_name(part);
        let Some(rels_bytes) = pkg.part(&rels_part) else {
            println!("  {part}: (no rels part)");
            continue;
        };

        let relationships = match parse_relationships(rels_bytes) {
            Ok(r) => r,
            Err(err) => {
                println!("  {part}: (failed to parse rels: {err})");
                continue;
            }
        };

        if relationships.is_empty() {
            println!("  {part}: (no relationships)");
            continue;
        }

        println!("  {part}:");
        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
            {
                println!("    {} [{}] -> {} (external)", rel.id, rel.type_uri, rel.target);
                continue;
            }
            let resolved = resolve_target(part, &rel.target);
            println!(
                "    {} [{}] -> {} (resolved: {})",
                rel.id, rel.type_uri, rel.target, resolved
            );
        }
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_rich_value_image_targets(
    vm_to_rv: Option<&HashMap<u32, u32>>,
    cell_images: Option<&CellImages>,
    rv_to_targets: &BTreeMap<u32, Vec<String>>,
) -> Result<(), Box<dyn Error>> {
    if rv_to_targets.is_empty() {
        return Ok(());
    }

    // Filter down to targets that look image-related.
    let mut rv_to_images: BTreeMap<u32, Vec<String>> = BTreeMap::new();
    for (rv, targets) in rv_to_targets {
        let mut out: Vec<String> = Vec::new();
        for target in targets {
            let is_cell_images = is_cell_images_part_name(target);
            if is_probable_image_target(target) {
                out.push(target.clone());
            }

            if is_cell_images {
                if let Some(expanded) = expand_cellimages_target(*rv, target, cell_images) {
                    out.extend(expanded);
                }
            }
        }

        out.sort();
        out.dedup();
        if !out.is_empty() {
            rv_to_images.insert(*rv, out);
        }
    }

    if rv_to_images.is_empty() {
        return Ok(());
    }

    println!();
    println!("rich_value_index -> image target (best-effort):");

    let referenced_rvs: Option<HashSet<u32>> = vm_to_rv.map(|map| map.values().copied().collect());

    for (rv, targets) in rv_to_images {
        let referenced = referenced_rvs
            .as_ref()
            .is_some_and(|set| set.contains(&rv));
        if referenced {
            println!("  rv {rv} (referenced by metadata.xml):");
        } else {
            println!("  rv {rv}:");
        }
        for target in targets {
            println!("    {target}");
        }
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_vm_cell_mappings(
    pkg: &XlsxPackage,
    vm_to_rv: Option<&HashMap<u32, u32>>,
    rv_to_targets: &BTreeMap<u32, Vec<String>>,
) {
    let Some(vm_to_rv) = vm_to_rv else {
        return;
    };

    let sheets = match pkg.worksheet_parts() {
        Ok(parts) => parts,
        Err(_) => return,
    };

    let mut rows: Vec<(String, String, u32, Option<u32>, Option<String>)> = Vec::new();

    for sheet in sheets {
        let Some(bytes) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };

        let mut reader = quick_xml::Reader::from_reader(Cursor::new(bytes));
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut in_sheet_data = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(e))
                    if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"sheetData") =>
                {
                    in_sheet_data = true;
                }
                Ok(quick_xml::events::Event::End(e))
                    if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"sheetData") =>
                {
                    in_sheet_data = false;
                }
                Ok(quick_xml::events::Event::Start(e))
                | Ok(quick_xml::events::Event::Empty(e))
                    if in_sheet_data && local_name(e.name().as_ref()).eq_ignore_ascii_case(b"c") =>
                {
                    let mut r: Option<String> = None;
                    let mut vm: Option<u32> = None;
                    for attr in e.attributes().with_checks(false) {
                        let Ok(attr) = attr else {
                            continue;
                        };
                        let key = local_name(attr.key.as_ref());
                        let Ok(value) = attr.unescape_value() else {
                            continue;
                        };
                        if key.eq_ignore_ascii_case(b"r") {
                            r = Some(value.into_owned());
                        } else if key.eq_ignore_ascii_case(b"vm") {
                            vm = value.parse::<u32>().ok();
                        }
                    }
                    let Some(r) = r else { continue };
                    let Some(vm) = vm else { continue };

                    let rv = vm_to_rv.get(&vm).copied();
                    let target = rv.and_then(|rv| {
                        rv_to_targets
                            .get(&rv)
                            .and_then(|targets| {
                                targets
                                    .iter()
                                    .find(|t| is_probable_image_target(t) || is_cell_images_part_name(t))
                                    .or_else(|| targets.first())
                                    .cloned()
                            })
                    });
                    rows.push((sheet.name.clone(), r, vm, rv, target));
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }

    if rows.is_empty() {
        return;
    }

    rows.sort_by(|a, b| a.0.cmp(&b.0).then_with(|| a.1.cmp(&b.1)));

    let rv_to_rel_values: BTreeMap<u32, Vec<u32>> = find_part_case_insensitive(pkg, "xl/richData/richValue.xml")
        .and_then(|(_name, bytes)| extract_rich_value_record_rel_values(&bytes).ok())
        .unwrap_or_default();

    println!();
    println!("vm cell mappings (sheet, cell, vm -> rv -> target):");
    let max = 50usize;
    for (idx, (sheet, cell, vm, rv, target)) in rows.iter().enumerate() {
        if idx >= max {
            println!("  ... (more omitted)");
            break;
        }
        println!(
            "  {sheet}!{cell} vm={vm} -> rv={} -> {}",
            rv.map(|n| n.to_string()).unwrap_or_else(|| "-".to_string()),
            target.clone().unwrap_or_else(|| "-".to_string())
        );
    }

    // Machine-friendly output for quick copy/paste/grepping.
    //
    // Format:
    //   sheet<TAB>cell<TAB>vm<TAB>rv<TAB>rel<TAB>target
    println!();
    println!("vm cell mappings (tsv):");
    println!("sheet\tcell\tvm\trv\trel\ttarget");
    for (idx, (sheet, cell, vm, rv, target)) in rows.into_iter().enumerate() {
        if idx >= max {
            println!("... (more omitted)");
            break;
        }

        let rel = rv
            .and_then(|rv| rv_to_rel_values.get(&rv))
            .and_then(|rels| rels.first().copied());

        println!(
            "{sheet}\t{cell}\t{vm}\t{}\t{}\t{}",
            rv.map(|n| n.to_string()).unwrap_or_else(|| "-".to_string()),
            rel.map(|n| n.to_string()).unwrap_or_else(|| "-".to_string()),
            target.unwrap_or_else(|| "-".to_string())
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn expand_cellimages_target(
    rv: u32,
    cellimages_part: &str,
    cell_images: Option<&CellImages>,
) -> Option<Vec<String>> {
    let cell_images = cell_images?;
    let part = cell_images
        .parts
        .iter()
        .find(|p| p.path.as_str() == cellimages_part)?;

    if part.images.is_empty() {
        return None;
    }

    // Heuristics:
    // - If there's only one image, assume it's the one referenced.
    // - Otherwise, if `rv` is in-bounds, assume `rv` is the image entry index.
    let mut out = Vec::new();
    if part.images.len() == 1 {
        out.push(format!(
            "{} (via {}[0])",
            part.images[0].target_path, part.path
        ));
        return Some(out);
    }

    let idx = rv as usize;
    if let Some(img) = part.images.get(idx) {
        out.push(format!("{} (via {}[{idx}])", img.target_path, part.path));
        return Some(out);
    }

    None
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_rich_value_targets(pkg: &XlsxPackage) -> Result<BTreeMap<u32, Vec<String>>, Box<dyn Error>> {
    let mut out: BTreeMap<u32, Vec<String>> = BTreeMap::new();

    if let Some((part_name, bytes)) = find_part_case_insensitive(pkg, "xl/richData/richValue.xml") {
        let rels_part = rels_part_name(&part_name);
        let rels = pkg
            .part(&rels_part)
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let rel_by_id: HashMap<String, Relationship> =
            rels.into_iter().map(|r| (r.id.clone(), r)).collect();

        if let Ok(rv_ids) = extract_rich_value_record_rids(&bytes) {
            for (rv, rids) in rv_ids {
                for rid in rids {
                    if let Some(rel) = rel_by_id.get(&rid) {
                        if rel
                            .target_mode
                            .as_deref()
                            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
                        {
                            continue;
                        }
                        let resolved = resolve_target(&part_name, &rel.target);
                        out.entry(rv).or_default().push(resolved);
                    }
                }
            }
        }
    }

    // Some workbooks appear to store relationship IDs in `richValueRel.xml` instead of directly
    // in `richValue.xml`. Treat the nth discovered rid in that file as `rv=n` as a best-effort
    // fallback.
    if let Some((part_name, bytes)) = find_part_case_insensitive(pkg, "xl/richData/richValueRel.xml") {
        let rels_part = rels_part_name(&part_name);
        let rels = pkg
            .part(&rels_part)
            .map(parse_relationships)
            .transpose()?
            .unwrap_or_default();
        let rel_by_id: HashMap<String, Relationship> =
            rels.into_iter().map(|r| (r.id.clone(), r)).collect();

        if let Ok(rids) = extract_relationship_ids_in_order(&bytes) {
            for (idx, rid) in rids.into_iter().enumerate() {
                let rv = idx as u32;
                if out.get(&rv).is_some() {
                    // Prefer direct `richValue.xml` resolution when available.
                    continue;
                }
                let Some(rel) = rel_by_id.get(&rid) else {
                    continue;
                };
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }
                let resolved = resolve_target(&part_name, &rel.target);
                out.entry(rv).or_default().push(resolved);
            }
        }
    }

    for targets in out.values_mut() {
        targets.sort();
        targets.dedup();
    }

    Ok(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_rich_value_record_rids(bytes: &[u8]) -> Result<BTreeMap<u32, Vec<String>>, Box<dyn Error>> {
    let xml = std::str::from_utf8(bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

    let rv_nodes: Vec<roxmltree::Node<'_, '_>> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        .collect();
    if rv_nodes.is_empty() {
        return Ok(BTreeMap::new());
    }

    let mut out: BTreeMap<u32, Vec<String>> = BTreeMap::new();
    for (idx, rv_node) in rv_nodes.into_iter().enumerate() {
        let mut rids: Vec<String> = Vec::new();
        let mut seen: HashSet<String> = HashSet::new();

        for node in rv_node.descendants() {
            if !node.is_element() {
                continue;
            }
            for attr in node.attributes() {
                let val = attr.value();
                if looks_like_rid(val) && seen.insert(val.to_string()) {
                    rids.push(val.to_string());
                }
            }
        }

        if !rids.is_empty() {
            out.insert(idx as u32, rids);
        }
    }

    Ok(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_rich_value_record_rel_values(
    bytes: &[u8],
) -> Result<BTreeMap<u32, Vec<u32>>, Box<dyn Error>> {
    let xml = std::str::from_utf8(bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

    let rv_nodes: Vec<roxmltree::Node<'_, '_>> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        .collect();
    if rv_nodes.is_empty() {
        return Ok(BTreeMap::new());
    }

    let mut out: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    for (idx, rv_node) in rv_nodes.into_iter().enumerate() {
        let mut rels: Vec<u32> = Vec::new();
        for node in rv_node.descendants() {
            if !node.is_element() {
                continue;
            }
            if node.tag_name().name() != "v" {
                continue;
            }
            let kind = node.attribute("kind").unwrap_or_default();
            if !kind.eq_ignore_ascii_case("rel") {
                continue;
            }
            let Some(text) = node.text() else {
                continue;
            };
            if let Ok(value) = text.trim().parse::<u32>() {
                rels.push(value);
            }
        }

        rels.sort();
        rels.dedup();
        if !rels.is_empty() {
            out.insert(idx as u32, rels);
        }
    }

    Ok(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn extract_relationship_ids_in_order(bytes: &[u8]) -> Result<Vec<String>, Box<dyn Error>> {
    let xml = std::str::from_utf8(bytes)?;
    let doc = roxmltree::Document::parse(xml)?;

    let mut out: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();

    for node in doc.descendants().filter(|n| n.is_element()) {
        // Take at most one relationship ID per element to preserve the most likely ordering.
        let mut found: Option<&str> = None;
        for attr in node.attributes() {
            if looks_like_rid(attr.value()) {
                found = Some(attr.value());
                break;
            }
        }

        if let Some(rid) = found {
            if seen.insert(rid.to_string()) {
                out.push(rid.to_string());
            }
        }
    }

    Ok(out)
}

#[cfg(not(target_arch = "wasm32"))]
fn looks_like_rid(value: &str) -> bool {
    let value = value.trim();
    let lower = value.to_ascii_lowercase();
    let Some(rest) = lower.strip_prefix("rid") else {
        return false;
    };
    !rest.is_empty() && rest.bytes().all(|b| b.is_ascii_digit())
}

#[cfg(not(target_arch = "wasm32"))]
fn is_cell_images_part_name(path: &str) -> bool {
    let Some(rest) = path.strip_prefix("xl/") else {
        return false;
    };
    if rest.contains('/') {
        return false;
    }
    let lower = rest.to_ascii_lowercase();
    lower.starts_with("cellimages") && lower.ends_with(".xml")
}

#[cfg(not(target_arch = "wasm32"))]
fn is_probable_image_target(path: &str) -> bool {
    if path.starts_with("xl/media/") {
        return true;
    }
    let lower = path.to_ascii_lowercase();
    matches!(
        lower.rsplit_once('.').map(|(_, ext)| ext),
        Some("png" | "jpg" | "jpeg" | "gif" | "bmp" | "tif" | "tiff" | "emf" | "wmf")
    )
}

#[cfg(not(target_arch = "wasm32"))]
fn find_part_case_insensitive(pkg: &XlsxPackage, desired: &str) -> Option<(String, Vec<u8>)> {
    let desired_lower = desired.to_ascii_lowercase();
    for name in pkg.part_names() {
        if name.to_ascii_lowercase() == desired_lower {
            return pkg.part(name).map(|b| (name.to_string(), b.to_vec()));
        }
    }
    None
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    type_uri: String,
    target: String,
    target_mode: Option<String>,
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_relationships(xml: &[u8]) -> Result<Vec<Relationship>, Box<dyn Error>> {
    let mut reader = quick_xml::Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut relationships = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            quick_xml::events::Event::Start(start) | quick_xml::events::Event::Empty(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"Relationship") {
                    let mut id = None;
                    let mut target = None;
                    let mut type_uri = None;
                    let mut target_mode = None;
                    for attr in start.attributes().with_checks(false) {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"Id") {
                            id = Some(value);
                        } else if key.eq_ignore_ascii_case(b"Target") {
                            target = Some(value);
                        } else if key.eq_ignore_ascii_case(b"Type") {
                            type_uri = Some(value);
                        } else if key.eq_ignore_ascii_case(b"TargetMode") {
                            target_mode = Some(value);
                        }
                    }
                    if let (Some(id), Some(target), Some(type_uri)) = (id, target, type_uri) {
                        relationships.push(Relationship {
                            id,
                            target,
                            type_uri,
                            target_mode,
                        });
                    }
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(relationships)
}

#[cfg(not(target_arch = "wasm32"))]
fn rels_part_name(part_name: &str) -> String {
    let (dir, file) = part_name.rsplit_once('/').unwrap_or(("", part_name));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_target(base_part: &str, target: &str) -> String {
    let (target, is_absolute) = match target.strip_prefix('/') {
        Some(target) => (target, true),
        None => (target, false),
    };
    let base_dir = if is_absolute {
        ""
    } else {
        base_part
            .rsplit_once('/')
            .map(|(dir, _)| dir)
            .unwrap_or("")
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

#[cfg(not(target_arch = "wasm32"))]
fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}
