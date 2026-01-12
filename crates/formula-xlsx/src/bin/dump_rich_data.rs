#[cfg(not(target_arch = "wasm32"))]
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
use formula_xlsx::openxml;
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::rich_data::{rich_value, rich_value_rel};
#[cfg(not(target_arch = "wasm32"))]
use formula_xlsx::{parse_value_metadata_vm_to_rich_value_index_map, XlsxPackage};

#[cfg(not(target_arch = "wasm32"))]
fn usage() -> &'static str {
    "dump_rich_data <path.xlsx> [--print-parts] [--extract-cell-images] [--extract-cell-images-out <dir>]\n\
\n\
Debug helper for inspecting Excel rich data (linked entities, images-in-cell).\n\
\n\
Options:\n\
  --print-parts           List rich-data related ZIP parts found in the workbook\n\
  --extract-cell-images   Extract rich-data in-cell images by cell and print a summary\n\
  --extract-cell-images-out <dir>\n\
                         Write extracted in-cell images to <dir> as files\n\
"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

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
    dump_one(
        &xlsx_path,
        print_parts,
        extract_cell_images,
        extract_cell_images_out.as_deref(),
    )?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_one(
    xlsx_path: &Path,
    print_parts: bool,
    extract_cell_images: bool,
    extract_cell_images_out: Option<&Path>,
) -> Result<(), Box<dyn Error>> {
    let bytes = fs::read(xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    println!("workbook: {}", xlsx_path.display());

    if print_parts {
        print_interesting_parts(&pkg);
    }

    let vm_to_rich_value_index = dump_metadata(&pkg);
    dump_metadata_relationships(&pkg);
    let cell_images = dump_cell_images(&pkg);
    let usage = scan_worksheet_vm_cm_usage(&pkg);

    if let Some(vm_to_rich_value_index) = vm_to_rich_value_index.as_ref() {
        dump_vm_to_rv_usage(vm_to_rich_value_index, &usage);
    }

    dump_rich_data_graph(&pkg)?;

    let rv_to_targets = resolve_rich_value_targets(&pkg)?;

    dump_rich_value_image_targets(vm_to_rich_value_index.as_ref(), cell_images.as_ref(), &rv_to_targets)?;

    dump_vm_cell_mappings(&pkg, vm_to_rich_value_index.as_ref(), &rv_to_targets);

    if extract_cell_images {
        dump_rich_cell_images_by_cell(&pkg, extract_cell_images_out);
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn print_interesting_parts(pkg: &XlsxPackage) {
    let mut parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| {
            let normalized = name.strip_prefix('/').unwrap_or(name);
            let lower = normalized.to_ascii_lowercase();
            lower.starts_with("xl/richdata/")
                || lower.starts_with("xl/metadata.xml")
                || lower == "xl/_rels/metadata.xml.rels"
                || lower.starts_with("xl/cellimages")
                || (lower.starts_with("xl/_rels/cellimages") && lower.ends_with(".rels"))
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
    let Some((part_name, bytes)) = find_part_case_insensitive(pkg, "xl/metadata.xml") else {
        println!();
        println!("xl/metadata.xml: (missing)");
        return None;
    };

    println!();
    println!("{part_name}:");

    match parse_metadata_xml(&bytes) {
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

    match parse_value_metadata_vm_to_rich_value_index_map(&bytes) {
        Ok(primary) => {
            if primary.is_empty() {
                println!("  vm -> rich_value_index: (none resolved)");
                return Some(primary);
            }

            // Excel's `vm` appears in the wild as both 0-based and 1-based. To be tolerant, insert
            // both the original key and its 0-based equivalent.
            let primary_len = primary.len();
            let mut map: HashMap<u32, u32> = HashMap::with_capacity(primary_len.saturating_mul(2));
            for (vm, rv) in primary {
                map.entry(vm).or_insert(rv);
                if vm > 0 {
                    map.entry(vm - 1).or_insert(rv);
                }
            }

            println!("  vm -> rich_value_index ({}):", map.len());
            if map.len() != primary_len {
                println!(
                    "    (tolerant mapping includes both vm and vm-1 keys; primary entries: {primary_len})"
                );
            }

            let mut entries: Vec<(u32, u32)> = map.iter().map(|(vm, rv)| (*vm, *rv)).collect();
            entries.sort_by_key(|(vm, _)| *vm);

            let max = 50usize;
            for (idx, (vm, rv)) in entries.iter().enumerate() {
                if idx >= max {
                    println!("    ... ({} more)", entries.len().saturating_sub(max));
                    break;
                }
                println!("    vm {vm} -> rv {rv}");
            }

            Some(map)
        }
        Err(err) => {
            println!("  vm -> rich_value_index: (failed to resolve: {err})");
            None
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_metadata_relationships(pkg: &XlsxPackage) {
    let rels_part = openxml::rels_part_name("xl/metadata.xml");
    let Some((part_name, bytes)) = find_part_case_insensitive(pkg, &rels_part) else {
        return;
    };

    println!();
    println!("{part_name}:");

    let relationships = match openxml::parse_relationships(&bytes) {
        Ok(r) => r,
        Err(err) => {
            println!("  (failed to parse relationships: {err})");
            return;
        }
    };

    if relationships.is_empty() {
        println!("  (no relationships)");
        return;
    }

    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            println!("  {} [{}] -> {} (external)", rel.id, rel.type_uri, rel.target);
            continue;
        }
        let resolved = openxml::resolve_target("xl/metadata.xml", &rel.target);
        println!(
            "  {} [{}] -> {} (resolved: {})",
            rel.id, rel.type_uri, rel.target, resolved
        );
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
                    println!(
                        "      [{idx}] {} -> {}",
                        img.embed_rel_id,
                        img.target.as_deref().unwrap_or("<unresolved>")
                    );
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
                    if openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"c") {
                        sheet_total_cells += 1;
                        let mut has_vm = false;
                        let mut has_cm = false;
                        for attr in e.attributes().with_checks(false) {
                            let attr = match attr {
                                Ok(attr) => attr,
                                Err(_) => continue,
                            };
                            let key = openxml::local_name(attr.key.as_ref());
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
        .filter(|name| {
            let normalized = name.strip_prefix('/').unwrap_or(name);
            let lower = normalized.to_ascii_lowercase();
            lower.starts_with("xl/richdata/") && lower.ends_with(".xml")
        })
        .collect();
    rich_parts.sort();
    if rich_parts.is_empty() {
        return Ok(());
    }

    println!();
    println!("richData part graph (relationships):");

    for part in rich_parts {
        let rels_part = openxml::rels_part_name(part);
        let Some((_rels_name, rels_bytes)) = find_part_case_insensitive(pkg, &rels_part) else {
            println!("  {part}: (no rels part)");
            continue;
        };

        let relationships = match openxml::parse_relationships(&rels_bytes) {
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
            let resolved = openxml::resolve_target(part, &rel.target);
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
                    if openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"sheetData") =>
                {
                    in_sheet_data = true;
                }
                Ok(quick_xml::events::Event::End(e))
                    if openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"sheetData") =>
                {
                    in_sheet_data = false;
                }
                Ok(quick_xml::events::Event::Start(e))
                | Ok(quick_xml::events::Event::Empty(e))
                    if in_sheet_data
                        && openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"c") =>
                {
                    let mut r: Option<String> = None;
                    let mut vm: Option<u32> = None;
                    for attr in e.attributes().with_checks(false) {
                        let Ok(attr) = attr else {
                            continue;
                        };
                        let key = openxml::local_name(attr.key.as_ref());
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

    let mut rich_value_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| rich_value_part_sort_key(name).is_some())
        .collect();
    rich_value_parts.sort_by(|a, b| {
        rich_value_part_sort_key(a)
            .cmp(&rich_value_part_sort_key(b))
            .then_with(|| a.cmp(b))
    });

    // Best-effort table of rich value index -> first "relationship index" payload found in the
    // rich value record.
    let mut rv_to_rel_values: BTreeMap<u32, Vec<u32>> = BTreeMap::new();
    let mut rv_idx: u32 = 0;
    for part_name in rich_value_parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let Ok(rel_indices) = rich_value::parse_rich_value_relationship_indices(bytes) else {
            continue;
        };
        for rel_idx in rel_indices {
            if let Some(rel_idx) = rel_idx {
                rv_to_rel_values.insert(rv_idx, vec![rel_idx as u32]);
            }
            rv_idx = rv_idx.saturating_add(1);
        }
    }

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
fn dump_rich_cell_images_by_cell(pkg: &XlsxPackage, out_dir: Option<&Path>) {
    let images_by_cell = match pkg.extract_rich_cell_images_by_cell() {
        Ok(v) => v,
        Err(err) => {
            println!();
            println!("rich-data in-cell images (by cell): (failed to extract: {err})");
            return;
        }
    };

    if images_by_cell.is_empty() {
        println!();
        println!("rich-data in-cell images (by cell): (none)");
        return;
    }

    println!();
    println!(
        "rich-data in-cell images (by cell): {} cell(s)",
        images_by_cell.len()
    );

    let mut counts_by_sheet: BTreeMap<&str, usize> = BTreeMap::new();
    for ((sheet, _cell), _bytes) in &images_by_cell {
        *counts_by_sheet.entry(sheet.as_str()).or_insert(0) += 1;
    }
    if !counts_by_sheet.is_empty() {
        println!("  sheets:");
        for (sheet, count) in counts_by_sheet {
            println!("    {sheet}: {count}");
        }
    }

    let mut entries: Vec<(&(String, formula_model::CellRef), &Vec<u8>)> =
        images_by_cell.iter().collect();
    entries.sort_by(|(a_key, _), (b_key, _)| {
        a_key
            .0
            .cmp(&b_key.0)
            .then_with(|| (a_key.1.row, a_key.1.col).cmp(&(b_key.1.row, b_key.1.col)))
    });

    println!("  examples:");
    let max = 20usize;
    for (idx, ((sheet, cell), bytes)) in entries.into_iter().enumerate() {
        if idx >= max {
            println!("    ... ({} more)", images_by_cell.len().saturating_sub(max));
            break;
        }
        println!("    {sheet}!{cell}: {} bytes", bytes.len());
    }

    if let Some(out_dir) = out_dir {
        if let Err(err) = fs::create_dir_all(out_dir) {
            println!();
            println!(
                "rich-data in-cell images (by cell): failed to create output dir {}: {err}",
                out_dir.display()
            );
            return;
        }

        let mut manifest: Vec<String> = Vec::with_capacity(images_by_cell.len() + 1);
        manifest.push("sheet\tcell\tbytes\tfile".to_string());

        let mut written = 0usize;
        let mut failed = 0usize;

        for ((sheet, cell), bytes) in &images_by_cell {
            let sheet_sanitized = sanitize_filename_component(sheet);
            let cell_a1 = cell.to_string();
            let ext = guess_image_extension(bytes).unwrap_or("bin");

            let mut file_name = format!("{sheet_sanitized}_{cell_a1}.{ext}");
            let mut path = out_dir.join(&file_name);
            if path.exists() {
                // Avoid overwriting in case of name collisions.
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
                    manifest.push(format!(
                        "{}\t{}\t{}\t{}",
                        tsv_escape(sheet),
                        cell_a1,
                        bytes.len(),
                        file_name
                    ));
                }
                Err(_) => failed += 1,
            }
        }

        let manifest_path = out_dir.join("manifest.tsv");
        if let Err(err) = fs::write(&manifest_path, manifest.join("\n") + "\n") {
            println!();
            println!(
                "rich-data in-cell images (by cell): failed to write manifest {}: {err}",
                manifest_path.display()
            );
        }

        println!();
        println!(
            "rich-data in-cell images (by cell): wrote {written} file(s) to {} (failed: {failed})",
            out_dir.display()
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
        let target = part.images[0].target.as_deref().unwrap_or("<unresolved>");
        out.push(format!(
            "{} (via {}[0])",
            target, part.path
        ));
        return Some(out);
    }

    let idx = rv as usize;
    if let Some(img) = part.images.get(idx) {
        let target = img.target.as_deref().unwrap_or("<unresolved>");
        out.push(format!("{} (via {}[{idx}])", target, part.path));
        return Some(out);
    }

    None
}

#[cfg(not(target_arch = "wasm32"))]
fn rich_value_part_sort_key(part_name: &str) -> Option<(u8, u32)> {
    let normalized = part_name.strip_prefix('/').unwrap_or(part_name);
    let lower = normalized.to_ascii_lowercase();
    if !lower.starts_with("xl/richdata/") || !lower.ends_with(".xml") {
        return None;
    }

    let file = lower.rsplit('/').next()?;
    let stem = file.strip_suffix(".xml")?;

    // See `crates/formula-xlsx/src/rich_data/mod.rs` for richer logic. This is a minimal
    // best-effort sorter for the debug CLI.
    let (family, suffix) = if let Some(rest) = stem.strip_prefix("richvalue") {
        (0u8, rest)
    } else if let Some(rest) = stem.strip_prefix("rdrichvalue") {
        (1u8, rest)
    } else {
        return None;
    };

    let idx = if suffix.is_empty() {
        0
    } else if suffix.chars().all(|c| c.is_ascii_digit()) {
        suffix.parse::<u32>().ok()?
    } else {
        return None;
    };

    Some((family, idx))
}

#[cfg(not(target_arch = "wasm32"))]
fn resolve_rich_value_targets(pkg: &XlsxPackage) -> Result<BTreeMap<u32, Vec<String>>, Box<dyn Error>> {
    let mut out: BTreeMap<u32, Vec<String>> = BTreeMap::new();

    // Scan all `xl/richData/richValue*.xml` parts (and forward-compatible variants like
    // `rdrichvalue*.xml`) in numeric order.
    let mut rich_value_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| rich_value_part_sort_key(name).is_some())
        .collect();
    rich_value_parts.sort_by(|a, b| {
        rich_value_part_sort_key(a)
            .cmp(&rich_value_part_sort_key(b))
            .then_with(|| a.cmp(b))
    });

    // Track the per-rich-value relationship index so we can resolve via `richValueRel.xml` later.
    let mut rel_indices_by_rv: Vec<Option<usize>> = Vec::new();

    let mut rv_base: u32 = 0;
    for part_name in rich_value_parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };

        let part_rel_indices = match rich_value::parse_rich_value_relationship_indices(bytes) {
            Ok(v) => v,
            Err(err) => {
                println!("  warning: failed to parse rich value indices in {part_name}: {err}");
                continue;
            }
        };
        let part_rv_count = part_rel_indices.len() as u32;

        // Resolve any embedded relationship IDs (e.g. `r:embed="rIdN"`) via this part's `.rels`
        // file, if present.
        let rels_part = openxml::rels_part_name(part_name);
        let rel_by_id: HashMap<String, openxml::Relationship> =
            if let Some((_rels_name, rels_bytes)) = find_part_case_insensitive(pkg, &rels_part) {
                match openxml::parse_relationships(&rels_bytes) {
                    Ok(rels) => rels.into_iter().map(|r| (r.id.clone(), r)).collect(),
                    Err(err) => {
                        println!("  warning: failed to parse {rels_part}: {err}");
                        HashMap::new()
                    }
                }
            } else {
                HashMap::new()
            };

        if !rel_by_id.is_empty() {
            if let Ok(rv_ids) = extract_rich_value_record_rids(bytes) {
                for (local_rv, rids) in rv_ids {
                    let rv = rv_base.saturating_add(local_rv);
                    for rid in rids {
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
                        out.entry(rv)
                            .or_default()
                            .push(openxml::resolve_target(part_name, &rel.target));
                    }
                }
            }
        }

        rel_indices_by_rv.extend(part_rel_indices);
        rv_base = rv_base.saturating_add(part_rv_count);
    }

    // Resolve relationship indices via `xl/richData/richValueRel.xml` when present.
    if !rel_indices_by_rv.is_empty() {
        if let Some((rel_part_name, rel_xml_bytes)) =
            find_part_case_insensitive(pkg, "xl/richData/richValueRel.xml")
        {
            let rel_id_table = match rich_value_rel::parse_rich_value_rel_table(&rel_xml_bytes) {
                Ok(v) => v,
                Err(err) => {
                    println!("  warning: failed to parse {rel_part_name}: {err}");
                    Vec::new()
                }
            };

            let rels_part = openxml::rels_part_name(&rel_part_name);
            let targets_by_id: HashMap<String, String> =
                if let Some((_rels_name, rels_bytes)) = find_part_case_insensitive(pkg, &rels_part)
                {
                    match openxml::parse_relationships(&rels_bytes) {
                        Ok(rels) => {
                            let mut out = HashMap::new();
                            for rel in rels {
                                if rel
                                    .target_mode
                                    .as_deref()
                                    .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
                                {
                                    continue;
                                }
                                let resolved = openxml::resolve_target(&rel_part_name, &rel.target);
                                out.insert(rel.id, resolved);
                            }
                            out
                        }
                        Err(err) => {
                            println!("  warning: failed to parse {rels_part}: {err}");
                            HashMap::new()
                        }
                    }
                } else {
                    HashMap::new()
                };

            if !rel_id_table.is_empty() && !targets_by_id.is_empty() {
                for (rv_idx, rel_idx) in rel_indices_by_rv.iter().enumerate() {
                    let Some(rel_idx) = rel_idx else {
                        continue;
                    };
                    let Some(rid) = rel_id_table.get(*rel_idx) else {
                        continue;
                    };
                    if rid.is_empty() {
                        continue;
                    }
                    if let Some(target) = targets_by_id.get(rid) {
                        out.entry(rv_idx as u32).or_default().push(target.clone());
                    }
                }

                // If `richValue.xml` did not contain explicit relationship indices, Excel-like
                // producers sometimes implicitly align rich value record indices with
                // `richValueRel.xml` relationship indices. As a best-effort fallback, if we didn't
                // resolve anything above, assume `rv == rel_idx`.
                if out.is_empty() && rel_indices_by_rv.iter().all(|v| v.is_none()) {
                    for rv_idx in 0..rel_indices_by_rv.len() {
                        let Some(rid) = rel_id_table.get(rv_idx) else {
                            continue;
                        };
                        if rid.is_empty() {
                            continue;
                        }
                        if let Some(target) = targets_by_id.get(rid) {
                            out.entry(rv_idx as u32).or_default().push(target.clone());
                        }
                    }
                }
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
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
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
    let path = path.strip_prefix('/').unwrap_or(path);
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
    let path = path.strip_prefix('/').unwrap_or(path);
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
fn sanitize_filename_component(value: &str) -> String {
    // Avoid path separators and other awkward filename chars across platforms.
    // Keep it readable for debugging.
    let mut out = String::with_capacity(value.len().min(64));
    for ch in value.chars() {
        if out.len() >= 64 {
            break;
        }
        match ch {
            'a'..='z' | 'A'..='Z' | '0'..='9' | '-' | '_' => out.push(ch),
            ' ' => out.push('_'),
            _ => out.push('_'),
        }
    }
    if out.is_empty() {
        out.push_str("sheet");
    }
    out
}

#[cfg(not(target_arch = "wasm32"))]
fn guess_image_extension(bytes: &[u8]) -> Option<&'static str> {
    // Common image signatures used by Excel media parts.
    const PNG: &[u8] = b"\x89PNG\r\n\x1a\n";
    if bytes.starts_with(PNG) {
        return Some("png");
    }
    if bytes.len() >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF {
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
fn tsv_escape(value: &str) -> String {
    // Sheet names are user-controlled; keep the manifest format robust.
    value.replace(['\t', '\r', '\n'], " ")
}

#[cfg(not(target_arch = "wasm32"))]
fn find_part_case_insensitive(pkg: &XlsxPackage, desired: &str) -> Option<(String, Vec<u8>)> {
    let desired = desired.strip_prefix('/').unwrap_or(desired);
    let desired_lower = desired.to_ascii_lowercase();
    for name in pkg.part_names() {
        let normalized = name.strip_prefix('/').unwrap_or(name);
        if normalized.to_ascii_lowercase() == desired_lower {
            return pkg.part(name).map(|b| (name.to_string(), b.to_vec()));
        }
    }
    None
}
