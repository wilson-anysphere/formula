//! Developer utility to inspect Excel RichData / RichValue parts.
//!
//! Usage:
//!   cargo run -p formula-xlsx --example dump_richvalue -- path/to/file.xlsx

use std::collections::HashMap;
use std::error::Error;
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
use std::path::PathBuf;

use formula_xlsx::{openxml, parse_rich_value_structure_xml, parse_rich_value_types_xml, XlsxPackage};
use roxmltree::Document;

const EXPECTED_PARTS: &[&str] = &[
    "xl/richData/richValue.xml",
    "xl/richData/richValueRel.xml",
    "xl/richData/richValueTypes.xml",
    "xl/richData/richValueStructure.xml",
];

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug)]
struct DumpError(String);

impl std::fmt::Display for DumpError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl Error for DumpError {}

fn usage() -> &'static str {
    "dump_richvalue <path.xlsx>\n\
\n\
Prints a compact, copy/pastable summary of Excel RichData parts:\n\
  xl/richData/richValue.xml\n\
  xl/richData/richValueRel.xml (+ its .rels, for target resolution)\n\
  xl/richData/richValueTypes.xml\n\
  xl/richData/richValueStructure.xml\n"
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
fn main() -> Result<(), Box<dyn Error>> {
    let xlsx_path = parse_args()?;
    let bytes = fs::read(&xlsx_path)?;
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    println!("workbook: {}", xlsx_path.display());

    let mut missing_parts = Vec::new();
    let mut present_any = false;

    println!("parts:");
    for part in EXPECTED_PARTS {
        let present = pkg.part(part).is_some();
        present_any |= present;
        if present {
            println!("  {part}: present");
        } else {
            println!("  {part}: missing");
            missing_parts.push(*part);
        }
    }

    if !present_any {
        println!();
        println!("No RichData parts found in this workbook.");
        println!("If the file is expected to contain images-in-cell / RichData, please verify it was saved by a recent version of Excel.");
        println!();
        print_expected_paths();
        return Ok(());
    }

    if !missing_parts.is_empty() {
        println!();
        println!("Note: some RichData parts are missing.");
        print_expected_paths();
    }

    let (types_count, types_parse_err) = match pkg.part(EXPECTED_PARTS[2]) {
        None => (None, None),
        Some(bytes) => match parse_rich_value_types_xml(bytes) {
            Ok(types) => (Some(types.len()), None),
            Err(err) => (None, Some(err.to_string())),
        },
    };

    let (structures_count, structures_parse_err) = match pkg.part(EXPECTED_PARTS[3]) {
        None => (None, None),
        Some(bytes) => match parse_rich_value_structure_xml(bytes) {
            Ok(structures) => (Some(structures.len()), None),
            Err(err) => (None, Some(err.to_string())),
        },
    };

    let rich_value_parts = find_rich_value_parts(&pkg);
    let (values_count, values_parse_err) = if rich_value_parts.is_empty() {
        (None, None)
    } else {
        match count_rich_value_entries(&pkg, &rich_value_parts) {
            Ok(count) => (Some(count), None),
            Err(err) => (None, Some(err.to_string())),
        }
    };

    let (rel_entries, rel_parse_err) = match pkg.part(EXPECTED_PARTS[1]) {
        None => (None, None),
        Some(bytes) => match parse_rich_value_rel(bytes, EXPECTED_PARTS[1]) {
            Ok(v) => (Some(v), None),
            Err(err) => (None, Some(err)),
        },
    };

    println!();
    println!("summary:");
    print_count("richValueTypes", types_count, types_parse_err.as_deref());
    print_count(
        "richValueStructure",
        structures_count,
        structures_parse_err.as_deref(),
    );
    match (values_count, values_parse_err.as_deref()) {
        (Some(n), _) => println!("  richValue: {n} (parts: {})", rich_value_parts.len()),
        (None, Some(_)) => println!("  richValue: <parse error>"),
        (None, None) => println!("  richValue: missing"),
    }
    match &rel_entries {
        None if pkg.part(EXPECTED_PARTS[1]).is_none() => println!("  richValueRel: missing"),
        None => println!("  richValueRel: <parse error>"),
        Some(entries) => println!("  richValueRel: {}", entries.len()),
    }

    if let Some(err) = types_parse_err {
        println!("  richValueTypes parse error: {err}");
    }
    if let Some(err) = structures_parse_err {
        println!("  richValueStructure parse error: {err}");
    }
    if let Some(err) = values_parse_err {
        println!("  richValue parse error: {err}");
    }
    if let Some(err) = rel_parse_err {
        println!("  richValueRel parse error: {err}");
    }

    if pkg.part(EXPECTED_PARTS[0]).is_none() && !rich_value_parts.is_empty() {
        println!();
        println!("Found richValue parts (note: xl/richData/richValue.xml was missing):");
        for part in &rich_value_parts {
            println!("  - {part}");
        }
    } else if rich_value_parts.len() > 1 {
        println!();
        println!("Found richValue parts:");
        for part in &rich_value_parts {
            println!("  - {part}");
        }
    }

    // richValueRel.xml.rels is needed to resolve rId* -> target part.
    let rich_value_rel_rels_part = rels_for_part(EXPECTED_PARTS[1]);
    let rels_target_map = match pkg.part(&rich_value_rel_rels_part) {
        None => None,
        Some(bytes) => Some(parse_relationship_targets(
            bytes,
            &rich_value_rel_rels_part,
            EXPECTED_PARTS[1],
        )?),
    };

    if pkg.part(EXPECTED_PARTS[1]).is_some() {
        println!(
            "  richValueRel.rels: {}",
            if rels_target_map.is_some() {
                "present"
            } else {
                "missing"
            }
        );
        println!("  richValueRel.rels part: {rich_value_rel_rels_part}");
    }

    if let Some(entries) = rel_entries {
        println!();
        println!("richValueRel entries:");
        for entry in entries {
            let idx = entry.index.unwrap_or_else(|| entry.ordinal.to_string());
            let rid = entry
                .rel_id
                .clone()
                .unwrap_or_else(|| "<missing r:id>".to_string());
            let resolved = entry.rel_id.as_deref().and_then(|rid| {
                rels_target_map
                    .as_ref()
                    .and_then(|m| m.get(rid))
                    .map(|t| t.as_str())
            });

            match resolved {
                Some(target) => println!("  [{idx}] {rid} -> {target}"),
                None => println!("  [{idx}] {rid} -> <unresolved>"),
            }
        }

        if rels_target_map.is_none() {
            println!();
            println!(
                "Note: unable to resolve richValueRel targets because the .rels part is missing."
            );
        }
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn parse_args() -> Result<PathBuf, Box<dyn Error>> {
    let mut args = std::env::args().skip(1);
    let Some(arg) = args.next() else {
        return Err(Box::new(DumpError(format!(
            "missing <path.xlsx>\n\n{}",
            usage()
        ))));
    };
    if arg == "--help" || arg == "-h" {
        eprintln!("{}", usage());
        std::process::exit(0);
    }
    if args.next().is_some() {
        return Err(Box::new(DumpError(format!(
            "unexpected extra arguments\n\n{}",
            usage()
        ))));
    }
    Ok(PathBuf::from(arg))
}

fn print_expected_paths() {
    println!("Expected part paths:");
    for part in EXPECTED_PARTS {
        println!("  - {part}");
    }
}

fn print_count(label: &str, count: Option<usize>, parse_error: Option<&str>) {
    match (count, parse_error) {
        (Some(n), _) => println!("  {label}: {n}"),
        (None, Some(_)) => println!("  {label}: <parse error>"),
        (None, None) => println!("  {label}: missing"),
    }
}

fn find_rich_value_parts(pkg: &XlsxPackage) -> Vec<String> {
    // Value parts are `xl/richData/richValue.xml` or `xl/richData/richValue<N>.xml`.
    //
    // Avoid `richValueTypes.xml` / `richValueStructure.xml` / `richValueRel.xml`.
    pkg.part_names()
        .filter(|name| is_rich_value_part_name(name))
        .map(|s| s.to_string())
        .collect()
}

fn is_rich_value_part_name(name: &str) -> bool {
    const PREFIX: &str = "xl/richData/richValue";
    const SUFFIX: &str = ".xml";
    if !name.starts_with(PREFIX) || !name.ends_with(SUFFIX) {
        return false;
    }
    if name == "xl/richData/richValueRel.xml"
        || name == "xl/richData/richValueTypes.xml"
        || name == "xl/richData/richValueStructure.xml"
    {
        return false;
    }
    let mid = &name[PREFIX.len()..name.len() - SUFFIX.len()];
    mid.chars().all(|c| c.is_ascii_digit())
}

fn count_rich_value_entries(pkg: &XlsxPackage, parts: &[String]) -> Result<usize, DumpError> {
    let mut total = 0usize;
    for part_name in parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        total += count_rv_elements(bytes, part_name)?;
    }
    Ok(total)
}

fn count_rv_elements(xml_bytes: &[u8], part_name: &str) -> Result<usize, DumpError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| DumpError(format!("{part_name}: not UTF-8: {e}")))?;
    let doc = Document::parse(xml)
        .map_err(|e| DumpError(format!("{part_name}: XML parse error: {e}")))?;
    Ok(doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        .count())
}

#[cfg(test)]
mod rich_value_part_name_tests {
    use super::*;

    #[test]
    fn rich_value_part_name_filters() {
        assert!(is_rich_value_part_name("xl/richData/richValue.xml"));
        assert!(is_rich_value_part_name("xl/richData/richValue1.xml"));
        assert!(!is_rich_value_part_name("xl/richData/richValueTypes.xml"));
        assert!(!is_rich_value_part_name(
            "xl/richData/richValueStructure.xml"
        ));
        assert!(!is_rich_value_part_name("xl/richData/richValueRel.xml"));
    }
}

#[derive(Debug, Clone)]
struct RichValueRelEntry {
    ordinal: usize,
    index: Option<String>,
    rel_id: Option<String>,
}

fn parse_rich_value_rel(
    xml_bytes: &[u8],
    part_name: &str,
) -> Result<Vec<RichValueRelEntry>, DumpError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| DumpError(format!("{part_name}: not UTF-8: {e}")))?;
    let doc = Document::parse(xml)
        .map_err(|e| DumpError(format!("{part_name}: XML parse error: {e}")))?;

    let root = doc.root_element();

    // Prefer direct children of the root (expected shape), but fall back to any
    // descendant element with an r:id attribute.
    let mut entries = Vec::new();
    for (ordinal, node) in root.children().filter(|n| n.is_element()).enumerate() {
        if let Some(rel_id) = rel_id_attr(&node) {
            entries.push(RichValueRelEntry {
                ordinal,
                index: index_attr(&node),
                rel_id: Some(rel_id),
            });
        }
    }

    if entries.is_empty() {
        for (ordinal, node) in root
            .descendants()
            .filter(|n| n.is_element())
            .filter(|n| *n != root)
            .enumerate()
        {
            if let Some(rel_id) = rel_id_attr(&node) {
                entries.push(RichValueRelEntry {
                    ordinal,
                    index: index_attr(&node),
                    rel_id: Some(rel_id),
                });
            }
        }
    }

    Ok(entries)
}

fn rel_id_attr(node: &roxmltree::Node<'_, '_>) -> Option<String> {
    node.attribute((REL_NS, "id"))
        .or_else(|| node.attribute("r:id"))
        .or_else(|| node.attribute("id"))
        .map(str::to_string)
}

fn index_attr(node: &roxmltree::Node<'_, '_>) -> Option<String> {
    for key in ["i", "idx", "index"] {
        if let Some(v) = node.attribute(key) {
            return Some(v.to_string());
        }
    }
    None
}

fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

fn parse_relationship_targets(
    rels_bytes: &[u8],
    rels_part_name: &str,
    source_part: &str,
) -> Result<HashMap<String, String>, DumpError> {
    let xml = std::str::from_utf8(rels_bytes)
        .map_err(|e| DumpError(format!("{rels_part_name}: not UTF-8: {e}")))?;
    let doc = Document::parse(xml)
        .map_err(|e| DumpError(format!("{rels_part_name}: XML parse error: {e}")))?;

    let mut out = HashMap::new();
    for rel in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        if rel
            .attribute("TargetMode")
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let Some(id) = rel.attribute("Id") else {
            continue;
        };
        let Some(target) = rel.attribute("Target") else {
            continue;
        };
        if target.starts_with("http://") || target.starts_with("https://") {
            continue;
        }
        out.insert(id.to_string(), openxml::resolve_target(source_part, target));
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_target_handles_parent_dir() {
        assert_eq!(
            openxml::resolve_target("xl/richData/richValueRel.xml", "../media/image1.png"),
            "xl/media/image1.png"
        );
    }

    #[test]
    fn rels_for_part_matches_opc_convention() {
        assert_eq!(
            rels_for_part("xl/richData/richValueRel.xml"),
            "xl/richData/_rels/richValueRel.xml.rels"
        );
    }
}
