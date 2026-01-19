#![cfg_attr(target_arch = "wasm32", allow(dead_code, unused_imports))]

use std::collections::BTreeSet;
use std::error::Error;
use std::io::{self, Write};
#[cfg(not(target_arch = "wasm32"))]
use std::fs;
use std::path::{Path, PathBuf};

use formula_xlsx::{openxml, XlsxPackage};
use serde::Serialize;

#[cfg(not(target_arch = "wasm32"))]
mod workbook_open;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct WorkbookCellImagesDump {
    workbook: String,

    cell_images_part: Option<String>,
    cell_images_rels_part: Option<String>,

    content_type_override: Option<String>,
    workbook_relationship_type_uri: Option<String>,

    root_element_local_name: Option<String>,
    root_element_namespace_uri: Option<String>,

    blip_embed_count: Option<usize>,
    cell_images_rels_relationship_type_uris: Option<Vec<String>>,
}

fn usage() -> &'static str {
    "dump_cell_images [--password <pw>] <path.xlsx|dir> [--print-xml] [--json]\n\
\n\
Inspects `xl/cellimages.xml` and related parts in one or more workbooks.\n\
\n\
Options:\n\
  --password <pw>  Password for Office-encrypted workbooks (OLE `EncryptedPackage`; use --password '' for empty password)\n\
  --print-xml  Print the raw `cellimages.xml` (UTF-8 only)\n\
  --json       Emit one JSON object per workbook (one per line)\n\
"
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
    let mut input_path: Option<PathBuf> = None;
    let mut password: Option<String> = None;
    let mut print_xml = false;
    let mut json = false;

    let mut args = std::env::args().skip(1).peekable();
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--help" | "-h" => {
                eprintln!("{}", usage());
                return Ok(());
            }
            "--print-xml" => {
                print_xml = true;
            }
            "--json" => {
                json = true;
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
                if input_path.is_some() {
                    return Err(format!("unexpected extra argument: {value}\n\n{}", usage()).into());
                }
                input_path = Some(PathBuf::from(value));
            }
        }
    }

    let input_path =
        input_path.ok_or_else(|| format!("missing <path.xlsx|dir>\n\n{}", usage()))?;

    let inputs = if input_path.is_dir() {
        collect_xlsx_files(&input_path)?
    } else {
        vec![input_path]
    };

    for xlsx_path in inputs {
        if let Err(err) = dump_one(&xlsx_path, password.as_deref(), print_xml, json) {
            eprintln!("{}: {err}", xlsx_path.display());
        }
    }

    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn dump_one(
    xlsx_path: &Path,
    password: Option<&str>,
    print_xml: bool,
    json: bool,
) -> Result<(), Box<dyn Error>> {
    let pkg = workbook_open::open_xlsx_package(xlsx_path, password)?;
    let workbook = xlsx_path.display().to_string();

    let dump = inspect_package(&pkg, workbook);

    let mut out = io::BufWriter::new(io::stdout());

    if json {
        writeln!(out, "{}", serde_json::to_string(&dump)?)?;
        // Keep JSON output machine-readable. If the caller also asked for XML, put it on stderr.
        if print_xml {
            print_xml_to_stderr(&pkg, &dump);
        }
    } else {
        print_human(&pkg, &dump, print_xml, &mut out)?;
    }

    out.flush()?;
    Ok(())
}

fn inspect_package(pkg: &XlsxPackage, workbook: String) -> WorkbookCellImagesDump {
    let part_info = match pkg.cell_images_part_info() {
        Ok(Some(info)) => Some((info.part_path, info.rels_path)),
        Ok(None) => None,
        // Best-effort: if the library helper fails (malformed file), still attempt to inspect the
        // canonical part location if present.
        Err(_) => fallback_cell_images_part(pkg),
    };

    let Some((part_path, rels_path)) = part_info else {
        return WorkbookCellImagesDump {
            workbook,
            cell_images_part: None,
            cell_images_rels_part: None,
            content_type_override: None,
            workbook_relationship_type_uri: None,
            root_element_local_name: None,
            root_element_namespace_uri: None,
            blip_embed_count: None,
            cell_images_rels_relationship_type_uris: None,
        };
    };

    let content_type_override = content_type_override_for_part(pkg, &part_path);
    let workbook_relationship_type_uri = discover_workbook_relationship_type_uri(pkg, &part_path);

    let (root_local_name, root_namespace_uri, blip_embed_count) =
        parse_cell_images_root_and_blips(pkg, &part_path);

    let cell_images_rels_relationship_type_uris = parse_rels_type_uris(pkg, &rels_path);

    WorkbookCellImagesDump {
        workbook,
        cell_images_part: Some(part_path),
        cell_images_rels_part: Some(rels_path),
        content_type_override,
        workbook_relationship_type_uri,
        root_element_local_name: root_local_name,
        root_element_namespace_uri: root_namespace_uri,
        blip_embed_count,
        cell_images_rels_relationship_type_uris,
    }
}

fn print_human(
    pkg: &XlsxPackage,
    dump: &WorkbookCellImagesDump,
    print_xml: bool,
    out: &mut dyn Write,
) -> io::Result<()> {
    writeln!(out, "workbook: {}", dump.workbook)?;

    let Some(cell_images_part) = dump.cell_images_part.as_deref() else {
        writeln!(out, "  no cell images part")?;
        writeln!(out)?;
        return Ok(());
    };

    writeln!(out, "  cell images part: {cell_images_part}")?;
    writeln!(
        out,
        "  cell images rels: {}",
        dump.cell_images_rels_part
            .as_deref()
            .unwrap_or("unknown")
    )?;
    writeln!(
        out,
        "  content type override: {}",
        dump.content_type_override
            .as_deref()
            .unwrap_or("unknown")
    )?;
    writeln!(
        out,
        "  workbook relationship type: {}",
        dump.workbook_relationship_type_uri
            .as_deref()
            .unwrap_or("unknown")
    )?;
    writeln!(
        out,
        "  root element: {} namespace={}",
        dump.root_element_local_name
            .as_deref()
            .unwrap_or("unknown"),
        dump.root_element_namespace_uri
            .as_deref()
            .unwrap_or("unknown")
    )?;
    writeln!(
        out,
        "  a:blip embeds: {}",
        dump.blip_embed_count
            .map(|n| n.to_string())
            .as_deref()
            .unwrap_or("unknown")
    )?;
    writeln!(
        out,
        "  rels relationship types: {}",
        dump.cell_images_rels_relationship_type_uris
            .as_deref()
            .map(|v| v.join(", "))
            .unwrap_or_else(|| "unknown".to_string())
    )?;

    if print_xml {
        if let Some(bytes) = pkg.part(cell_images_part) {
            match std::str::from_utf8(bytes) {
                Ok(xml) => {
                    writeln!(out)?;
                    writeln!(out, "cellimages.xml:\n{xml}")?;
                    writeln!(out)?;
                }
                Err(err) => {
                    writeln!(
                        out,
                        "  warning: cellimages.xml is not UTF-8; cannot print ({err})"
                    )?;
                    writeln!(out)?;
                }
            }
        }
    } else {
        writeln!(out)?;
    }

    Ok(())
}

fn print_xml_to_stderr(pkg: &XlsxPackage, dump: &WorkbookCellImagesDump) {
    let Some(cell_images_part) = dump.cell_images_part.as_deref() else {
        return;
    };

    let Some(bytes) = pkg.part(cell_images_part) else {
        return;
    };

    match std::str::from_utf8(bytes) {
        Ok(xml) => {
            eprintln!("\nworkbook: {}", dump.workbook);
            eprintln!("cellimages.xml:\n{xml}\n");
        }
        Err(err) => {
            eprintln!(
                "workbook: {}: warning: cellimages.xml is not UTF-8; cannot print ({err})",
                dump.workbook
            );
        }
    }
}

fn content_type_override_for_part(pkg: &XlsxPackage, part_path: &str) -> Option<String> {
    let ct_bytes = pkg.part("[Content_Types].xml")?;
    let ct_xml = std::str::from_utf8(ct_bytes).ok()?;

    let doc = roxmltree::Document::parse(ct_xml).ok()?;
    let needle = if part_path.starts_with('/') {
        part_path.to_string()
    } else {
        format!("/{part_path}")
    };

    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Override" {
            continue;
        }
        if node.attribute("PartName") != Some(needle.as_str()) {
            continue;
        }
        if let Some(content_type) = node.attribute("ContentType") {
            return Some(content_type.to_string());
        }
    }

    None
}

fn discover_workbook_relationship_type_uri(pkg: &XlsxPackage, cell_images_part: &str) -> Option<String> {
    let rels_bytes = pkg.part("xl/_rels/workbook.xml.rels")?;
    let rels_xml = std::str::from_utf8(rels_bytes).ok()?;
    let doc = roxmltree::Document::parse(rels_xml).ok()?;

    let needle = cell_images_part.strip_prefix('/').unwrap_or(cell_images_part);
    let mut found = BTreeSet::new();

    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            continue;
        }
        let Some(target) = node.attribute("Target") else {
            continue;
        };
        let Some(type_uri) = node.attribute("Type") else {
            continue;
        };

        let resolved = openxml::resolve_target("xl/workbook.xml", target);
        if resolved == needle {
            found.insert(type_uri.to_string());
        }
    }

    found.into_iter().next()
}

fn parse_cell_images_root_and_blips(
    pkg: &XlsxPackage,
    part_path: &str,
) -> (Option<String>, Option<String>, Option<usize>) {
    let bytes = match pkg.part(part_path) {
        Some(bytes) => bytes,
        None => return (None, None, None),
    };
    let xml = match std::str::from_utf8(bytes) {
        Ok(xml) => xml,
        Err(_) => return (None, None, None),
    };

    let doc = match roxmltree::Document::parse(xml) {
        Ok(doc) => doc,
        Err(_) => return (None, None, None),
    };

    let root = doc.root_element();
    let root_local_name = Some(root.tag_name().name().to_string());
    let root_namespace_uri = root.tag_name().namespace().map(|ns| ns.to_string());

    let mut embed_count = 0usize;
    for node in root.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "blip" {
            continue;
        }

        // Count only embeds (not links).
        if node.attribute((REL_NS, "embed")).or_else(|| node.attribute("r:embed")).is_some() {
            embed_count += 1;
        }
    }

    (root_local_name, root_namespace_uri, Some(embed_count))
}

fn parse_rels_type_uris(pkg: &XlsxPackage, rels_part_path: &str) -> Option<Vec<String>> {
    let bytes = pkg.part(rels_part_path)?;
    let xml = std::str::from_utf8(bytes).ok()?;
    let doc = roxmltree::Document::parse(xml).ok()?;

    let mut out: BTreeSet<String> = BTreeSet::new();
    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            continue;
        }
        if let Some(type_uri) = node.attribute("Type") {
            out.insert(type_uri.to_string());
        }
    }

    Some(out.into_iter().collect())
}

fn fallback_cell_images_part(pkg: &XlsxPackage) -> Option<(String, String)> {
    // Prefer the canonical Excel part location when present.
    if pkg.part("xl/cellimages.xml").is_some() {
        return Some((
            "xl/cellimages.xml".to_string(),
            "xl/_rels/cellimages.xml.rels".to_string(),
        ));
    }

    // Otherwise scan `xl/` for `cellimages*.xml` parts (allowing numeric suffix).
    for name in pkg.part_names() {
        let lower = name.to_ascii_lowercase();
        if !lower.starts_with("xl/") || !lower.ends_with(".xml") {
            continue;
        }
        let file_name = lower.rsplit('/').next().unwrap_or("");
        let Some(stem) = file_name.strip_suffix(".xml") else {
            continue;
        };
        let Some(suffix) = stem.strip_prefix("cellimages") else {
            continue;
        };
        if !suffix.is_empty() && !suffix.chars().all(|c| c.is_ascii_digit()) {
            continue;
        }
        return Some((name.to_string(), rels_for_part(name)));
    }

    None
}

fn rels_for_part(part_name: &str) -> String {
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
    let (dir, file) = part_name.rsplit_once('/').unwrap_or(("", part_name));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn collect_xlsx_files(dir: &Path) -> Result<Vec<PathBuf>, Box<dyn Error>> {
    let mut out: Vec<PathBuf> = fs::read_dir(dir)?
        .filter_map(|entry| entry.ok())
        .map(|entry| entry.path())
        .filter(|path| path.extension().and_then(|s| s.to_str()) == Some("xlsx"))
        .collect();
    out.sort();
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn inspects_synthetic_cellimages_part() {
        let ct_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.example.cellimages+xml"/>
</Types>"#;

        let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCellImages" Type="http://schemas.example.com/relationships/cellImages" Target="cellimages.xml"/>
</Relationships>"#;

        let cell_images_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ci:cellImages xmlns:ci="http://schemas.example.com/cellimages"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <a:blip r:embed="rId1"/>
  <a:blip r:link="rId2"/>
</ci:cellImages>"#;

        let cell_images_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.example.com/relationships/other" Target="media/other.bin"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("[Content_Types].xml", ct_xml),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/cellimages.xml", cell_images_xml),
            ("xl/_rels/cellimages.xml.rels", cell_images_rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let dump = inspect_package(&pkg, "synthetic.xlsx".to_string());
        assert_eq!(dump.cell_images_part.as_deref(), Some("xl/cellimages.xml"));
        assert_eq!(
            dump.cell_images_rels_part.as_deref(),
            Some("xl/_rels/cellimages.xml.rels")
        );
        assert_eq!(
            dump.content_type_override.as_deref(),
            Some("application/vnd.example.cellimages+xml")
        );
        assert_eq!(
            dump.workbook_relationship_type_uri.as_deref(),
            Some("http://schemas.example.com/relationships/cellImages")
        );
        assert_eq!(dump.root_element_local_name.as_deref(), Some("cellImages"));
        assert_eq!(
            dump.root_element_namespace_uri.as_deref(),
            Some("http://schemas.example.com/cellimages")
        );
        assert_eq!(dump.blip_embed_count, Some(1));
        let expected_rels_types = vec![
            "http://schemas.example.com/relationships/other".to_string(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
        ];
        assert_eq!(
            dump.cell_images_rels_relationship_type_uris.as_deref(),
            Some(expected_rels_types.as_slice())
        );
    }
}
