//! Workbook-level "images in cells" (`xl/cellimages.xml`) parsing.
//!
//! Modern Excel features like "Place in Cell" images and the `IMAGE()` function
//! appear to rely on a workbook-level `xl/cellimages.xml` part containing
//! DrawingML `<pic>` payloads that reference media via relationships.

use std::collections::{BTreeMap, HashMap};

use formula_model::drawings::{ImageData, ImageId};
use roxmltree::Document;

use crate::drawings::content_type_for_extension;
use crate::path::resolve_target;
use crate::XlsxError;

type Result<T> = std::result::Result<T, XlsxError>;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// Best-effort loader for Excel "in-cell" images.
///
/// This mirrors the old `cell_images.rs` behavior: missing parts, missing rels, or parse errors
/// should not prevent the workbook from loading. Successfully parsed image relationships are
/// loaded into `workbook.images`.
pub fn load_cell_images_from_parts(parts: &BTreeMap<String, Vec<u8>>, workbook: &mut formula_model::Workbook) {
    for path in parts.keys() {
        if !is_cell_images_part(path) {
            continue;
        }
        let _ = parse_cell_images_part(path, parts, workbook);
    }
}

/// Parsed workbook-level cell images parts.
#[derive(Debug, Clone, Default)]
pub struct CellImages {
    /// The discovered `xl/cellimages*.xml` parts.
    pub parts: Vec<CellImagesPart>,
}

#[derive(Debug, Clone)]
pub struct CellImagesPart {
    /// Part name (e.g. `xl/cellimages.xml`).
    pub path: String,
    /// Part name for the relationships part (e.g. `xl/_rels/cellimages.xml.rels`).
    pub rels_path: String,
    /// Cell images in document order.
    pub images: Vec<CellImageEntry>,
}

#[derive(Debug, Clone)]
pub struct CellImageEntry {
    /// Relationship ID referenced by `<a:blip r:embed="…">` (e.g. `rId1`).
    pub embed_rel_id: String,
    /// Resolved target part path (e.g. `xl/media/image1.png`).
    pub target_path: String,
    /// Workbook image identifier (usually the media file name, like `image1.png`).
    pub image_id: ImageId,
    /// Best-effort raw XML for the `<pic>` subtree for future round-trip support.
    pub pic_xml: Option<String>,
}

/// Best-effort loader for Excel "in-cell" images (`xl/cellimages*.xml`).
///
/// This is intended as a lightweight helper for callers that only need to load
/// referenced media payloads into `workbook.images` during XLSX parsing.
///
/// Missing parts / parse errors are ignored by design.
pub fn load_cell_images_from_parts(parts: &BTreeMap<String, Vec<u8>>, workbook: &mut formula_model::Workbook) {
    let _ = CellImages::parse_from_parts(parts, workbook);
}

impl CellImages {
    /// Discover and parse workbook-level cell images parts (`xl/cellimages*.xml`).
    ///
    /// In addition to returning a structured representation of the part(s), this
    /// loads the referenced `xl/media/*` payloads into `workbook.images`.
    pub fn parse_from_parts(
        parts: &BTreeMap<String, Vec<u8>>,
        workbook: &mut formula_model::Workbook,
    ) -> Result<Self> {
        let mut cell_images_parts = Vec::new();

        // Discover `xl/cellimages*.xml` parts. Excel emits `xl/cellimages.xml` today, but allow
        // a numeric suffix for forward compatibility.
        for path in parts.keys() {
            if !is_cell_images_part(path) {
                continue;
            }

            let part = parse_cell_images_part(path, parts, workbook)?;
            cell_images_parts.push(part);
        }

        Ok(Self {
            parts: cell_images_parts,
        })
    }
}

fn is_cell_images_part(path: &str) -> bool {
    let Some(rest) = path.strip_prefix("xl/") else {
        return false;
    };
    if rest.contains('/') {
        // workbook-level parts live directly under `xl/` (not in `xl/worksheets/`, etc.).
        return false;
    }
    let lower = rest.to_ascii_lowercase();
    lower.starts_with("cellimages") && lower.ends_with(".xml")
}

fn parse_cell_images_part(
    path: &str,
    parts: &BTreeMap<String, Vec<u8>>,
    workbook: &mut formula_model::Workbook,
) -> Result<CellImagesPart> {
    let rels_path = crate::openxml::rels_part_name(path);

    let rels_bytes = parts
        .get(&rels_path)
        .ok_or_else(|| XlsxError::MissingPart(format!("missing cell images rels: {rels_path}")))?;
    let relationships = crate::openxml::parse_relationships(rels_bytes)?;
    let rels_by_id: HashMap<String, crate::openxml::Relationship> = relationships
        .into_iter()
        .map(|rel| (rel.id.clone(), rel))
        .collect();

    let xml_bytes = parts
        .get(path)
        .ok_or_else(|| XlsxError::MissingPart(path.to_string()))?;
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("cell images xml not utf-8: {e}")))?;

    let doc = Document::parse(xml)?;

    let mut images = Vec::new();
    for pic in doc
        .root_element()
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "pic")
    {
        let Some(blip) = pic
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "blip")
        else {
            continue;
        };

        let Some(embed_rel_id) = get_blip_embed_rel_id(&blip) else {
            continue;
        };

        let rel = rels_by_id.get(&embed_rel_id).ok_or_else(|| {
            XlsxError::Invalid(format!(
                "cellimages.xml references missing image relationship {embed_rel_id}"
            ))
        })?;

        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            // External relationships are not backed by OPC parts.
            continue;
        }
        let target_path = resolve_target_best_effort(path, &rels_path, &rel.target, parts)?;

        let file_name = target_path
            .strip_prefix("xl/media/")
            .unwrap_or(&target_path)
            .to_string();
        let image_id = ImageId::new(file_name);

        if workbook.images.get(&image_id).is_none() {
            let bytes = parts
                .get(&target_path)
                .ok_or_else(|| XlsxError::MissingPart(target_path.clone()))?
                .clone();
            let ext = image_id
                .as_str()
                .rsplit_once('.')
                .map(|(_, ext)| ext)
                .unwrap_or("");
            workbook.images.insert(
                image_id.clone(),
                ImageData {
                    bytes,
                    content_type: Some(content_type_for_extension(ext).to_string()),
                },
            );
        }

        images.push(CellImageEntry {
            embed_rel_id,
            target_path,
            image_id,
            pic_xml: slice_node_xml(&pic, xml),
        });
    }

    // Some Excel-generated `cellimages.xml` payloads use a lightweight schema where the
    // relationship ID is stored directly on a `<cellImage r:id="…">` element (rather than
    // embedding a full DrawingML `<pic>` subtree).
    for cell_image in doc
        .root_element()
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cellImage")
    {
        let Some(embed_rel_id) = get_cell_image_rel_id(&cell_image) else {
            continue;
        };

        // Avoid duplicating images already discovered via `<pic>`.
        if images.iter().any(|img| img.embed_rel_id == embed_rel_id) {
            continue;
        }

        let rel = rels_by_id.get(&embed_rel_id).ok_or_else(|| {
            XlsxError::Invalid(format!(
                "cellimages.xml references missing image relationship {embed_rel_id}"
            ))
        })?;

        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_uri != REL_TYPE_IMAGE {
            continue;
        }

        let target_path = resolve_target(path, &rel.target);
        let file_name = target_path
            .strip_prefix("xl/media/")
            .unwrap_or(&target_path)
            .to_string();
        let image_id = ImageId::new(file_name);

        if workbook.images.get(&image_id).is_none() {
            let bytes = parts
                .get(&target_path)
                .ok_or_else(|| XlsxError::MissingPart(target_path.clone()))?
                .clone();
            let ext = image_id
                .as_str()
                .rsplit_once('.')
                .map(|(_, ext)| ext)
                .unwrap_or("");
            workbook.images.insert(
                image_id.clone(),
                ImageData {
                    bytes,
                    content_type: Some(content_type_for_extension(ext).to_string()),
                },
            );
        }

        images.push(CellImageEntry {
            embed_rel_id,
            target_path,
            image_id,
            pic_xml: slice_node_xml(&cell_image, xml),
        });
    }

    Ok(CellImagesPart {
        path: path.to_string(),
        rels_path,
        images,
    })
}

fn resolve_target_best_effort(
    source_part: &str,
    rels_part: &str,
    target: &str,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<String> {
    // OPC relationship targets are typically resolved relative to the source part's directory.
    // However, some producers appear to emit paths relative to the `.rels` directory instead.
    // We'll resolve using the standard rule first and fall back to alternative interpretations
    // only when the referenced part cannot be found.
    let direct = resolve_target(source_part, target);
    if parts.contains_key(&direct) {
        return Ok(direct);
    }

    // Fallback: treat the target as relative to the relationships part location.
    let rels_relative = resolve_target(rels_part, target);
    if parts.contains_key(&rels_relative) {
        return Ok(rels_relative);
    }

    // Fallback: if the resolved path accidentally escaped the `xl/` prefix (e.g. `../media/...`
    // from a workbook-level part), try re-rooting it under `xl/`.
    if !direct.starts_with("xl/") {
        let xl_prefixed = format!("xl/{direct}");
        if parts.contains_key(&xl_prefixed) {
            return Ok(xl_prefixed);
        }
    }

    Err(XlsxError::MissingPart(direct))
}

fn get_blip_embed_rel_id(blip_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    // The canonical namespaces form.
    blip_node
        .attribute((REL_NS, "embed"))
        .or_else(|| blip_node.attribute("r:embed"))
        .or_else(|| {
            // Some XML libraries represent namespaced attributes using Clark notation:
            // `{namespace}localname`.
            let clark = format!("{{{REL_NS}}}embed");
            blip_node.attribute(clark.as_str())
        })
        .map(|s| s.to_string())
}

fn get_cell_image_rel_id(cell_image_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    cell_image_node
        .attribute((REL_NS, "id"))
        .or_else(|| cell_image_node.attribute("r:id"))
        .or_else(|| {
            let clark = format!("{{{REL_NS}}}id");
            cell_image_node.attribute(clark.as_str())
        })
        .map(|s| s.to_string())
}

fn slice_node_xml(node: &roxmltree::Node<'_, '_>, doc: &str) -> Option<String> {
    let range = node.range();
    doc.get(range).map(|s| s.to_string())
}
