//! Workbook-level "images in cells" (`xl/cellimages*.xml`) parsing.
//!
//! Modern Excel features like "Place in Cell" images and the `IMAGE()` function
//! appear to rely on a workbook-level `xl/cellimages*.xml` part containing
//! DrawingML `<pic>` payloads that reference media via relationships.
mod part_info;
pub use part_info::{CellImageEmbed, CellImagesPartInfo};

use std::collections::{BTreeMap, HashMap};

use formula_model::drawings::{ImageData, ImageId};
use roxmltree::Document;

use crate::drawings::{content_type_for_extension, REL_TYPE_IMAGE};
use crate::path::resolve_target;
use crate::XlsxError;

mod part_info;

pub use part_info::{CellImageEmbed, CellImagesPartInfo};

type Result<T> = std::result::Result<T, XlsxError>;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Best-effort loader for workbook-level "in-cell" images.
///
/// This is intentionally best-effort/tolerant of incomplete or malformed workbooks:
/// - Missing `cellimages*.xml` parts → no-op
/// - Missing `.rels` → skip that part
/// - Missing referenced media part → skip that image/part
/// - Parse errors → skip that part
///
/// This is used by the workbook reader to opportunistically populate
/// `workbook.images` during import.
pub fn load_cell_images_from_parts(
    parts: &BTreeMap<String, Vec<u8>>,
    workbook: &mut formula_model::Workbook,
) {
    for path in parts.keys() {
        if !is_cell_images_part(path) {
            continue;
        }
        // Best-effort: ignore parse errors for cell image parts.
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
    let Some(file_name) = rest.rsplit('/').next() else {
        return false;
    };
    let lower = file_name.to_ascii_lowercase();
    let Some(stem) = lower.strip_suffix(".xml") else {
        return false;
    };
    let Some(suffix) = stem.strip_prefix("cellimages") else {
        return false;
    };
    suffix.is_empty() || suffix.chars().all(|c| c.is_ascii_digit())
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
    // Excel emits DrawingML `<xdr:pic>` payloads for in-cell images, which reference the media via
    // `<a:blip r:embed="…">`.
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

        let Some(rel) = rels_by_id.get(&embed_rel_id) else {
            // Best-effort: skip broken references instead of failing the whole workbook load.
            continue;
        };

        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            // External relationships are not backed by OPC parts.
            continue;
        }
        if rel.type_uri != REL_TYPE_IMAGE {
            // Be conservative: `<a:blip r:embed>` should refer to an image relationship.
            continue;
        }

        let target_path = match resolve_target_best_effort(path, &rels_path, &rel.target, parts) {
            Ok(path) => path,
            Err(_) => continue,
        };
        let image_id = image_id_from_target_path(&target_path);

        if workbook.images.get(&image_id).is_none() {
            let Some(bytes) = parts.get(&target_path) else {
                // Best-effort: missing media should not prevent the workbook from loading.
                continue;
            };
            let bytes = bytes.clone();
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

        let Some(rel) = rels_by_id.get(&embed_rel_id) else {
            // Best-effort: skip broken references instead of failing the whole workbook load.
            continue;
        };
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

        let target_path = match resolve_target_best_effort(path, &rels_path, &rel.target, parts) {
            Ok(path) => path,
            Err(_) => continue,
        };
        let image_id = image_id_from_target_path(&target_path);

        if workbook.images.get(&image_id).is_none() {
            let Some(bytes) = parts.get(&target_path) else {
                // Best-effort: missing media should not prevent the workbook from loading.
                continue;
            };
            let bytes = bytes.clone();
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

    // Some producers emit targets rooted at `media/` (or similar) even for workbook-level parts
    // under `xl/`. Try toggling the `xl/` prefix before other fallbacks.
    if let Some(stripped) = direct.strip_prefix("xl/") {
        if parts.contains_key(stripped) {
            return Ok(stripped.to_string());
        }
    } else {
        let xl_prefixed = format!("xl/{direct}");
        if parts.contains_key(&xl_prefixed) {
            return Ok(xl_prefixed);
        }
    }

    // Fallback: treat the target as relative to the relationships part location.
    let rels_relative = resolve_target(rels_part, target);
    if parts.contains_key(&rels_relative) {
        return Ok(rels_relative);
    }

    // Apply the same `xl/` prefix toggling for the rels-relative candidate.
    if let Some(stripped) = rels_relative.strip_prefix("xl/") {
        if parts.contains_key(stripped) {
            return Ok(stripped.to_string());
        }
    } else {
        let xl_prefixed = format!("xl/{rels_relative}");
        if parts.contains_key(&xl_prefixed) {
            return Ok(xl_prefixed);
        }
    }

    Err(XlsxError::MissingPart(direct))
}

fn get_blip_embed_rel_id(blip_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    get_relationship_id_attr(blip_node, "embed")
}

fn get_relationship_id_attr(node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    let prefixed = format!("r:{local}");
    let clark = format!("{{{REL_NS}}}{local}");
    // Prefer the canonical namespace-aware lookup.
    node.attribute((REL_NS, local))
        .or_else(|| node.attribute(prefixed.as_str()))
        // Some XML libraries represent namespaced attributes using Clark notation:
        // `{namespace}localname`.
        .or_else(|| node.attribute(clark.as_str()))
        .map(|s| s.to_string())
}

fn get_cell_image_rel_id(cell_image_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    // Most recent Excel builds seem to use `r:id` on `<cellImage>`, but older variants may use
    // `r:embed` (mirroring `<a:blip r:embed>`). Treat either as a relationship ID.
    cell_image_node
        .attribute((REL_NS, "id"))
        .or_else(|| cell_image_node.attribute("r:id"))
        .or_else(|| {
            // Clark notation: `{namespace}localname`.
            let clark = format!("{{{REL_NS}}}id");
            cell_image_node.attribute(clark.as_str())
        })
        .or_else(|| cell_image_node.attribute((REL_NS, "embed")))
        .or_else(|| cell_image_node.attribute("r:embed"))
        .or_else(|| {
            let clark = format!("{{{REL_NS}}}embed");
            cell_image_node.attribute(clark.as_str())
        })
        .map(|s| s.to_string())
}

fn image_id_from_target_path(target_path: &str) -> ImageId {
    let file_name = target_path
        .strip_prefix("xl/media/")
        .or_else(|| target_path.strip_prefix("media/"))
        .unwrap_or(target_path)
        .to_string();
    ImageId::new(file_name)
}
fn slice_node_xml(node: &roxmltree::Node<'_, '_>, doc: &str) -> Option<String> {
    let range = node.range();
    doc.get(range).map(|s| s.to_string())
}
