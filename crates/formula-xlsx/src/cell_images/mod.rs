//! Workbook-level "images in cells" (`xl/cellimages*.xml`) parsing (when present).
//!
//! Some workbooks include a workbook-level `xl/cellimages*.xml` part containing DrawingML image
//! payloads (`<pic>`, `<a:blip r:embed="…">`, etc.) that reference `xl/media/*` via relationships.
//! This module provides best-effort parsing for that optional schema.
//!
//! Modern Excel "Place in Cell" embedded images commonly do **not** use `xl/cellimages.xml`.
//! Instead Excel typically stores the image association via `xl/metadata.xml` plus the rich value
//! store: `xl/richData/{rdrichvalue*.xml,rdrichvaluestructure.xml,rdRichValueTypes.xml,richValueRel.xml}`
//! (and relationships) together with `xl/media/*`.
//!
//! For a concrete walkthrough of the current Excel schema(s), see
//! `docs/xlsx-embedded-images-in-cells.md`.
mod part_info;

pub use part_info::{CellImageEmbed, CellImagesPartInfo};

use std::collections::{BTreeMap, HashMap, HashSet};

use formula_model::drawings::{ImageData, ImageId};
use roxmltree::Document;

use crate::drawings::{content_type_for_extension, REL_TYPE_IMAGE};
use crate::path::resolve_target;
use crate::XlsxError;

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
        if let Ok(part) = parse_cell_images_part(path, parts) {
            load_cell_images_part_media(&part, parts, workbook);
        }
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
    /// Resolved target part path (e.g. `xl/media/image1.png`), if the embed relationship can be
    /// resolved to an image target.
    pub target: Option<String>,
    /// Best-effort raw XML for the `<cellImage>` subtree (or `<xdr:pic>` when present).
    pub raw_xml: String,
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

            let part = parse_cell_images_part(path, parts)?;
            load_cell_images_part_media(&part, parts, workbook);
            cell_images_parts.push(part);
        }

        Ok(Self {
            parts: cell_images_parts,
        })
    }
}
impl CellImagesPart {
    /// Parse the canonical workbook-level `xl/cellImages.xml` part (if present).
    ///
    /// This mirrors the low-level "parts map" parsing style used by other parsers in this crate
    /// (e.g. drawings). It intentionally does *not* attempt to map images to cells; it only parses
    /// the `cellImages.xml` schema and resolves image relationship IDs to media targets.
    pub fn from_parts(parts: &BTreeMap<String, Vec<u8>>) -> Result<Option<Self>> {
        // Excel uses `xl/cellImages.xml` (capital I) today, but be permissive and accept any
        // casing.
        let part_path = parts
            .keys()
            .find(|p| crate::zip_util::zip_part_names_equivalent(p.as_str(), "xl/cellimages.xml"))
            .cloned();
        let Some(part_path) = part_path else {
            return Ok(None);
        };
        parse_cell_images_part(&part_path, parts).map(Some)
    }
}
fn is_cell_images_part(path: &str) -> bool {
    let Some(rest) = path.strip_prefix("xl/") else {
        return false;
    };
    let Some(file_name) = rest.rsplit('/').next() else {
        return false;
    };
    let Some(stem) = crate::ascii::strip_suffix_ignore_case(file_name, ".xml") else {
        return false;
    };
    let Some(suffix) = crate::ascii::strip_prefix_ignore_case(stem, "cellimages") else {
        return false;
    };
    suffix.is_empty() || suffix.as_bytes().iter().all(u8::is_ascii_digit)
}

fn parse_cell_images_part(path: &str, parts: &BTreeMap<String, Vec<u8>>) -> Result<CellImagesPart> {
    let rels_path = crate::openxml::rels_part_name(path);

    fn strip_fragment(target: &str) -> &str {
        target
            .split_once('#')
            .map(|(base, _)| base)
            .unwrap_or(target)
    }

    let rid_to_target: HashMap<String, String> = parts
        .get(&rels_path)
        .and_then(|rels_bytes| crate::openxml::parse_relationships(rels_bytes).ok())
        .map(|relationships| {
            relationships
                .into_iter()
                .filter(|rel| {
                    !rel.target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                        && rel.type_uri == REL_TYPE_IMAGE
                })
                .filter_map(|rel| {
                    let target = strip_fragment(&rel.target);
                    if target.is_empty() {
                        return None;
                    }

                    let target_path = resolve_target_best_effort(path, &rels_path, target, parts)
                        .unwrap_or_else(|_| {
                            // Fall back to strict URI resolution if the target cannot be found in
                            // `parts`. This preserves a deterministic output even for incomplete
                            // workbooks.
                            let mut strict = resolve_target(path, target);
                            if strict.starts_with("media/") {
                                strict = format!("xl/{strict}");
                            }
                            strict
                        });
                    Some((rel.id, target_path))
                })
                .collect()
        })
        .unwrap_or_default();

    let xml_bytes = parts
        .get(path)
        .ok_or_else(|| XlsxError::MissingPart(path.to_string()))?;
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("cell images xml not utf-8: {e}")))?;

    let doc = Document::parse(xml)?;

    // Basic validation: ensure this is actually a `cellImages` document. We only match on local
    // name to ignore prefixes like `cx:cellImages` / `etc:cellImages`.
    let root = doc.root_element();
    if !root.tag_name().name().eq_ignore_ascii_case("cellImages") {
        return Err(XlsxError::Invalid(format!(
            "unexpected root element in {path}: {}",
            root.tag_name().name()
        )));
    }

    let mut images = Vec::new();
    // Track subtrees we've already materialized to avoid double-counting when scanning for `<blip>`
    // nodes. Use the byte range in the source XML as a cheap stable identity.
    let mut seen_ranges: HashSet<(usize, usize)> = HashSet::new();
    // Collect images in document order. Excel can represent the image relationship ID either:
    // - as DrawingML `<xdr:pic>` payloads referencing `<a:blip r:embed="…">`, or
    // - via a lightweight `<cellImage r:id="…">` schema.
    for node in root.descendants().filter(|n| n.is_element()) {
        let name = node.tag_name().name();
        if name.eq_ignore_ascii_case("pic") {
            // Excel emits DrawingML `<xdr:pic>` payloads for in-cell images, which reference the
            // media via `<a:blip r:embed="…">`.
            let Some(blip) = node
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("blip"))
            else {
                continue;
            };
            let Some(embed_rel_id) = get_blip_embed_rel_id(&blip) else {
                continue;
            };
            let target = rid_to_target.get(&embed_rel_id).cloned();
            images.push(CellImageEntry {
                embed_rel_id,
                target,
                raw_xml: slice_node_xml(&node, xml).unwrap_or_default(),
            });
            let range = node.range();
            seen_ranges.insert((range.start, range.end));
        }
        if name.eq_ignore_ascii_case("cellImage") {
            let range = node.range();
            // If this `<cellImage>` includes a full `<pic>` payload, we'll pick it up when we hit
            // the `<pic>` node itself. Avoid double-counting in that case.
            if node
                .descendants()
                .any(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("pic"))
            {
                // Still mark the `<cellImage>` subtree as seen so that any nested `<blip>` elements
                // that aren't within `<pic>` aren't treated as separate images (this matches our
                // historical "only parse `<pic>` when present" behavior).
                seen_ranges.insert((range.start, range.end));
                continue;
            }

            let embed_rel_id = get_cell_image_rel_id(&node).or_else(|| {
                // Some minimal schemas store the relationship on a nested `<a:blip r:embed="...">`.
                node.descendants()
                    .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("blip"))
                    .and_then(|blip| get_blip_embed_rel_id(&blip))
            });
            let Some(embed_rel_id) = embed_rel_id else {
                continue;
            };
            let target = rid_to_target.get(&embed_rel_id).cloned();
            images.push(CellImageEntry {
                embed_rel_id,
                target,
                raw_xml: slice_node_xml(&node, xml).unwrap_or_default(),
            });
            seen_ranges.insert((range.start, range.end));
        }

        // Some schema variants store `<a:blip r:embed="...">` nodes without a `<pic>` wrapper and
        // even without a `<cellImage>` wrapper. Scan all `blip` elements as a fallback.
        if name.eq_ignore_ascii_case("blip") {
            let Some(embed_rel_id) = get_blip_embed_rel_id(&node) else {
                continue;
            };

            let container = node
                .ancestors()
                .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("pic"))
                .or_else(|| {
                    node.ancestors().find(|n| {
                        n.is_element() && n.tag_name().name().eq_ignore_ascii_case("cellImage")
                    })
                })
                .unwrap_or(node);

            let range = container.range();
            if !seen_ranges.insert((range.start, range.end)) {
                continue;
            }

            let target = rid_to_target.get(&embed_rel_id).cloned();
            images.push(CellImageEntry {
                embed_rel_id,
                target,
                raw_xml: slice_node_xml(&container, xml).unwrap_or_default(),
            });
        }
    }

    Ok(CellImagesPart {
        path: path.to_string(),
        rels_path,
        images,
    })
}

fn load_cell_images_part_media(
    part: &CellImagesPart,
    parts: &BTreeMap<String, Vec<u8>>,
    workbook: &mut formula_model::Workbook,
) {
    for image in &part.images {
        let Some(target_path) = image.target.as_deref() else {
            continue;
        };
        let Some(bytes) = parts.get(target_path) else {
            continue;
        };
        let image_id = image_id_from_target_path(target_path);
        if workbook.images.get(&image_id).is_some() {
            continue;
        }
        let content_type = {
            let ext = image_id
                .as_str()
                .rsplit_once('.')
                .map(|(_, ext)| ext)
                .unwrap_or("");
            content_type_for_extension(ext).to_string()
        };
        workbook.images.insert(
            image_id,
            ImageData {
                bytes: bytes.clone(),
                content_type: Some(content_type),
            },
        );
    }
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

fn get_relationship_attr(blip_node: &roxmltree::Node<'_, '_>, local: &str) -> Option<String> {
    // The canonical namespaces form.
    blip_node
        .attribute((REL_NS, local))
        .or_else(|| blip_node.attribute(format!("r:{local}").as_str()))
        .or_else(|| {
            // Some XML libraries represent namespaced attributes using Clark notation:
            // `{namespace}localname`.
            let clark = format!("{{{REL_NS}}}{local}");
            blip_node.attribute(clark.as_str())
        })
        .map(|s| s.to_string())
}

fn get_blip_embed_rel_id(blip_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    get_relationship_attr(blip_node, "embed")
}

fn get_cell_image_rel_id(cell_image_node: &roxmltree::Node<'_, '_>) -> Option<String> {
    // Most recent Excel builds seem to use `r:id` on `<cellImage>`, but older variants may use
    // `r:embed` (mirroring `<a:blip r:embed>`). Treat either as a relationship ID.
    get_relationship_attr(cell_image_node, "id")
        .or_else(|| get_relationship_attr(cell_image_node, "embed"))
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
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_images_part_from_parts_extracts_embeds_and_resolves_targets() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:cellImage>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
    </xdr:pic>
  </cx:cellImage>
  <cx:cellImage>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill>
    </xdr:pic>
  </cx:cellImage>
</cx:cellImages>"#;

        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#frag"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png#frag"/>
</Relationships>"#;

        let parts: BTreeMap<String, Vec<u8>> = [
            ("xl/cellImages.xml".to_string(), xml.to_vec()),
            ("xl/_rels/cellImages.xml.rels".to_string(), rels.to_vec()),
        ]
        .into_iter()
        .collect();

        let part = CellImagesPart::from_parts(&parts)
            .expect("parse")
            .expect("expected part");
        assert_eq!(part.path, "xl/cellImages.xml");
        assert_eq!(part.rels_path, "xl/_rels/cellImages.xml.rels");
        assert_eq!(part.images.len(), 2);
        assert_eq!(part.images[0].embed_rel_id, "rId1");
        assert_eq!(
            part.images[0].target.as_deref(),
            Some("xl/media/image1.png")
        );
        assert!(!part.images[0].raw_xml.is_empty());

        assert_eq!(part.images[1].embed_rel_id, "rId2");
        assert_eq!(
            part.images[1].target.as_deref(),
            Some("xl/media/image2.png")
        );
    }

    #[test]
    fn cell_images_part_preserves_document_order_when_mixing_schemas() {
        // Mix a lightweight `<cellImage r:id="...">` entry with a `<cellImage><xdr:pic>...</xdr:pic>`
        // entry and ensure we preserve document order (lightweight first).
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:cellImage r:id="rId1"/>
  <cx:cellImage>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill>
    </xdr:pic>
  </cx:cellImage>
</cx:cellImages>"#;

        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>"#;

        let parts: BTreeMap<String, Vec<u8>> = [
            ("xl/cellImages.xml".to_string(), xml.to_vec()),
            ("xl/_rels/cellImages.xml.rels".to_string(), rels.to_vec()),
        ]
        .into_iter()
        .collect();

        let part = CellImagesPart::from_parts(&parts)
            .expect("parse")
            .expect("expected part");
        assert_eq!(part.images.len(), 2);
        assert_eq!(part.images[0].embed_rel_id, "rId1");
        assert_eq!(part.images[1].embed_rel_id, "rId2");
    }

    #[test]
    fn cell_images_part_supports_nested_blip_embeds_without_pic_payload() {
        // Some producers store the relationship ID on a nested `<a:blip r:embed="...">` directly
        // under `<cellImage>`, without embedding a full DrawingML `<xdr:pic>` subtree.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:cellImages xmlns:cx="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:cellImage>
    <a:blip r:embed="rId1"/>
  </cx:cellImage>
</cx:cellImages>"#;

        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

        let parts: BTreeMap<String, Vec<u8>> = [
            ("xl/cellImages.xml".to_string(), xml.to_vec()),
            ("xl/_rels/cellImages.xml.rels".to_string(), rels.to_vec()),
        ]
        .into_iter()
        .collect();

        let part = CellImagesPart::from_parts(&parts)
            .expect("parse")
            .expect("expected part");
        assert_eq!(part.images.len(), 1);
        assert_eq!(part.images[0].embed_rel_id, "rId1");
        assert_eq!(
            part.images[0].target.as_deref(),
            Some("xl/media/image1.png")
        );
        assert!(!part.images[0].raw_xml.is_empty());
    }
}
