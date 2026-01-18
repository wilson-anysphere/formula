//! Package-level helpers for workbook `xl/cellimages*.xml` parts.

use std::collections::HashMap;
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{local_name, parse_relationships};
use crate::path::{rels_for_part, resolve_target};
use crate::{XlsxError, XlsxPackage};

type Result<T> = std::result::Result<T, XlsxError>;

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const FALLBACK_CELL_IMAGES_PART: &str = "xl/cellimages.xml";

/// Standard OpenXML relationship type for embedded images.
const IMAGE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

fn strip_fragment(target: &str) -> &str {
    // OPC relationship targets are URIs; some producers include a fragment (e.g. `foo.xml#bar`).
    // OPC part names do not include fragments, so they must be ignored when resolving ZIP entries.
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellImagesPartInfo {
    /// ZIP entry name for the cell images part (typically `xl/cellimages.xml`).
    pub part_path: String,
    /// ZIP entry name for the `.rels` associated with [`Self::part_path`]
    /// (typically `xl/_rels/cellimages.xml.rels`).
    pub rels_path: String,
    /// All image relationship IDs discovered in `cellimages.xml`, resolved to concrete `xl/media/*`
    /// parts via `cellimages.xml.rels`.
    pub embeds: Vec<CellImageEmbed>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellImageEmbed {
    /// Relationship ID referenced by `cellimages.xml` (e.g. `rId1`).
    pub embed_rid: String,
    /// Resolved relationship target part name (e.g. `xl/media/image1.png`).
    pub target_part: String,
    /// Bytes for the resolved target part (if present in the package).
    pub target_bytes: Vec<u8>,
}

impl XlsxPackage {
    /// Discover and parse the optional `cellimages.xml` part used by Excel's "image in cell"
    /// feature.
    ///
    /// This implementation intentionally does *not* rely on a fixed relationship type URI for the
    /// workbook -> cellimages relationship. Instead it prefers scanning
    /// `xl/_rels/workbook.xml.rels` for any relationship whose `Target` resolves to a
    /// `xl/cellimages*.xml` part (case-insensitive).
    pub fn cell_images_part_info(&self) -> Result<Option<super::CellImagesPartInfo>> {
        let Some(part_path) = discover_cell_images_part(self)? else {
            return Ok(None);
        };

        let part_bytes = self
            .part(&part_path)
            .ok_or_else(|| XlsxError::MissingPart(part_path.clone()))?;

        // Optional validation that the discovered part is actually a cellImages document.
        validate_cell_images_root(part_bytes, &part_path)?;

        let rels_path = rels_for_part(&part_path);
        let rels_bytes = self
            .part(&rels_path)
            .ok_or_else(|| XlsxError::MissingPart(rels_path.clone()))?;

        let embed_rids = parse_cell_images_embeds(part_bytes)?;
        let rid_to_target = parse_cell_images_image_relationships(&part_path, rels_bytes)?;

        let mut embeds = Vec::new();
        if embeds.try_reserve_exact(embed_rids.len()).is_err() {
            return Err(XlsxError::AllocationFailure("cell_images_part_info embeds"));
        }
        for rid in embed_rids {
            let target_part = rid_to_target.get(&rid).cloned().ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "cellimages.xml references missing image relationship {rid}"
                ))
            })?;
            let target_bytes = self
                .part(&target_part)
                .ok_or_else(|| XlsxError::MissingPart(target_part.clone()))?
                .to_vec();

            embeds.push(super::CellImageEmbed {
                embed_rid: rid,
                target_part,
                target_bytes,
            });
        }

        Ok(Some(super::CellImagesPartInfo {
            part_path,
            rels_path,
            embeds,
        }))
    }
}

fn discover_cell_images_part(pkg: &XlsxPackage) -> Result<Option<String>> {
    // Preferred heuristic: scan workbook.xml.rels for a relationship target that resolves to a
    // `xl/cellimages*.xml` part (case-insensitive). Do not assume a specific relationship type.
    if let Some(rels_bytes) = pkg.part(WORKBOOK_RELS_PART) {
        // Best-effort: if workbook rels are malformed, fall back to direct part existence checks.
        if let Ok(relationships) = parse_relationships(rels_bytes) {
            for rel in relationships {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                let target = strip_fragment(&rel.target);
                if target.is_empty() {
                    continue;
                }
                let candidate = resolve_target(WORKBOOK_PART, target);
                if super::is_cell_images_part(&candidate) && pkg.part(&candidate).is_some() {
                    return Ok(Some(candidate));
                }
            }
        }
    }

    // Fallback: accept a well-known canonical part name if present.
    if pkg.part(FALLBACK_CELL_IMAGES_PART).is_some() {
        return Ok(Some(FALLBACK_CELL_IMAGES_PART.to_string()));
    }

    // Fallback: scan for any `xl/cellimages*.xml` part when workbook rels are missing or do not
    // reference the part.
    for name in pkg.part_names() {
        if super::is_cell_images_part(name) {
            return Ok(Some(name.to_string()));
        }
    }

    Ok(None)
}

fn validate_cell_images_root(xml: &[u8], part_name: &str) -> Result<()> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => {
                let name = e.name();
                let root = local_name(name.as_ref());
                if root.eq_ignore_ascii_case(b"cellImages") {
                    return Ok(());
                }
                let root_str = std::str::from_utf8(root).unwrap_or("<non-utf8>");
                return Err(XlsxError::Invalid(format!(
                    "unexpected root element in {part_name}: {root_str}"
                )));
            }
            Event::Eof => {
                return Err(XlsxError::Invalid(format!(
                    "{part_name} is empty (expected cellImages document)"
                )))
            }
            _ => {}
        }
        buf.clear();
    }
}

fn parse_cell_images_embeds(xml: &[u8]) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut embeds: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) => {
                let name = e.name();
                let elem = local_name(name.as_ref());

                // DrawingML `a:blip r:embed="…"` (commonly under `<pic>`).
                if elem.eq_ignore_ascii_case(b"blip") {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"embed") {
                            embeds.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }

                // Lightweight schema: `<cellImage r:id="…">` or `<cellImage r:embed="…">`.
                if elem.eq_ignore_ascii_case(b"cellImage") {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        if key.eq_ignore_ascii_case(b"id") || key.eq_ignore_ascii_case(b"embed") {
                            embeds.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(embeds)
}

fn parse_cell_images_image_relationships(
    cell_images_part: &str,
    rels_xml: &[u8],
) -> Result<HashMap<String, String>> {
    let relationships = parse_relationships(rels_xml)?;
    let mut out: HashMap<String, String> = HashMap::new();

    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        // Be conservative: only resolve image relationships.
        if !rel.type_uri.eq_ignore_ascii_case(IMAGE_REL_TYPE) {
            continue;
        }

        let target = strip_fragment(&rel.target);
        if target.is_empty() {
            continue;
        }

        let mut target_part = resolve_target(cell_images_part, target);
        // Some producers emit targets like `../media/image1.png` for workbook-level parts such as
        // `xl/cellimages.xml`, which resolves to `media/image1.png` with strict URI resolution.
        // In XLSX packages, worksheet media lives under `xl/media/*`, so normalize these to an
        // `xl/`-prefixed path.
        if target_part.starts_with("media/") {
            target_part = format!("xl/{target_part}");
        }

        out.insert(rel.id, target_part);
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_cell_images_embeds_extracts_blip_embeds() {
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

        let embeds = parse_cell_images_embeds(xml).expect("parse embeds");
        assert_eq!(embeds, vec!["rId1".to_string(), "rId2".to_string()]);
    }

    #[test]
    fn parse_cell_images_image_relationships_normalizes_parent_media_targets() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#frag"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png#frag"/>
</Relationships>"#;

        let map =
            parse_cell_images_image_relationships("xl/cellImages.xml", rels).expect("rels parse");
        assert_eq!(map.get("rId1").map(String::as_str), Some("xl/media/image1.png"));
        assert_eq!(map.get("rId2").map(String::as_str), Some("xl/media/image2.png"));
    }
}
