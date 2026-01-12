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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellImagesPartInfo {
    /// ZIP entry name for the cell images part (typically `xl/cellimages.xml`).
    pub part_path: String,
    /// ZIP entry name for the `.rels` associated with [`Self::part_path`]
    /// (typically `xl/_rels/cellimages.xml.rels`).
    pub rels_path: String,
    /// All `r:embed` references discovered under `<a:blip>` elements in `cellimages.xml`,
    /// resolved to concrete `xl/media/*` parts via `cellimages.xml.rels`.
    pub embeds: Vec<CellImageEmbed>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellImageEmbed {
    /// Relationship ID from `<a:blip r:embed="...">`.
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
    /// `xl/_rels/workbook.xml.rels` for any relationship whose `Target` ends with `cellimages.xml`
    /// (case-insensitive).
    pub fn cell_images_part_info(&self) -> Result<Option<CellImagesPartInfo>> {
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

        let mut embeds = Vec::with_capacity(embed_rids.len());
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

            embeds.push(CellImageEmbed {
                embed_rid: rid,
                target_part,
                target_bytes,
            });
        }

        Ok(Some(CellImagesPartInfo {
            part_path,
            rels_path,
            embeds,
        }))
    }
}

fn discover_cell_images_part(pkg: &XlsxPackage) -> Result<Option<String>> {
    // Preferred heuristic: scan workbook.xml.rels for a relationship target ending with
    // cellimages.xml (case-insensitive). Do not assume a specific relationship type.
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

                if rel.target.to_ascii_lowercase().ends_with("cellimages.xml") {
                    let candidate = resolve_target(WORKBOOK_PART, &rel.target);
                    if pkg.part(&candidate).is_some() {
                        return Ok(Some(candidate));
                    }
                }
            }
        }
    }

    // Fallback: accept a well-known canonical part name if present.
    if pkg.part(FALLBACK_CELL_IMAGES_PART).is_some() {
        return Ok(Some(FALLBACK_CELL_IMAGES_PART.to_string()));
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
                if local_name(name.as_ref()).eq_ignore_ascii_case(b"blip") {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"embed") {
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

        if !rel.type_uri.eq_ignore_ascii_case(IMAGE_REL_TYPE) {
            continue;
        }

        let target_part = resolve_target(cell_images_part, &rel.target);
        out.insert(rel.id, target_part);
    }

    Ok(out)
}

