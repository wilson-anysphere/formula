use std::collections::BTreeMap;
use std::io::{Read, Seek};

use formula_model::drawings::{
    Anchor, AnchorPoint, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize, ImageData,
    ImageId,
};
use roxmltree::{Document, Node};

use crate::path::resolve_target;
use crate::relationships::{Relationship, Relationships};
use crate::zip_util::open_zip_part;
use crate::XlsxError;
use zip::ZipArchive;

type Result<T> = std::result::Result<T, XlsxError>;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const GRAPHIC_DATA_CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DrawingRef(pub usize);

#[derive(Debug, Clone)]
pub struct DrawingPart {
    pub sheet_index: usize,
    pub path: String,
    pub rels_path: String,
    pub objects: Vec<DrawingObject>,
    relationships: Relationships,
    /// Namespace declarations (`xmlns` / `xmlns:*`) found on the root `<xdr:wsDr>` element.
    ///
    /// Some DrawingML producers (e.g. SmartArt) declare additional namespaces on `<xdr:wsDr>`.
    /// When we preserve raw XML snippets (e.g. `<dgm:relIds/>` inside a SmartArt graphic frame),
    /// these namespace declarations must be preserved to avoid emitting invalid XML with
    /// undeclared prefixes.
    root_xmlns: BTreeMap<String, String>,
}

impl DrawingPart {
    pub fn new_empty(sheet_index: usize, path: String, rels_path: String) -> Self {
        Self {
            sheet_index,
            path,
            rels_path,
            objects: Vec::new(),
            relationships: Relationships::default(),
            root_xmlns: BTreeMap::new(),
        }
    }

    pub fn parse_from_parts(
        sheet_index: usize,
        path: &str,
        parts: &BTreeMap<String, Vec<u8>>,
        workbook: &mut formula_model::Workbook,
    ) -> Result<Self> {
        let rels_path = drawing_rels_path(path);
        let rels_bytes = parts
            .get(&rels_path)
            .ok_or_else(|| XlsxError::MissingPart(format!("missing drawing rels: {rels_path}")))?;
        let rels_xml = std::str::from_utf8(rels_bytes)
            .map_err(|e| XlsxError::Invalid(format!("drawing rels not utf-8: {e}")))?;
        let relationships = Relationships::from_xml(rels_xml)?;

        let drawing_bytes = parts
            .get(path)
            .ok_or_else(|| XlsxError::MissingPart(path.to_string()))?;
        let drawing_xml = std::str::from_utf8(drawing_bytes)
            .map_err(|e| XlsxError::Invalid(format!("drawing xml not utf-8: {e}")))?;

        let doc = Document::parse(drawing_xml)?;
        let root_xmlns = extract_root_xmlns(doc.root_element());
        let mut objects = Vec::new();

        for (z, anchor_node) in doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
            .enumerate()
        {
            let anchor_tag = anchor_node.tag_name().name();
            if anchor_tag != "oneCellAnchor"
                && anchor_tag != "twoCellAnchor"
                && anchor_tag != "absoluteAnchor"
            {
                continue;
            }

            let anchor = parse_anchor(&anchor_node)?;
            let anchor_preserved = parse_anchor_preserved(&anchor_node, drawing_xml);

            if let Some(pic) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "pic")
            {
                let (id, pic_xml, embed) = parse_pic(&pic, drawing_xml)?;
                let image_id = resolve_image_id(&relationships, &embed, path, parts, workbook)?;
                let mut preserved = anchor_preserved.clone();
                preserved.insert("xlsx.embed_rel_id".to_string(), embed.clone());
                preserved.insert("xlsx.pic_xml".to_string(), pic_xml);

                let size = extract_size_from_transform(&pic).or_else(|| size_from_anchor(anchor));

                objects.push(DrawingObject {
                    id,
                    kind: DrawingObjectKind::Image { image_id },
                    anchor,
                    z_order: z as i32,
                    size,
                    preserved,
                });
                continue;
            }

            if let Some(sp) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "sp")
            {
                let (id, sp_xml) = parse_named_node(&sp, drawing_xml, "cNvPr")?;
                let size = extract_size_from_transform(&sp).or_else(|| size_from_anchor(anchor));
                objects.push(DrawingObject {
                    id,
                    kind: DrawingObjectKind::Shape { raw_xml: sp_xml },
                    anchor,
                    z_order: z as i32,
                    size,
                    preserved: anchor_preserved.clone(),
                });
                continue;
            }

            if let Some(frame) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicFrame")
            {
                let (id, frame_xml) = parse_named_node(&frame, drawing_xml, "cNvPr")?;

                // `xdr:graphicFrame` is used for multiple object types (charts, SmartArt diagrams,
                // etc). Treat it as a chart placeholder only when it actually references a chart.
                //
                // Excel chart frames contain either:
                // - A `<c:chart r:id="...">` (or `cx:chart`) element, or
                // - An `a:graphicData` node with `uri=".../chart"`.
                let chart_node = frame
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "chart");

                let graphic_data_is_chart = frame
                    .descendants()
                    .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
                    .and_then(|n| n.attribute("uri"))
                    .is_some_and(|uri| uri == GRAPHIC_DATA_CHART_URI);

                let size = extract_size_from_transform(&frame).or_else(|| size_from_anchor(anchor));

                if chart_node.is_some() || graphic_data_is_chart {
                    if let Some(chart_rel_id) = chart_node
                        .and_then(|n| {
                            n.attribute((REL_NS, "id"))
                                .or_else(|| n.attribute("r:id"))
                                .or_else(|| n.attribute("id"))
                        })
                        .map(|s| s.to_string())
                    {
                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::ChartPlaceholder {
                                rel_id: chart_rel_id,
                                raw_xml: frame_xml,
                            },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    } else {
                        // Chart frame without a relationship id: preserve as an unknown anchor
                        // rather than inventing a placeholder rel id.
                        let raw_anchor =
                            slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    }
                } else {
                    // Non-chart `graphicFrame` (e.g. SmartArt): preserve the entire anchor subtree.
                    let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                    objects.push(DrawingObject {
                        id,
                        kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    });
                }
                continue;
            }

            // Unknown anchor type: preserve the entire anchor subtree.
            let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
            objects.push(DrawingObject {
                id: DrawingObjectId((z + 1) as u32),
                kind: DrawingObjectKind::Unknown {
                    raw_xml: raw_anchor,
                },
                anchor,
                z_order: z as i32,
                size: None,
                preserved: anchor_preserved,
            });
        }

        Ok(Self {
            sheet_index,
            path: path.to_string(),
            rels_path,
            objects,
            relationships,
            root_xmlns,
        })
    }

    /// Streaming-friendly parser that reads only the required drawing parts from a ZIP archive.
    ///
    /// This is best-effort: missing relationship parts or media payloads will result in fewer
    /// objects/images being materialized, but should not prevent parsing the rest of the drawing.
    pub fn parse_from_archive<R: Read + Seek>(
        sheet_index: usize,
        path: &str,
        archive: &mut ZipArchive<R>,
        workbook: &mut formula_model::Workbook,
    ) -> Result<Self> {
        let rels_path = drawing_rels_path(path);
        let relationships = match read_zip_part_optional(archive, &rels_path)? {
            Some(bytes) => std::str::from_utf8(&bytes)
                .ok()
                .and_then(|xml| Relationships::from_xml(xml).ok())
                .unwrap_or_default(),
            None => Relationships::default(),
        };

        let drawing_bytes = read_zip_part_optional(archive, path)?
            .ok_or_else(|| XlsxError::MissingPart(path.to_string()))?;
        let drawing_xml = std::str::from_utf8(&drawing_bytes)
            .map_err(|e| XlsxError::Invalid(format!("drawing xml not utf-8: {e}")))?;

        let doc = Document::parse(drawing_xml)?;
        let root_xmlns = extract_root_xmlns(doc.root_element());
        let mut objects = Vec::new();

        for (z, anchor_node) in doc
            .root_element()
            .children()
            .filter(|n| n.is_element())
            .enumerate()
        {
            let anchor_tag = anchor_node.tag_name().name();
            if anchor_tag != "oneCellAnchor"
                && anchor_tag != "twoCellAnchor"
                && anchor_tag != "absoluteAnchor"
            {
                continue;
            }

            let anchor = match parse_anchor(&anchor_node) {
                Ok(a) => a,
                Err(_) => continue,
            };
            let anchor_preserved = parse_anchor_preserved(&anchor_node, drawing_xml);

            if let Some(pic) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "pic")
            {
                if let Ok((id, pic_xml, embed)) = parse_pic(&pic, drawing_xml) {
                    if let Ok(image_id) = resolve_image_id_from_archive(
                        &relationships,
                        &embed,
                        path,
                        archive,
                        workbook,
                    ) {
                        let mut preserved = anchor_preserved.clone();
                        preserved.insert("xlsx.embed_rel_id".to_string(), embed.clone());
                        preserved.insert("xlsx.pic_xml".to_string(), pic_xml);

                        let size =
                            extract_size_from_transform(&pic).or_else(|| size_from_anchor(anchor));

                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::Image { image_id },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved,
                        });
                        continue;
                    }
                }
                // Fall back to preserving the full anchor subtree when we can't resolve the image.
            }

            if let Some(sp) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "sp")
            {
                if let Ok((id, sp_xml)) = parse_named_node(&sp, drawing_xml, "cNvPr") {
                    let size =
                        extract_size_from_transform(&sp).or_else(|| size_from_anchor(anchor));
                    objects.push(DrawingObject {
                        id,
                        kind: DrawingObjectKind::Shape { raw_xml: sp_xml },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    });
                    continue;
                }
                // Fall back to preserving the full anchor subtree when we can't parse the shape.
            }

            if let Some(frame) = anchor_node
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "graphicFrame")
            {
                if let Ok((id, frame_xml)) = parse_named_node(&frame, drawing_xml, "cNvPr") {
                    let chart_node = frame
                        .descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "chart");

                    let graphic_data_is_chart = frame
                        .descendants()
                        .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
                        .and_then(|n| n.attribute("uri"))
                        .is_some_and(|uri| uri == GRAPHIC_DATA_CHART_URI);

                    let size =
                        extract_size_from_transform(&frame).or_else(|| size_from_anchor(anchor));

                    if chart_node.is_some() || graphic_data_is_chart {
                        if let Some(chart_rel_id) = chart_node
                            .and_then(|n| {
                                n.attribute((REL_NS, "id"))
                                    .or_else(|| n.attribute("r:id"))
                                    .or_else(|| n.attribute("id"))
                            })
                            .map(|s| s.to_string())
                        {
                            objects.push(DrawingObject {
                                id,
                                kind: DrawingObjectKind::ChartPlaceholder {
                                    rel_id: chart_rel_id,
                                    raw_xml: frame_xml,
                                },
                                anchor,
                                z_order: z as i32,
                                size,
                                preserved: anchor_preserved.clone(),
                            });
                        } else {
                            let raw_anchor =
                                slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                            objects.push(DrawingObject {
                                id,
                                kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                                anchor,
                                z_order: z as i32,
                                size,
                                preserved: anchor_preserved.clone(),
                            });
                        }
                    } else {
                        let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    }
                    continue;
                }
                // Fall back to preserving the full anchor subtree when we can't parse the frame.
            }

            // Unknown anchor type: preserve the entire anchor subtree.
            let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
            objects.push(DrawingObject {
                id: DrawingObjectId((z + 1) as u32),
                kind: DrawingObjectKind::Unknown {
                    raw_xml: raw_anchor,
                },
                anchor,
                z_order: z as i32,
                size: None,
                preserved: anchor_preserved,
            });
        }

        Ok(Self {
            sheet_index,
            path: path.to_string(),
            rels_path,
            objects,
            relationships,
            root_xmlns,
        })
    }

    pub fn create_new(sheet_index: usize) -> Result<(Self, String)> {
        // Default new part names. Callers may rename by updating `path`/`rels_path`.
        let path = format!("xl/drawings/drawing{}.xml", sheet_index + 1);
        let rels_path = drawing_rels_path(&path);

        let drawing_rel_id = "rId1".to_string();

        Ok((
            Self::new_empty(sheet_index, path, rels_path),
            drawing_rel_id,
        ))
    }

    pub fn insert_image_object(&mut self, image_id: &ImageId, anchor: Anchor) -> DrawingObject {
        let next_object_id = self
            .objects
            .iter()
            .map(|o| o.id.0)
            .max()
            .unwrap_or(0)
            .saturating_add(1);

        let embed_rel_id = self.relationships.next_r_id();
        let target = format!("../media/{}", image_id.as_str());
        self.relationships.push(Relationship {
            id: embed_rel_id.clone(),
            type_: REL_TYPE_IMAGE.to_string(),
            target,
            target_mode: None,
        });

        let size = match anchor {
            Anchor::OneCell { ext, .. } | Anchor::Absolute { ext, .. } => Some(ext),
            Anchor::TwoCell { .. } => None,
        };

        let pic_xml = build_pic_xml(next_object_id, &embed_rel_id, size);
        let mut preserved = std::collections::HashMap::new();
        preserved.insert("xlsx.embed_rel_id".to_string(), embed_rel_id);
        preserved.insert("xlsx.pic_xml".to_string(), pic_xml);

        let object = DrawingObject {
            id: DrawingObjectId(next_object_id),
            kind: DrawingObjectKind::Image {
                image_id: image_id.clone(),
            },
            anchor,
            z_order: self.objects.len() as i32,
            size,
            preserved,
        };

        self.objects.push(object.clone());
        object
    }

    pub fn write_into_parts(
        &mut self,
        parts: &mut BTreeMap<String, Vec<u8>>,
        workbook: &formula_model::Workbook,
    ) -> Result<()> {
        // Keep objects ordered by z-order.
        self.objects.sort_by_key(|o| o.z_order);

        let mut xml = String::new();
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(&build_wsdr_root_start_tag(&self.root_xmlns));

        for object in &self.objects {
            match &object.kind {
                DrawingObjectKind::Unknown { raw_xml } => {
                    xml.push_str(raw_xml);
                }
                DrawingObjectKind::Image { .. } => {
                    let pic_xml = object
                        .preserved
                        .get("xlsx.pic_xml")
                        .cloned()
                        .unwrap_or_else(|| build_pic_xml(object.id.0, "rId1", object.size));
                    xml.push_str(&build_anchor_xml(
                        &object.anchor,
                        &pic_xml,
                        &object.preserved,
                    ));
                }
                DrawingObjectKind::Shape { raw_xml } => {
                    xml.push_str(&build_anchor_xml(
                        &object.anchor,
                        raw_xml,
                        &object.preserved,
                    ));
                }
                DrawingObjectKind::ChartPlaceholder { raw_xml, rel_id: _ } => {
                    // Keep chart relationships as they existed in the source file.
                    xml.push_str(&build_anchor_xml(
                        &object.anchor,
                        raw_xml,
                        &object.preserved,
                    ));
                }
            }
        }

        xml.push_str("</xdr:wsDr>");
        parts.insert(self.path.clone(), xml.into_bytes());
        parts.insert(self.rels_path.clone(), self.relationships.to_xml());

        // Ensure any images referenced by relationships are present in the package.
        for rel in self.relationships.iter() {
            if rel.type_ == REL_TYPE_IMAGE {
                let target_path = resolve_target(&self.path, &rel.target);
                let filename = target_path
                    .strip_prefix("xl/media/")
                    .unwrap_or(&target_path)
                    .to_string();
                let image_id = ImageId::new(filename);
                if let Some(img) = workbook.images.get(&image_id) {
                    parts.insert(target_path, img.bytes.clone());
                }
            }
        }

        Ok(())
    }
}

const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn extract_root_xmlns(root: Node<'_, '_>) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();
    for ns in root.namespaces() {
        let prefix = ns.name().unwrap_or("").to_string();
        out.insert(prefix, ns.uri().to_string());
    }
    out
}

fn build_wsdr_root_start_tag(root_xmlns: &BTreeMap<String, String>) -> String {
    let mut xmlns = BTreeMap::new();
    // Required namespaces for drawings generated by this module.
    xmlns.insert("xdr".to_string(), XDR_NS.to_string());
    xmlns.insert("a".to_string(), A_NS.to_string());
    xmlns.insert("r".to_string(), REL_NS.to_string());

    // Merge in preserved namespaces from the parsed drawing root. When there are conflicts (same
    // prefix), prefer the preserved value to avoid changing prefix bindings in preserved raw XML.
    for (k, v) in root_xmlns {
        xmlns.insert(k.clone(), v.clone());
    }

    let mut out = String::new();
    out.push_str("<xdr:wsDr");
    for (prefix, uri) in xmlns {
        out.push(' ');
        if prefix.is_empty() {
            out.push_str("xmlns=\"");
        } else {
            out.push_str("xmlns:");
            out.push_str(&prefix);
            out.push_str("=\"");
        }
        escape_xml_attr(&mut out, &uri);
        out.push('"');
    }
    out.push('>');
    out
}

fn escape_xml_attr(out: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
}

fn read_zip_part_optional<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>> {
    match open_zip_part(archive, name) {
        Ok(mut file) => {
            if file.is_dir() {
                return Ok(None);
            }
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)?;
            Ok(Some(buf))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(err.into()),
    }
}

fn resolve_image_id_from_archive<R: Read + Seek>(
    relationships: &Relationships,
    embed_rel_id: &str,
    drawing_path: &str,
    archive: &mut ZipArchive<R>,
    workbook: &mut formula_model::Workbook,
) -> Result<ImageId> {
    let rel = relationships.get(embed_rel_id).ok_or_else(|| {
        XlsxError::Invalid(format!(
            "drawing references missing image relationship {embed_rel_id}"
        ))
    })?;
    if rel
        .target_mode
        .as_deref()
        .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
    {
        return Err(XlsxError::Invalid(format!(
            "image relationship {embed_rel_id} is external"
        )));
    }

    let target_path = resolve_target(drawing_path, &rel.target);
    let file_name = target_path
        .strip_prefix("xl/media/")
        .unwrap_or(&target_path)
        .to_string();
    let image_id = ImageId::new(file_name);

    if workbook.images.get(&image_id).is_none() {
        // Best-effort: if the media part is missing, still return the image id.
        if let Some(bytes) = read_zip_part_optional(archive, &target_path)? {
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
    }
    Ok(image_id)
}

pub fn load_media_parts(workbook: &mut formula_model::Workbook, parts: &BTreeMap<String, Vec<u8>>) {
    for (path, bytes) in parts {
        let Some(file_name) = path.strip_prefix("xl/media/") else {
            continue;
        };
        let image_id = ImageId::new(file_name);
        if workbook.images.get(&image_id).is_some() {
            continue;
        }

        let ext = file_name.rsplit_once('.').map(|(_, ext)| ext).unwrap_or("");
        workbook.images.insert(
            image_id,
            ImageData {
                bytes: bytes.clone(),
                content_type: Some(content_type_for_extension(ext).to_string()),
            },
        );
    }
}

pub fn content_type_for_extension(ext: &str) -> &'static str {
    match ext.to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        _ => "application/octet-stream",
    }
}

fn drawing_rels_path(drawing_path: &str) -> String {
    let (dir, file) = drawing_path.rsplit_once('/').unwrap_or(("", drawing_path));
    let dir = if dir.is_empty() {
        "".to_string()
    } else {
        format!("{dir}/")
    };
    format!("{dir}_rels/{file}.rels")
}

fn parse_anchor(anchor_node: &roxmltree::Node<'_, '_>) -> Result<Anchor> {
    crate::drawingml::anchor::parse_anchor(anchor_node).ok_or_else(|| {
        XlsxError::Invalid(format!(
            "failed to parse anchor <{}>",
            anchor_node.tag_name().name()
        ))
    })
}

fn parse_pic(
    pic_node: &roxmltree::Node<'_, '_>,
    drawing_xml: &str,
) -> Result<(DrawingObjectId, String, String)> {
    let (id, pic_xml) = parse_named_node(pic_node, drawing_xml, "cNvPr")?;
    let blip = pic_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "blip")
        .ok_or_else(|| XlsxError::Invalid("pic missing a:blip".to_string()))?;
    let embed = blip
        .attribute((REL_NS, "embed"))
        .or_else(|| blip.attribute("r:embed"))
        .ok_or_else(|| XlsxError::Invalid("blip missing r:embed".to_string()))?
        .to_string();

    Ok((id, pic_xml, embed))
}

fn parse_named_node(
    node: &roxmltree::Node<'_, '_>,
    doc_xml: &str,
    name_tag: &str,
) -> Result<(DrawingObjectId, String)> {
    let nv_id = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == name_tag)
        .and_then(|n| n.attribute("id"))
        .unwrap_or("0");
    let id = nv_id
        .parse::<u32>()
        .map_err(|e| XlsxError::Invalid(format!("invalid object id {nv_id:?}: {e}")))?;

    let xml = slice_node_xml(node, doc_xml).unwrap_or_default();
    Ok((DrawingObjectId(id), xml))
}

fn slice_node_xml(node: &roxmltree::Node<'_, '_>, doc: &str) -> Option<String> {
    let range = node.range();
    doc.get(range).map(|s| s.to_string())
}

fn size_from_anchor(anchor: Anchor) -> Option<EmuSize> {
    match anchor {
        Anchor::OneCell { ext, .. } | Anchor::Absolute { ext, .. } => Some(ext),
        Anchor::TwoCell { .. } => None,
    }
}

/// Best-effort `a:xfrm/a:ext` extraction.
///
/// This is used to populate `DrawingObject.size`, particularly for `xdr:twoCellAnchor`
/// objects where the anchor itself does not include an `<xdr:ext>`.
fn extract_size_from_transform(node: &roxmltree::Node<'_, '_>) -> Option<EmuSize> {
    let xfrm = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "xfrm")?;
    let ext = xfrm
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "ext")?;
    let cx = ext.attribute("cx")?.parse::<i64>().ok()?;
    let cy = ext.attribute("cy")?.parse::<i64>().ok()?;
    Some(EmuSize::new(cx, cy))
}

fn resolve_image_id(
    relationships: &Relationships,
    embed_rel_id: &str,
    drawing_path: &str,
    parts: &BTreeMap<String, Vec<u8>>,
    workbook: &mut formula_model::Workbook,
) -> Result<ImageId> {
    let target = relationships.target_for(embed_rel_id).ok_or_else(|| {
        XlsxError::Invalid(format!(
            "drawing references missing image relationship {embed_rel_id}"
        ))
    })?;
    let target_path = resolve_target(drawing_path, target);
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

    Ok(image_id)
}

fn build_anchor_xml(
    anchor: &Anchor,
    inner_xml: &str,
    preserved: &std::collections::HashMap<String, String>,
) -> String {
    let mut out = String::new();

    let anchor_attrs = preserved_anchor_attrs(preserved);

    match anchor {
        Anchor::OneCell { from, ext } => {
            out.push_str("<xdr:oneCellAnchor");
            out.push_str(&anchor_attrs);
            out.push('>');
            out.push_str(&build_from_to_xml("from", from));
            out.push_str(&format!(r#"<xdr:ext cx="{}" cy="{}"/>"#, ext.cx, ext.cy));
        }
        Anchor::TwoCell { from, to } => {
            out.push_str("<xdr:twoCellAnchor");
            out.push_str(&anchor_attrs);
            out.push('>');
            out.push_str(&build_from_to_xml("from", from));
            out.push_str(&build_from_to_xml("to", to));
        }
        Anchor::Absolute { pos, ext } => {
            out.push_str("<xdr:absoluteAnchor");
            out.push_str(&anchor_attrs);
            out.push('>');
            out.push_str(&format!(
                r#"<xdr:pos x="{}" y="{}"/>"#,
                pos.x_emu, pos.y_emu
            ));
            out.push_str(&format!(r#"<xdr:ext cx="{}" cy="{}"/>"#, ext.cx, ext.cy));
        }
    }

    out.push_str(inner_xml);
    out.push_str(&preserved_client_data_xml(preserved));

    match anchor {
        Anchor::OneCell { .. } => out.push_str("</xdr:oneCellAnchor>"),
        Anchor::TwoCell { .. } => out.push_str("</xdr:twoCellAnchor>"),
        Anchor::Absolute { .. } => out.push_str("</xdr:absoluteAnchor>"),
    }

    out
}

fn build_from_to_xml(tag: &str, point: &AnchorPoint) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
        point.cell.col, point.offset.x_emu, point.cell.row, point.offset.y_emu
    )
}

fn build_pic_xml(object_id: u32, embed_rel_id: &str, size: Option<EmuSize>) -> String {
    let (cx, cy) = size.map(|s| (s.cx, s.cy)).unwrap_or((0, 0));

    format!(
        r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{object_id}" name="Picture {object_id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{embed_rel_id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"#
    )
}

fn parse_anchor_preserved(
    anchor_node: &roxmltree::Node<'_, '_>,
    doc_xml: &str,
) -> std::collections::HashMap<String, String> {
    let mut preserved = std::collections::HashMap::new();

    let mut attrs: BTreeMap<String, String> = BTreeMap::new();
    for attr in anchor_node.attributes() {
        // Skip namespace declarations since the drawing serializer always emits the worksheet
        // drawing namespaces on `<xdr:wsDr>`.
        if attr.name() == "xmlns"
            || attr
                .namespace()
                .is_some_and(|ns| ns == "http://www.w3.org/2000/xmlns/")
        {
            continue;
        }
        attrs.insert(attr.name().to_string(), attr.value().to_string());
    }

    if let Some(edit_as) = anchor_node.attribute("editAs") {
        preserved.insert("xlsx.anchor_edit_as".to_string(), edit_as.to_string());
    }
    if !attrs.is_empty() {
        if let Ok(json) = serde_json::to_string(&attrs) {
            preserved.insert("xlsx.anchor_attrs".to_string(), json);
        }
    }

    if let Some(client_data) = anchor_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "clientData")
        .and_then(|n| slice_node_xml(&n, doc_xml))
    {
        preserved.insert("xlsx.client_data_xml".to_string(), client_data);
    }

    preserved
}

fn preserved_anchor_attrs(preserved: &std::collections::HashMap<String, String>) -> String {
    // Best-effort attribute preservation: if present, we store a JSON map of attribute key/value
    // pairs from the original `<xdr:*Anchor>` element.
    let Some(attrs_json) = preserved.get("xlsx.anchor_attrs") else {
        // Back-compat: emit at least `editAs` if it was captured separately.
        if let Some(edit_as) = preserved.get("xlsx.anchor_edit_as") {
            return format!(r#" editAs="{}""#, escape_attr_value(edit_as));
        }
        return String::new();
    };

    let Ok(attrs) = serde_json::from_str::<BTreeMap<String, String>>(attrs_json) else {
        return String::new();
    };

    let mut out = String::new();
    for (k, v) in attrs {
        out.push(' ');
        out.push_str(&k);
        out.push_str(r#"=""#);
        out.push_str(&escape_attr_value(&v));
        out.push('"');
    }
    out
}

fn preserved_client_data_xml(preserved: &std::collections::HashMap<String, String>) -> String {
    match preserved.get("xlsx.client_data_xml") {
        Some(xml) => client_data_with_xdr_prefix(xml),
        None => "<xdr:clientData/>".to_string(),
    }
}

fn client_data_with_xdr_prefix(raw: &str) -> String {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return "<xdr:clientData/>".to_string();
    }

    if trimmed.starts_with("<xdr:clientData") {
        return trimmed.to_string();
    }

    // Some producers use the spreadsheetDrawing namespace as a default namespace (no prefix),
    // resulting in `<clientData .../>`. Our drawing writer always emits `xdr:` prefixed tags, so
    // rewrite to match.
    if trimmed.starts_with("<clientData")
        || trimmed.starts_with("<x:clientData")
        || trimmed.starts_with("<d:clientData")
    {
        // Best-effort: rewrite both start and end tags if present.
        // This is safe for the empty-element form (`<clientData .../>`) which is the common case.
        let mut out = trimmed.to_string();
        out = out.replacen("<clientData", "<xdr:clientData", 1);
        out = out.replacen("<x:clientData", "<xdr:clientData", 1);
        out = out.replacen("<d:clientData", "<xdr:clientData", 1);
        out = out.replace("</clientData>", "</xdr:clientData>");
        out = out.replace("</x:clientData>", "</xdr:clientData>");
        out = out.replace("</d:clientData>", "</xdr:clientData>");
        return out;
    }

    // Fall back to the raw fragment.
    trimmed.to_string()
}

fn escape_attr_value(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}
