use std::collections::BTreeMap;
use std::io::{Read, Seek};

use formula_model::drawings::{
    Anchor, AnchorPoint, DrawingObject, DrawingObjectId, DrawingObjectKind, EmuSize, ImageData,
    ImageId,
};
use quick_xml::events::Event as XmlEvent;
use quick_xml::Reader as XmlReader;
use roxmltree::{Document, Node};

use crate::path::resolve_target;
use crate::relationships::{Relationship, Relationships};
use crate::XlsxError;
use zip::ZipArchive;

type Result<T> = std::result::Result<T, XlsxError>;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const GRAPHIC_DATA_CHART_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

fn is_pic_node(node: Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().name() == "pic"
}

fn is_sp_node(node: Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().name() == "sp"
}

fn is_chart_node(node: Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().name() == "chart"
}

fn is_graphic_frame_node(node: Node<'_, '_>) -> bool {
    node.is_element() && node.tag_name().name() == "graphicFrame"
}

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
    /// Non-namespace attributes found on the root `<xdr:wsDr>` element.
    ///
    /// Office producers sometimes add additional root-level attributes (e.g. `mc:Ignorable`) that
    /// are required for proper Markup Compatibility behavior. Preserve them when round-tripping
    /// existing drawing parts.
    root_attrs: BTreeMap<String, String>,
}

impl DrawingPart {
    /// Compute the `.rels` part path for a drawing part (e.g. `xl/drawings/_rels/drawing1.xml.rels`).
    pub fn rels_path_for(drawing_path: &str) -> String {
        drawing_rels_path(drawing_path)
    }

    pub fn new_empty(sheet_index: usize, path: String, rels_path: String) -> Self {
        Self {
            sheet_index,
            path,
            rels_path,
            objects: Vec::new(),
            relationships: Relationships::default(),
            root_xmlns: BTreeMap::new(),
            root_attrs: BTreeMap::new(),
        }
    }

    /// Construct a [`DrawingPart`] from an in-memory object list.
    ///
    /// This is primarily used by the `XlsxDocument` writer, which stores DrawingML objects on the
    /// worksheet model (`Worksheet.drawings`) and needs to (re)emit the corresponding
    /// `xl/drawings/drawingN.xml` parts when saving.
    ///
    /// Relationship ID stability:
    /// - Existing image `r:embed` relationship IDs are preserved when present in
    ///   `DrawingObject.preserved["xlsx.embed_rel_id"]`.
    /// - New image relationships are allocated via [`Relationships::next_r_id`].
    /// - If `existing_rels_xml` is provided, non-image relationships (e.g. chart references) are
    ///   preserved verbatim where possible.
    pub fn from_objects(
        sheet_index: usize,
        path: String,
        objects: Vec<DrawingObject>,
        existing_rels_xml: Option<&str>,
    ) -> Result<Self> {
        let rels_path = drawing_rels_path(&path);
        let mut relationships = existing_rels_xml
            .map(Relationships::from_xml)
            .transpose()?
            .unwrap_or_default();

        let mut objects = objects;

        // First, ensure that all explicitly-preserved embed relationship IDs exist in the
        // relationship table. This prevents `next_r_id()` from accidentally reusing a preserved
        // ID for another image.
        for object in objects.iter_mut() {
            let DrawingObjectKind::Image { image_id } = &object.kind else {
                continue;
            };

            let Some(embed_rel_id) = object.preserved.get("xlsx.embed_rel_id").cloned() else {
                continue;
            };

            let desired_target = format!("../media/{}", image_id.as_str());
            match relationships.get_mut(&embed_rel_id) {
                Some(rel) => {
                    // Explicitly ensure the image relationship has the correct type/target even
                    // when the source `.rels` omits `Type` or points at stale media.
                    rel.type_ = REL_TYPE_IMAGE.to_string();
                    rel.target = desired_target;
                    rel.target_mode = None;
                }
                None => {
                    relationships.push(Relationship {
                        id: embed_rel_id.clone(),
                        type_: REL_TYPE_IMAGE.to_string(),
                        target: desired_target,
                        target_mode: None,
                    });
                }
            }

            // If we have an embed relationship id but no preserved pic XML (e.g. objects created
            // programmatically without using `insert_image_object`), synthesize a minimal `<pic>`
            // block that references the preserved relationship id.
            object
                .preserved
                .entry("xlsx.pic_xml".to_string())
                .or_insert_with(|| build_pic_xml(object.id.0, &embed_rel_id, object.size));
        }

        // Next, allocate relationships for any image objects that are missing preserved IDs.
        for object in objects.iter_mut() {
            let DrawingObjectKind::Image { image_id } = &object.kind else {
                continue;
            };

            let embed_rel_id = object
                .preserved
                .get("xlsx.embed_rel_id")
                .cloned()
                .unwrap_or_else(|| {
                    let id = relationships.next_r_id();
                    object
                        .preserved
                        .insert("xlsx.embed_rel_id".to_string(), id.clone());
                    id
                });

            let desired_target = format!("../media/{}", image_id.as_str());
            match relationships.get_mut(&embed_rel_id) {
                Some(rel) => {
                    rel.type_ = REL_TYPE_IMAGE.to_string();
                    rel.target = desired_target;
                    rel.target_mode = None;
                }
                None => relationships.push(Relationship {
                    id: embed_rel_id.clone(),
                    type_: REL_TYPE_IMAGE.to_string(),
                    target: desired_target,
                    target_mode: None,
                }),
            }

            object
                .preserved
                .entry("xlsx.pic_xml".to_string())
                .or_insert_with(|| build_pic_xml(object.id.0, &embed_rel_id, object.size));
        }

        Ok(Self {
            sheet_index,
            path,
            rels_path,
            objects,
            relationships,
            root_xmlns: BTreeMap::new(),
            root_attrs: BTreeMap::new(),
        })
    }

    /// Variant of [`Self::from_objects`] that also preserves the namespace declarations found on
    /// the original drawing XML root (`<xdr:wsDr>`).
    ///
    /// This is important when the object list includes preserved raw XML fragments (e.g. SmartArt
    /// `dgm:*` nodes) that rely on additional root-level namespace declarations.
    pub fn from_objects_with_existing_drawing_xml(
        sheet_index: usize,
        path: String,
        objects: Vec<DrawingObject>,
        existing_drawing_xml: Option<&str>,
        existing_rels_xml: Option<&str>,
    ) -> Result<Self> {
        let mut part = Self::from_objects(sheet_index, path, objects, existing_rels_xml)?;
        if let Some(xml) = existing_drawing_xml {
            if let Ok(doc) = Document::parse(xml) {
                let root = doc.root_element();
                part.root_xmlns = extract_root_xmlns(root);
                part.root_attrs = extract_root_attrs(xml);
            }
        }
        Ok(part)
    }

    pub fn parse_from_parts(
        sheet_index: usize,
        path: &str,
        parts: &BTreeMap<String, Vec<u8>>,
        workbook: &mut formula_model::Workbook,
    ) -> Result<Self> {
        let rels_path = drawing_rels_path(path);
        // Best-effort: a drawing part may legitimately exist without a relationships
        // part (`drawingN.xml.rels`), and real-world files sometimes include malformed
        // `.rels` payloads. Treat these cases as an empty relationships set so that we
        // can still parse and preserve anchors/shapes.
        let relationships = parts
            .get(&rels_path)
            .and_then(|rels_bytes| std::str::from_utf8(rels_bytes).ok())
            .and_then(|rels_xml| Relationships::from_xml(rels_xml).ok())
            .unwrap_or_default();

        let drawing_bytes = parts
            .get(path)
            .ok_or_else(|| XlsxError::MissingPart(path.to_string()))?;
        let drawing_xml = std::str::from_utf8(drawing_bytes)
            .map_err(|e| XlsxError::Invalid(format!("drawing xml not utf-8: {e}")))?;

        let doc = Document::parse(drawing_xml)?;
        let root = doc.root_element();
        let root_xmlns = extract_root_xmlns(root);
        let root_attrs = extract_root_attrs(drawing_xml);
        let mut objects = Vec::new();

        for (z, anchor_node) in crate::drawingml::anchor::wsdr_anchor_nodes(root)
            .into_iter()
            .enumerate()
        {

            // Preserve the anchor XML for best-effort/unknown objects.
            let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
            let anchor_preserved = parse_anchor_preserved(&anchor_node, drawing_xml);

            // Best-effort: if the anchor itself is malformed we cannot construct a valid
            // `formula_model::drawings::Anchor` value, so skip it instead of aborting the
            // entire drawing part.
            let anchor = match parse_anchor(&anchor_node) {
                Ok(anchor) => anchor,
                Err(_) => continue,
            };

            if let Some(pic) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_pic_node,
            )
            .into_iter()
            .find(|n| is_pic_node(*n))
            {
                match parse_pic(&pic, drawing_xml) {
                    Ok((id, pic_xml, embed)) => {
                        match resolve_image_id(&relationships, &embed, path, parts, workbook) {
                            Ok(image_id) => {
                                let mut preserved = anchor_preserved.clone();
                                preserved.insert("xlsx.embed_rel_id".to_string(), embed.clone());
                                preserved.insert("xlsx.pic_xml".to_string(), pic_xml);

                                let size = size_from_anchor(anchor)
                                    .or_else(|| extract_size_from_transform(&pic));

                                objects.push(DrawingObject {
                                    id,
                                    kind: DrawingObjectKind::Image { image_id },
                                    anchor,
                                    z_order: z as i32,
                                    size,
                                    preserved,
                                });
                            }
                            Err(_) => {
                                // Best-effort: preserve the entire anchor subtree so we can
                                // round-trip the original XML even when the image relationship
                                // is missing/invalid.
                                let size = size_from_anchor(anchor)
                                    .or_else(|| extract_size_from_transform(&pic));
                                objects.push(DrawingObject {
                                    id,
                                    kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                                    anchor,
                                    z_order: z as i32,
                                    size,
                                    preserved: anchor_preserved.clone(),
                                });
                            }
                        }
                    }
                    Err(_) => {
                        let id = extract_drawing_object_id(&pic)
                            .unwrap_or_else(|| DrawingObjectId((z + 1) as u32));
                        let size =
                            size_from_anchor(anchor).or_else(|| extract_size_from_transform(&pic));
                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    }
                }
                continue;
            }

            if let Some(sp) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_sp_node,
            )
            .into_iter()
            .find(|n| is_sp_node(*n))
            {
                let size = size_from_anchor(anchor).or_else(|| extract_size_from_transform(&sp));
                match parse_named_node(&sp, drawing_xml, "cNvPr") {
                    Ok((id, sp_xml)) => objects.push(DrawingObject {
                        id,
                        kind: DrawingObjectKind::Shape { raw_xml: sp_xml },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    }),
                    Err(_) => objects.push(DrawingObject {
                        id: DrawingObjectId((z + 1) as u32),
                        kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    }),
                }
                continue;
            }

            if let Some(frame) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_chart_node,
            )
            .into_iter()
            .find(|n| is_graphic_frame_node(*n))
            {
                let size =
                    size_from_anchor(anchor).or_else(|| extract_size_from_transform(&frame));
                match parse_named_node(&frame, drawing_xml, "cNvPr") {
                    Ok((id, frame_xml)) => {
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
                            objects.push(DrawingObject {
                                id,
                                kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                                anchor,
                                z_order: z as i32,
                                size,
                                preserved: anchor_preserved.clone(),
                            });
                        }
                    }
                    Err(_) => objects.push(DrawingObject {
                        id: DrawingObjectId((z + 1) as u32),
                        kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    }),
                }
                continue;
            }

            // Unknown anchor type: preserve the entire anchor subtree.
            let id = extract_drawing_object_id(&anchor_node)
                .unwrap_or_else(|| DrawingObjectId((z + 1) as u32));
            let size =
                size_from_anchor(anchor).or_else(|| extract_size_from_transform(&anchor_node));
            objects.push(DrawingObject {
                id,
                kind: DrawingObjectKind::Unknown {
                    raw_xml: raw_anchor,
                },
                anchor,
                z_order: z as i32,
                size,
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
            root_attrs,
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
        let root = doc.root_element();
        let root_xmlns = extract_root_xmlns(root);
        let root_attrs = extract_root_attrs(drawing_xml);
        let mut objects = Vec::new();

        for (z, anchor_node) in crate::drawingml::anchor::wsdr_anchor_nodes(root)
            .into_iter()
            .enumerate()
        {

            let anchor = match parse_anchor(&anchor_node) {
                Ok(a) => a,
                Err(_) => continue,
            };
            let anchor_preserved = parse_anchor_preserved(&anchor_node, drawing_xml);

            if let Some(pic) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_pic_node,
            )
            .into_iter()
            .find(|n| is_pic_node(*n))
            {
                // Best-effort: pictures may reference missing relationships or media parts (or be
                // malformed). In those cases, preserve the full anchor subtree as an unknown
                // drawing object, but keep the parsed DrawingML id and any size information we can
                // extract.
                let size =
                    size_from_anchor(anchor).or_else(|| extract_size_from_transform(&pic));

                match parse_pic(&pic, drawing_xml) {
                    Ok((id, pic_xml, embed)) => {
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

                            objects.push(DrawingObject {
                                id,
                                kind: DrawingObjectKind::Image { image_id },
                                anchor,
                                z_order: z as i32,
                                size,
                                preserved,
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
                        continue;
                    }
                    Err(_) => {
                        let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                        let id = extract_drawing_object_id(&pic)
                            .unwrap_or_else(|| DrawingObjectId((z + 1) as u32));
                        objects.push(DrawingObject {
                            id,
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                        continue;
                    }
                }
            }

            if let Some(sp) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_sp_node,
            )
            .into_iter()
            .find(|n| is_sp_node(*n))
            {
                let size = size_from_anchor(anchor).or_else(|| extract_size_from_transform(&sp));

                match parse_named_node(&sp, drawing_xml, "cNvPr") {
                    Ok((id, sp_xml)) => objects.push(DrawingObject {
                        id,
                        kind: DrawingObjectKind::Shape { raw_xml: sp_xml },
                        anchor,
                        z_order: z as i32,
                        size,
                        preserved: anchor_preserved.clone(),
                    }),
                    Err(_) => {
                        // Best-effort: if we can't parse the shape id, preserve the full anchor
                        // subtree but keep any size information we can extract.
                        let raw_anchor =
                            slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                        objects.push(DrawingObject {
                            id: DrawingObjectId((z + 1) as u32),
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    }
                }
                continue;
            }

            if let Some(frame) = crate::drawingml::anchor::element_children_selecting_alternate_content(
                anchor_node,
                is_chart_node,
            )
            .into_iter()
            .find(|n| is_graphic_frame_node(*n))
            {
                let size =
                    size_from_anchor(anchor).or_else(|| extract_size_from_transform(&frame));

                match parse_named_node(&frame, drawing_xml, "cNvPr") {
                    Ok((id, frame_xml)) => {
                        let chart_node = frame
                            .descendants()
                            .find(|n| n.is_element() && n.tag_name().name() == "chart");

                        let graphic_data_is_chart = frame
                            .descendants()
                            .find(|n| n.is_element() && n.tag_name().name() == "graphicData")
                            .and_then(|n| n.attribute("uri"))
                            .is_some_and(|uri| uri == GRAPHIC_DATA_CHART_URI);

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
                    }
                    Err(_) => {
                        // Best-effort: preserve malformed/unsupported frames as unknown anchors,
                        // but keep any size information we can extract.
                        let raw_anchor =
                            slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
                        objects.push(DrawingObject {
                            id: DrawingObjectId((z + 1) as u32),
                            kind: DrawingObjectKind::Unknown { raw_xml: raw_anchor },
                            anchor,
                            z_order: z as i32,
                            size,
                            preserved: anchor_preserved.clone(),
                        });
                    }
                }
                continue;
            }

            // Unknown anchor type: preserve the entire anchor subtree.
            let raw_anchor = slice_node_xml(&anchor_node, drawing_xml).unwrap_or_default();
            let id = extract_drawing_object_id(&anchor_node)
                .unwrap_or_else(|| DrawingObjectId((z + 1) as u32));
            let size =
                size_from_anchor(anchor).or_else(|| extract_size_from_transform(&anchor_node));
            objects.push(DrawingObject {
                id,
                kind: DrawingObjectKind::Unknown {
                    raw_xml: raw_anchor,
                },
                anchor,
                z_order: z as i32,
                size,
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
            root_attrs,
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
        xml.push_str(&build_wsdr_root_start_tag(&self.root_xmlns, &self.root_attrs));

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

fn extract_root_attrs(doc_xml: &str) -> BTreeMap<String, String> {
    // Preserve the root `<xdr:wsDr>` attribute names **exactly** as they appear in the source XML
    // (including prefixes). This matches anchor attribute preservation and avoids lossy
    // reconstruction when multiple prefixes map to the same namespace URI.
    //
    // We intentionally skip `xmlns` attributes here; namespace declarations are captured separately
    // via `root.namespaces()` and re-emitted in [`build_wsdr_root_start_tag`].
    let mut out = BTreeMap::new();
    let mut reader = XmlReader::from_reader(doc_xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(XmlEvent::Start(e)) | Ok(XmlEvent::Empty(e)) => {
                for attr in e.attributes().flatten() {
                    let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                    if key == "xmlns" || key.starts_with("xmlns:") {
                        continue;
                    }
                    let value = attr.unescape_value().unwrap_or_default().into_owned();
                    out.insert(key.to_string(), value);
                }
                break;
            }
            Ok(XmlEvent::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn build_wsdr_root_start_tag(
    root_xmlns: &BTreeMap<String, String>,
    root_attrs: &BTreeMap<String, String>,
) -> String {
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
    for (k, v) in root_attrs {
        out.push(' ');
        out.push_str(k);
        out.push_str("=\"");
        escape_xml_attr(&mut out, v);
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
    crate::zip_util::read_zip_part_optional_with_limit(
        archive,
        name,
        crate::zip_util::DEFAULT_MAX_ZIP_PART_BYTES,
    )
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
        // Enhanced Metafile / Windows Metafile are commonly used by Excel for vector images.
        // They require explicit Defaults in `[Content_Types].xml` for Excel to open the package
        // without repair prompts.
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
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
            "failed to parse DrawingML anchor <{}>",
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
    // Prefer the `xdr:*` non-visual properties block (`xdr:cNvPr`) over any other `*:cNvPr`
    // elements that might appear in the subtree (some producers emit a stray `a:cNvPr` inside
    // `a:graphicData`, etc.).
    let nv_id = node
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == name_tag
                && n.tag_name().namespace() == Some(XDR_NS)
        })
        .or_else(|| node.descendants().find(|n| n.is_element() && n.tag_name().name() == name_tag))
        .and_then(|n| n.attribute("id"))
        .unwrap_or("0");
    let id = nv_id
        .parse::<u32>()
        .map_err(|e| XlsxError::Invalid(format!("invalid object id {nv_id:?}: {e}")))?;

    let xml = slice_node_xml(node, doc_xml).unwrap_or_default();
    Ok((DrawingObjectId(id), xml))
}

fn extract_drawing_object_id(node: &roxmltree::Node<'_, '_>) -> Option<DrawingObjectId> {
    let nv_id = node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "cNvPr")?
        .attribute("id")?;
    let id = nv_id.trim().parse::<u32>().ok()?;
    if id == 0 {
        return None;
    }
    Some(DrawingObjectId(id))
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

    // Best-effort: preserve the **qualified** attribute names (including any prefixes) exactly as
    // they appear in the source XML. `roxmltree` exposes only local attribute names via
    // `Attribute::name()`, so we parse the anchor element's start tag with `quick-xml` instead.
    //
    // This allows round-tripping namespaced attributes like `xdr14:anchorId`.
    let mut attrs: BTreeMap<String, String> = BTreeMap::new();
    if let Some(raw_anchor_xml) = slice_node_xml(anchor_node, doc_xml) {
        let mut reader = XmlReader::from_reader(raw_anchor_xml.as_bytes());
        reader.config_mut().trim_text(false);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(XmlEvent::Start(e)) | Ok(XmlEvent::Empty(e)) => {
                    for attr in e.attributes().flatten() {
                        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("").to_string();
                        let value = attr.unescape_value().unwrap_or_default().into_owned();
                        attrs.insert(key, value);
                    }
                    break;
                }
                Ok(XmlEvent::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
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
