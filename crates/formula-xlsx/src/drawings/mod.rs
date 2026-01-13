//! DrawingML (images, shapes, charts) parsing and serialization.

mod part;

pub use part::*;

use std::collections::HashMap;

use formula_model::drawings::DrawingObject;
use roxmltree::Document;

use crate::path::{rels_for_part, resolve_target};
use crate::{XlsxError, XlsxPackage};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parsed DrawingML objects for a single sheet drawing part (`xl/drawings/*.xml`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtractedDrawingObjects {
    /// Workbook sheet index (0-based, ordered as in `xl/workbook.xml`).
    pub sheet_index: usize,
    /// Workbook sheet name.
    pub sheet_name: String,
    /// Worksheet XML part name (e.g. `xl/worksheets/sheet1.xml`).
    pub sheet_part: String,
    /// Drawing part name (e.g. `xl/drawings/drawing1.xml`).
    pub drawing_part: String,
    /// Parsed objects contained in the drawing part.
    pub objects: Vec<DrawingObject>,
}

impl XlsxPackage {
    /// Extract per-sheet DrawingML objects (images, shapes, chart placeholders) from `xl/drawings/*.xml`.
    ///
    /// This is a low-level API that:
    /// - walks workbook sheets (`xl/workbook.xml` + `xl/_rels/workbook.xml.rels`),
    /// - finds `worksheet/<drawing r:id="...">` references,
    /// - resolves those references via the worksheet `.rels`,
    /// - and parses each drawing part into [`DrawingObject`] entries with anchor/z-order info.
    pub fn extract_drawing_objects(&self) -> Result<Vec<ExtractedDrawingObjects>, XlsxError> {
        let sheets = self.worksheet_parts()?;

        // `DrawingPart::parse_from_parts` needs a workbook for image resolution. The extracted
        // `DrawingObjectKind::Image` values only retain an `ImageId`, but we still populate the
        // shared image store while parsing so `DrawingPart` behaves consistently with the
        // higher-level import pipeline.
        let mut workbook = formula_model::Workbook::default();

        let mut out = Vec::new();
        for (sheet_index, sheet) in sheets.iter().enumerate() {
            let Some(worksheet_xml) = self.part(&sheet.worksheet_part) else {
                continue;
            };
            let worksheet_xml = std::str::from_utf8(worksheet_xml)
                .map_err(|e| XlsxError::Invalid(format!("worksheet xml not utf-8: {e}")))?;
            let doc = Document::parse(worksheet_xml)?;

            let drawing_rids: Vec<String> = doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
                .map(|s| s.to_string())
                .collect();

            if drawing_rids.is_empty() {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet.worksheet_part);
            let Some(sheet_rels_bytes) = self.part(&sheet_rels_part) else {
                continue;
            };
            let relationships = crate::openxml::parse_relationships(sheet_rels_bytes)?;
            let rel_by_id: HashMap<String, crate::openxml::Relationship> = relationships
                .into_iter()
                .map(|rel| (rel.id.clone(), rel))
                .collect();

            for drawing_rid in drawing_rids {
                let Some(rel) = rel_by_id.get(&drawing_rid) else {
                    continue;
                };
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                let drawing_part = resolve_target(&sheet.worksheet_part, &rel.target);
                if self.part(&drawing_part).is_none() {
                    continue;
                }

                let drawing_part_parsed = DrawingPart::parse_from_parts(
                    sheet_index,
                    &drawing_part,
                    self.parts_map(),
                    &mut workbook,
                )?;

                out.push(ExtractedDrawingObjects {
                    sheet_index,
                    sheet_name: sheet.name.clone(),
                    sheet_part: sheet.worksheet_part.clone(),
                    drawing_part,
                    objects: drawing_part_parsed.objects,
                });
            }
        }

        Ok(out)
    }
}
