use std::collections::HashMap;
use std::io::Cursor;

use formula_model::{CellRef, HyperlinkTarget};
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::path::{rels_for_part, resolve_target};
use crate::{parse_worksheet_hyperlinks, XlsxError, XlsxPackage};

const RICH_VALUE_REL_PART: &str = "xl/richData/richValueRel.xml";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// An embedded image stored inside a cell using Excel's RichData / `vm=` mechanism.
///
/// These are distinct from DrawingML images (floating/anchored shapes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedCellImage {
    /// Resolved package part name for the image (e.g. `xl/media/image1.png`).
    pub image_part: String,
    /// Raw bytes for the image file.
    pub image_bytes: Vec<u8>,
    /// Optional hyperlink target attached to the same worksheet cell.
    pub hyperlink_target: Option<HyperlinkTarget>,
}

impl XlsxPackage {
    /// Extract embedded-in-cell images from the workbook package.
    ///
    /// Excel stores "in-cell" images by:
    /// - marking a cell as an error (`t="e"`) with a `vm="N"` attribute
    /// - storing RichData mappings in `xl/richData/richValueRel.xml`
    /// - resolving those `<rel r:id="...">` entries via `xl/richData/_rels/richValueRel.xml.rels`
    ///
    /// This API returns a mapping keyed by `(worksheet_part, cell_ref)`.
    pub fn extract_embedded_cell_images(
        &self,
    ) -> Result<HashMap<(String, CellRef), EmbeddedCellImage>, XlsxError> {
        let Some(rich_value_rel_bytes) = self.part(RICH_VALUE_REL_PART) else {
            // Workbooks without in-cell images omit the entire `xl/richData/` tree.
            return Ok(HashMap::new());
        };

        let rich_value_rel_xml = std::str::from_utf8(rich_value_rel_bytes)
            .map_err(|e| XlsxError::Invalid(format!("{RICH_VALUE_REL_PART} not utf-8: {e}")))?;
        let rich_value_rel_ids = parse_rich_value_rel_ids(rich_value_rel_xml)?;

        let rich_value_rels_part = rels_for_part(RICH_VALUE_REL_PART);
        let Some(rich_value_rel_rels_bytes) = self.part(&rich_value_rels_part) else {
            // If the richValueRel part exists, we expect its .rels as well. Be defensive and
            // treat a missing rels part as "no images" rather than erroring.
            return Ok(HashMap::new());
        };

        let image_targets_by_rel_id =
            parse_rich_value_rel_image_targets(rich_value_rel_rels_bytes)?;

        let mut out = HashMap::new();
        for sheet in self.worksheet_parts()? {
            let Some(sheet_xml_bytes) = self.part(&sheet.worksheet_part) else {
                continue;
            };
            let sheet_xml = std::str::from_utf8(sheet_xml_bytes).map_err(|e| {
                XlsxError::Invalid(format!("{} not utf-8: {e}", sheet.worksheet_part))
            })?;

            let sheet_rels_part = rels_for_part(&sheet.worksheet_part);
            let sheet_rels_xml = self
                .part(&sheet_rels_part)
                .and_then(|bytes| std::str::from_utf8(bytes).ok());

            // Best-effort: if hyperlink parsing fails (malformed file), still extract images.
            let hyperlinks =
                parse_worksheet_hyperlinks(sheet_xml, sheet_rels_xml).unwrap_or_default();

            for (cell_ref, vm) in parse_sheet_vm_cells(sheet_xml)? {
                let Some(image_part) =
                    resolve_rich_value_vm_to_image_part(&rich_value_rel_ids, &image_targets_by_rel_id, vm)
                else {
                    continue;
                };

                let Some(image_bytes) = self.part(&image_part) else {
                    continue;
                };

                let hyperlink_target = hyperlinks
                    .iter()
                    .find(|link| link.range.contains(cell_ref))
                    .map(|link| link.target.clone());

                out.insert(
                    (sheet.worksheet_part.clone(), cell_ref),
                    EmbeddedCellImage {
                        image_part,
                        image_bytes: image_bytes.to_vec(),
                        hyperlink_target,
                    },
                );
            }
        }

        Ok(out)
    }
}

fn parse_rich_value_rel_ids(xml: &str) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"rel" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r:id" | b"id" => {
                            out.push(attr.unescape_value()?.into_owned());
                            break;
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rich_value_rel_image_targets(rels_xml: &[u8]) -> Result<HashMap<String, String>, XlsxError> {
    let relationships = crate::openxml::parse_relationships(rels_xml)?;
    let mut out = HashMap::new();
    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_uri != REL_TYPE_IMAGE {
            continue;
        }
        let target = strip_fragment(&rel.target);
        if target.is_empty() {
            continue;
        }
        let resolved = resolve_target(RICH_VALUE_REL_PART, target);
        out.insert(rel.id, resolved);
    }
    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

fn parse_sheet_vm_cells(sheet_xml: &str) -> Result<Vec<(CellRef, u32)>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(sheet_xml.as_bytes()));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                let mut cell_ref: Option<CellRef> = None;
                let mut vm: Option<u32> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            let parsed = CellRef::from_a1(&a1)
                                .map_err(|e| XlsxError::Invalid(format!("invalid cell ref {a1}: {e}")))?;
                            cell_ref = Some(parsed);
                        }
                        b"vm" => {
                            vm = attr
                                .unescape_value()?
                                .into_owned()
                                .trim()
                                .parse::<u32>()
                                .ok();
                        }
                        _ => {}
                    }
                }

                if let (Some(cell_ref), Some(vm)) = (cell_ref, vm) {
                    out.push((cell_ref, vm));
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn resolve_rich_value_vm_to_image_part(
    rich_value_rel_ids: &[String],
    image_targets_by_rel_id: &HashMap<String, String>,
    vm: u32,
) -> Option<String> {
    if vm == 0 {
        return None;
    }
    let idx = (vm as usize).checked_sub(1)?;
    let rel_id = rich_value_rel_ids.get(idx)?;
    image_targets_by_rel_id.get(rel_id).cloned()
}
