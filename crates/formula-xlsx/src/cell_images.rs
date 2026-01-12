use std::collections::{BTreeMap, HashMap, HashSet};

use formula_model::drawings::{ImageData, ImageId};
use formula_model::Workbook;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::path::{rels_for_part, resolve_target};

const CELLIMAGES_PART: &str = "xl/cellimages.xml";

/// Best-effort loader for Excel "in-cell" images.
///
/// Excel stores the image catalog in `xl/cellimages.xml`, with image payloads
/// referenced via the part's relationships (`xl/_rels/cellimages.xml.rels`).
///
/// This is intentionally best-effort:
/// - Missing `cellimages.xml` → no-op
/// - Missing `.rels` → no-op
/// - Missing referenced media part → skip that image
/// - Parse errors → no-op
pub fn load_cell_images_from_parts(parts: &BTreeMap<String, Vec<u8>>, workbook: &mut Workbook) {
    let Some(cellimages_xml) = parts.get(CELLIMAGES_PART) else {
        return;
    };

    let rels_part = rels_for_part(CELLIMAGES_PART);
    let Some(rels_xml) = parts.get(&rels_part) else {
        return;
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return,
    };

    let mut target_by_id: HashMap<String, String> = HashMap::with_capacity(relationships.len());
    for rel in relationships {
        target_by_id.insert(rel.id, rel.target);
    }

    let referenced_ids = extract_relationship_ids(cellimages_xml);
    if referenced_ids.is_empty() {
        return;
    }

    for rid in referenced_ids {
        let Some(target) = target_by_id.get(&rid) else {
            continue;
        };
        let target_path = resolve_target(CELLIMAGES_PART, target);

        let Some(bytes) = parts.get(&target_path) else {
            continue;
        };

        let file_name = target_path
            .strip_prefix("xl/media/")
            .or_else(|| target_path.strip_prefix("media/"))
            .unwrap_or(&target_path)
            .to_string();
        let image_id = ImageId::new(file_name);
        if workbook.images.get(&image_id).is_some() {
            continue;
        }

        let ext = image_id.as_str().rsplit_once('.').map(|(_, ext)| ext).unwrap_or("");
        let content_type = crate::drawings::content_type_for_extension(ext).to_string();
        workbook.images.insert(
            image_id,
            ImageData {
                bytes: bytes.clone(),
                content_type: Some(content_type),
            },
        );
    }
}

fn extract_relationship_ids(xml: &[u8]) -> HashSet<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = HashSet::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes() {
                    let Ok(attr) = attr else {
                        continue;
                    };
                    let key_local = crate::openxml::local_name(attr.key.as_ref());
                    // Relationship attributes usually show up as `r:id` or `r:embed`.
                    if !key_local.eq_ignore_ascii_case(b"id")
                        && !key_local.eq_ignore_ascii_case(b"embed")
                    {
                        continue;
                    }
                    let Ok(value) = attr.unescape_value() else {
                        continue;
                    };
                    let value = value.into_owned();
                    if value.starts_with("rId") {
                        out.insert(value);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}
