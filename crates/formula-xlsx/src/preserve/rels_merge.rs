use std::collections::HashMap;

use crate::path::resolve_target;
use crate::relationships::{parse_relationships, Relationship, Relationships};
use crate::workbook::ChartExtractionError;

/// Minimal metadata needed to re-attach preserved relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RelationshipStub {
    pub rel_id: String,
    pub target: String,
}

pub(crate) fn ensure_rels_has_relationships(
    rels_xml: Option<&[u8]>,
    part_name: &str,
    base_part: &str,
    rel_type: &str,
    relationships: &[RelationshipStub],
) -> Result<(Vec<u8>, HashMap<String, String>), ChartExtractionError> {
    if relationships.is_empty() {
        return Ok((rels_xml.unwrap_or_default().to_vec(), HashMap::new()));
    }

    let mut xml = match rels_xml {
        Some(bytes) => std::str::from_utf8(bytes)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
            .to_string(),
        None => String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n</Relationships>\n",
        ),
    };

    let existing_rels = match rels_xml {
        Some(bytes) => parse_relationships(bytes, part_name)?,
        None => Vec::new(),
    };
    let mut rels = Relationships::new(existing_rels);

    let mut id_map: HashMap<String, String> = HashMap::new();
    let mut to_insert: Vec<Relationship> = Vec::new();

    for relationship in relationships {
        let desired_id = relationship.rel_id.as_str();
        let desired_target = relationship.target.as_str();

        if let Some(mapped) = id_map.get(desired_id) {
            // We've already allocated a stable replacement for this ID in this scope.
            // Ensure the relationship exists in the output but don't allocate again.
            if rels.get(mapped).is_none() {
                let rel = Relationship {
                    id: mapped.clone(),
                    type_: rel_type.to_string(),
                    target: desired_target.to_string(),
                };
                rels.push(rel.clone());
                to_insert.push(rel);
            }
            continue;
        }

        let final_id = match rels.get(desired_id) {
            None => desired_id.to_string(),
            Some(existing)
                if existing.type_ == rel_type
                    && resolve_target(base_part, &existing.target)
                        == resolve_target(base_part, desired_target) =>
            {
                desired_id.to_string()
            }
            Some(_) => {
                let new_id = rels.next_r_id();
                id_map.insert(desired_id.to_string(), new_id.clone());
                new_id
            }
        };

        if rels.get(&final_id).is_some() {
            continue;
        }

        let rel = Relationship {
            id: final_id.clone(),
            type_: rel_type.to_string(),
            target: desired_target.to_string(),
        };
        rels.push(rel.clone());
        to_insert.push(rel);
    }

    if !to_insert.is_empty() {
        let insert_idx = xml.rfind("</Relationships>").ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: missing </Relationships>"))
        })?;

        let mut insertion = String::new();
        for rel in &to_insert {
            insertion.push_str(&format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>\n",
                xml_escape(&rel.id),
                xml_escape(&rel.type_),
                xml_escape(&rel.target)
            ));
        }
        xml.insert_str(insert_idx, &insertion);
    }

    Ok((xml.into_bytes(), id_map))
}

pub(crate) fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

