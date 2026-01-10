use roxmltree::Document;

use crate::workbook::ChartExtractionError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub type_: String,
    pub target: String,
}

pub fn parse_relationships(
    xml: &[u8],
    part_name: &str,
) -> Result<Vec<Relationship>, ChartExtractionError> {
    let xml = std::str::from_utf8(xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc = Document::parse(xml)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

    let mut rels = Vec::new();
    for node in doc.descendants().filter(|n| n.is_element()) {
        if node.tag_name().name() != "Relationship" {
            continue;
        }

        let id = match node.attribute("Id") {
            Some(id) => id.to_string(),
            None => continue,
        };
        let type_ = node.attribute("Type").unwrap_or_default().to_string();
        let target = node.attribute("Target").unwrap_or_default().to_string();
        rels.push(Relationship { id, type_, target });
    }

    Ok(rels)
}
