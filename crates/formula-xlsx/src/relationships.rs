use std::collections::HashMap;

use roxmltree::Document;

use crate::workbook::ChartExtractionError;
use crate::XlsxError;

pub const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A single OPC relationship entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub type_: String,
    pub target: String,
}

/// Parse a `.rels` part using the chart extraction error type.
///
/// This function is used by the chart subsystem and intentionally does not
/// depend on the higher-level [`XlsxError`] to avoid mixing concerns.
pub fn parse_relationships(
    xml: &[u8],
    part_name: &str,
) -> Result<Vec<Relationship>, ChartExtractionError> {
    let xml = std::str::from_utf8(xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc =
        Document::parse(xml).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

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

/// Convenience wrapper for working with `.rels` entries when editing parts.
#[derive(Debug, Clone, Default)]
pub struct Relationships {
    rels: Vec<Relationship>,
    by_id: HashMap<String, usize>,
}

impl Relationships {
    pub fn from_xml(xml: &str) -> Result<Self, XlsxError> {
        let doc = Document::parse(xml)?;
        let mut rels = Vec::new();

        for node in doc.descendants().filter(|n| n.is_element()) {
            if node.tag_name().name() != "Relationship" {
                continue;
            }

            let id = node
                .attribute("Id")
                .ok_or_else(|| XlsxError::MissingAttr("Id"))?
                .to_string();
            let type_ = node
                .attribute("Type")
                .ok_or_else(|| XlsxError::MissingAttr("Type"))?
                .to_string();
            let target = node
                .attribute("Target")
                .ok_or_else(|| XlsxError::MissingAttr("Target"))?
                .to_string();

            rels.push(Relationship { id, type_, target });
        }

        Ok(Self::new(rels))
    }

    pub fn new(rels: Vec<Relationship>) -> Self {
        let mut by_id = HashMap::with_capacity(rels.len());
        for (idx, rel) in rels.iter().enumerate() {
            by_id.insert(rel.id.clone(), idx);
        }
        Self { rels, by_id }
    }

    pub fn target_for(&self, id: &str) -> Option<&str> {
        self.by_id
            .get(id)
            .and_then(|idx| self.rels.get(*idx))
            .map(|r| r.target.as_str())
    }

    pub fn get(&self, id: &str) -> Option<&Relationship> {
        self.by_id.get(id).and_then(|idx| self.rels.get(*idx))
    }

    pub fn iter(&self) -> impl Iterator<Item = &Relationship> {
        self.rels.iter()
    }

    pub fn push(&mut self, rel: Relationship) {
        let idx = self.rels.len();
        self.by_id.insert(rel.id.clone(), idx);
        self.rels.push(rel);
    }

    pub fn next_r_id(&self) -> String {
        let mut max = 0u32;
        for rel in &self.rels {
            if let Some(suffix) = rel.id.strip_prefix("rId") {
                if let Ok(n) = suffix.parse::<u32>() {
                    max = max.max(n);
                }
            }
        }
        format!("rId{}", max + 1)
    }

    pub fn to_xml(&self) -> Vec<u8> {
        let mut out = String::new();
        out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        out.push_str(&format!(r#"<Relationships xmlns="{PACKAGE_REL_NS}">"#));
        for rel in &self.rels {
            out.push_str(&format!(
                r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                xml_escape(&rel.id),
                xml_escape(&rel.type_),
                xml_escape(&rel.target)
            ));
        }
        out.push_str("</Relationships>");
        out.into_bytes()
    }
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

