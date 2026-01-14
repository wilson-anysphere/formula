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
    pub target_mode: Option<String>,
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
        let target_mode = node.attribute("TargetMode").map(|v| v.to_string());
        rels.push(Relationship {
            id,
            type_,
            target,
            target_mode,
        });
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

            // Best-effort: some producers emit incomplete relationship entries (missing `Type`).
            // For traversal/preservation, `Id` + `Target` are sufficient.
            let Some(id) = node.attribute("Id") else {
                continue;
            };
            let Some(target) = node.attribute("Target") else {
                continue;
            };
            let id = id.to_string();
            let target = target.to_string();
            let type_ = node.attribute("Type").unwrap_or_default().to_string();
            let target_mode = node.attribute("TargetMode").map(str::to_string);

            rels.push(Relationship {
                id,
                type_,
                target,
                target_mode,
            });
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

    pub fn get_mut(&mut self, id: &str) -> Option<&mut Relationship> {
        let idx = *self.by_id.get(id)?;
        self.rels.get_mut(idx)
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
            out.push_str(r#"<Relationship Id=""#);
            out.push_str(&xml_escape(&rel.id));
            out.push('"');
            if !rel.type_.is_empty() {
                out.push_str(r#" Type=""#);
                out.push_str(&xml_escape(&rel.type_));
                out.push('"');
            }
            out.push_str(r#" Target=""#);
            out.push_str(&xml_escape(&rel.target));
            out.push('"');
            if let Some(mode) = &rel.target_mode {
                out.push_str(&format!(r#" TargetMode="{}""#, xml_escape(mode)));
            }
            out.push_str("/>");
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn relationships_round_trip_preserves_target_mode() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let rels = Relationships::from_xml(xml).expect("parse rels");
        let serialized = String::from_utf8(rels.to_xml()).expect("utf8");
        assert!(
            serialized.contains(r#"TargetMode="External""#),
            "expected TargetMode to be preserved, got:\n{serialized}"
        );
    }

    #[test]
    fn relationships_from_xml_is_best_effort_about_missing_type() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image1.png" TargetMode="External"/>
  <Relationship Id="rId2" Type="urn:example:keep" Target="foo.xml"/>
</Relationships>"#;

        let rels = Relationships::from_xml(xml).expect("parse rels");
        assert_eq!(rels.iter().count(), 2);

        let r1 = rels.get("rId1").expect("rId1 present");
        assert_eq!(r1.type_, "", "missing Type should be tolerated");
        assert_eq!(r1.target, "../media/image1.png");
        assert_eq!(r1.target_mode.as_deref(), Some("External"));

        let serialized = String::from_utf8(rels.to_xml()).expect("utf8");
        assert!(
            serialized.contains(r#"Id="rId1""#) && serialized.contains(r#"Target="../media/image1.png""#),
            "expected rId1 to be preserved, got:\n{serialized}"
        );
        assert!(
            !serialized.contains(r#"Type="""#),
            "expected missing Type to remain omitted (not `Type=\"\"`), got:\n{serialized}"
        );
        assert!(
            serialized.contains(r#"TargetMode="External""#),
            "expected TargetMode to be preserved, got:\n{serialized}"
        );
        assert!(
            serialized.contains(r#"Id="rId2""#) && serialized.contains(r#"Type="urn:example:keep""#),
            "expected rId2 to be preserved, got:\n{serialized}"
        );
    }
}
