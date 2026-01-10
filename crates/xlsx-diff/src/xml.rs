use std::collections::BTreeMap;
use std::fmt;

use anyhow::{Context, Result};
use roxmltree::{Document, Node};

use crate::Severity;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NormalizedXml {
    pub root: XmlNode,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XmlNode {
    Element(XmlElement),
    Text(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XmlElement {
    pub name: QName,
    pub attrs: BTreeMap<QName, String>,
    pub children: Vec<XmlNode>,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct QName {
    pub ns: Option<String>,
    pub local: String,
}

impl fmt::Display for QName {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if let Some(ns) = &self.ns {
            write!(f, "{{{}}}{}", ns, self.local)
        } else {
            f.write_str(&self.local)
        }
    }
}

impl NormalizedXml {
    pub fn parse(part_name: &str, bytes: &[u8]) -> Result<Self> {
        let text = std::str::from_utf8(bytes)
            .with_context(|| format!("part {part_name} is not valid UTF-8"))?;
        let doc = Document::parse(text).with_context(|| format!("parse xml for {part_name}"))?;
        let root = doc.root_element();

        Ok(Self {
            root: build_node(root, false, part_name),
        })
    }
}

fn build_node(node: Node<'_, '_>, preserve_space: bool, part_name: &str) -> XmlNode {
    match node.node_type() {
        roxmltree::NodeType::Element => {
            let element = build_element(node, preserve_space, part_name);
            XmlNode::Element(element)
        }
        roxmltree::NodeType::Text => {
            let mut text = node.text().unwrap_or_default().replace("\r\n", "\n");
            if !preserve_space && text.trim().is_empty() {
                XmlNode::Text(String::new())
            } else {
                // Do not trim text: whitespace can be significant (e.g. shared strings).
                XmlNode::Text(std::mem::take(&mut text))
            }
        }
        _ => XmlNode::Text(String::new()),
    }
}

fn build_element(node: Node<'_, '_>, preserve_space: bool, part_name: &str) -> XmlElement {
    let name = QName {
        ns: node.tag_name().namespace().map(|s| s.to_string()),
        local: node.tag_name().name().to_string(),
    };

    let mut attrs: BTreeMap<QName, String> = BTreeMap::new();
    let mut xml_space_preserve = preserve_space;

    for attr in node.attributes() {
        // Ignore namespace declaration attributes ("xmlns" / "xmlns:*"). Namespace
        // differences are represented via resolved URIs on element/attribute names.
        if attr.name() == "xmlns"
            || attr
                .namespace()
                .is_some_and(|ns| ns == "http://www.w3.org/2000/xmlns/")
        {
            continue;
        }

        let qname = QName {
            ns: attr.namespace().map(|s| s.to_string()),
            local: attr.name().to_string(),
        };

        if qname.ns.as_deref() == Some("http://www.w3.org/XML/1998/namespace")
            && qname.local == "space"
        {
            if attr.value() == "preserve" {
                xml_space_preserve = true;
            }
        }

        attrs.insert(qname, attr.value().to_string());
    }

    let mut children: Vec<XmlNode> = node
        .children()
        .filter_map(|child| match child.node_type() {
            roxmltree::NodeType::Element => Some(build_node(child, xml_space_preserve, part_name)),
            roxmltree::NodeType::Text => {
                let built = build_node(child, xml_space_preserve, part_name);
                match &built {
                    XmlNode::Text(t) if t.is_empty() => None,
                    _ => Some(built),
                }
            }
            _ => None,
        })
        .collect();

    normalize_child_order(&name, &mut children, part_name);

    XmlElement {
        name,
        attrs,
        children,
    }
}

fn normalize_child_order(parent: &QName, children: &mut Vec<XmlNode>, part_name: &str) {
    // Relationships: order is not meaningful; sort by Id/Type/Target so
    // semantically identical files don't churn.
    if parent.local == "Relationships"
        && parent.ns.as_deref()
            == Some("http://schemas.openxmlformats.org/package/2006/relationships")
    {
        children.sort_by(|a, b| {
            let ka = relationship_sort_key(a);
            let kb = relationship_sort_key(b);
            ka.cmp(&kb)
        });
        return;
    }

    // [Content_Types].xml: order isn't meaningful; sort by element + key attribute.
    if part_name == "[Content_Types].xml"
        && parent.local == "Types"
        && parent.ns.as_deref()
            == Some("http://schemas.openxmlformats.org/package/2006/content-types")
    {
        children.sort_by(|a, b| {
            let ka = content_type_sort_key(a);
            let kb = content_type_sort_key(b);
            ka.cmp(&kb)
        });
    }
}

fn relationship_sort_key(node: &XmlNode) -> (String, String, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "Relationship" => (
            el.attrs
                .iter()
                .find(|(k, _)| k.local == "Id")
                .map(|(_, v)| v.clone())
                .unwrap_or_default(),
            el.attrs
                .iter()
                .find(|(k, _)| k.local == "Type")
                .map(|(_, v)| v.clone())
                .unwrap_or_default(),
            el.attrs
                .iter()
                .find(|(k, _)| k.local == "Target")
                .map(|(_, v)| v.clone())
                .unwrap_or_default(),
        ),
        _ => (String::new(), String::new(), String::new()),
    }
}

fn content_type_sort_key(node: &XmlNode) -> (String, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "Default" => (
            "Default".to_string(),
            el.attrs
                .iter()
                .find(|(k, _)| k.local == "Extension")
                .map(|(_, v)| v.clone())
                .unwrap_or_default(),
        ),
        XmlNode::Element(el) if el.name.local == "Override" => (
            "Override".to_string(),
            el.attrs
                .iter()
                .find(|(k, _)| k.local == "PartName")
                .map(|(_, v)| v.clone())
                .unwrap_or_default(),
        ),
        XmlNode::Element(el) => (el.name.local.clone(), String::new()),
        _ => (String::new(), String::new()),
    }
}

#[derive(Debug, Clone)]
pub struct XmlDiff {
    pub severity: Severity,
    pub path: String,
    pub kind: String,
    pub expected: Option<String>,
    pub actual: Option<String>,
}

pub fn diff_xml(
    expected: &NormalizedXml,
    actual: &NormalizedXml,
    base_severity: Severity,
) -> Vec<XmlDiff> {
    let mut diffs = Vec::new();
    diff_node(
        &expected.root,
        &actual.root,
        base_severity,
        "/".to_string(),
        &mut diffs,
    );
    diffs
}

fn diff_node(
    expected: &XmlNode,
    actual: &XmlNode,
    severity: Severity,
    path: String,
    diffs: &mut Vec<XmlDiff>,
) {
    match (expected, actual) {
        (XmlNode::Text(a), XmlNode::Text(b)) => {
            if a != b {
                diffs.push(XmlDiff {
                    severity,
                    path,
                    kind: "text_changed".to_string(),
                    expected: Some(truncate(a)),
                    actual: Some(truncate(b)),
                });
            }
        }
        (XmlNode::Element(a), XmlNode::Element(b)) => diff_element(a, b, severity, path, diffs),
        (XmlNode::Element(_), XmlNode::Text(b)) => diffs.push(XmlDiff {
            severity,
            path,
            kind: "node_kind_changed".to_string(),
            expected: Some("element".to_string()),
            actual: Some(format!("text({})", truncate(b))),
        }),
        (XmlNode::Text(a), XmlNode::Element(_)) => diffs.push(XmlDiff {
            severity,
            path,
            kind: "node_kind_changed".to_string(),
            expected: Some(format!("text({})", truncate(a))),
            actual: Some("element".to_string()),
        }),
    }
}

fn diff_element(
    expected: &XmlElement,
    actual: &XmlElement,
    severity: Severity,
    path: String,
    diffs: &mut Vec<XmlDiff>,
) {
    if expected.name != actual.name {
        diffs.push(XmlDiff {
            severity,
            path: path.clone(),
            kind: "element_name_changed".to_string(),
            expected: Some(expected.name.to_string()),
            actual: Some(actual.name.to_string()),
        });
        return;
    }

    for (k, expected_value) in &expected.attrs {
        match actual.attrs.get(k) {
            Some(actual_value) if actual_value == expected_value => {}
            Some(actual_value) => diffs.push(XmlDiff {
                severity,
                path: format!("{path}@{k}"),
                kind: "attribute_changed".to_string(),
                expected: Some(truncate(expected_value)),
                actual: Some(truncate(actual_value)),
            }),
            None => diffs.push(XmlDiff {
                severity,
                path: format!("{path}@{k}"),
                kind: "attribute_missing".to_string(),
                expected: Some(truncate(expected_value)),
                actual: None,
            }),
        }
    }

    for (k, actual_value) in &actual.attrs {
        if !expected.attrs.contains_key(k) {
            diffs.push(XmlDiff {
                severity,
                path: format!("{path}@{k}"),
                kind: "attribute_added".to_string(),
                expected: None,
                actual: Some(truncate(actual_value)),
            });
        }
    }

    let max_len = expected.children.len().max(actual.children.len());
    for idx in 0..max_len {
        let expected_child = expected.children.get(idx);
        let actual_child = actual.children.get(idx);
        match (expected_child, actual_child) {
            (Some(a), Some(b)) => {
                let child_path = format!(
                    "{}/{}[{}]",
                    path.trim_end_matches('/'),
                    child_name(a),
                    idx + 1
                );
                diff_node(a, b, severity, child_path, diffs);
            }
            (Some(a), None) => diffs.push(XmlDiff {
                severity,
                path: format!(
                    "{}/{}[{}]",
                    path.trim_end_matches('/'),
                    child_name(a),
                    idx + 1
                ),
                kind: "child_missing".to_string(),
                expected: Some(node_summary(a)),
                actual: None,
            }),
            (None, Some(b)) => diffs.push(XmlDiff {
                severity,
                path: format!(
                    "{}/{}[{}]",
                    path.trim_end_matches('/'),
                    child_name(b),
                    idx + 1
                ),
                kind: "child_added".to_string(),
                expected: None,
                actual: Some(node_summary(b)),
            }),
            (None, None) => {}
        }
    }
}

fn child_name(node: &XmlNode) -> String {
    match node {
        XmlNode::Element(el) => el.name.local.clone(),
        XmlNode::Text(_) => "text()".to_string(),
    }
}

fn node_summary(node: &XmlNode) -> String {
    match node {
        XmlNode::Text(t) => format!("text({})", truncate(t)),
        XmlNode::Element(el) => el.name.to_string(),
    }
}

fn truncate(value: &str) -> String {
    const MAX: usize = 120;
    if value.len() > MAX {
        format!("{}…", &value[..MAX])
    } else {
        value.to_string()
    }
}
