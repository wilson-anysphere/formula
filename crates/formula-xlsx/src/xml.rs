use std::collections::{BTreeMap, BTreeSet};

use roxmltree::{Document, Node};

#[derive(Debug, thiserror::Error)]
pub enum XmlDomError {
    #[error("xml is not valid UTF-8: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("failed to parse xml: {0}")]
    Parse(#[from] roxmltree::Error),
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

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct QName {
    pub ns: Option<String>,
    pub local: String,
}

impl XmlElement {
    pub fn parse(bytes: &[u8]) -> Result<Self, XmlDomError> {
        let text = std::str::from_utf8(bytes)?;
        let doc = Document::parse(text)?;
        let root = doc.root_element();
        Ok(build_element(root, false))
    }

    pub fn child(&self, local: &str) -> Option<&XmlElement> {
        self.children.iter().find_map(|child| match child {
            XmlNode::Element(el) if el.name.local == local => Some(el),
            _ => None,
        })
    }

    pub fn child_mut(&mut self, local: &str) -> Option<&mut XmlElement> {
        self.children.iter_mut().find_map(|child| match child {
            XmlNode::Element(el) if el.name.local == local => Some(el),
            _ => None,
        })
    }

    pub fn children_by_local<'a>(&'a self, local: &'a str) -> impl Iterator<Item = &'a XmlElement> {
        self.children.iter().filter_map(move |child| match child {
            XmlNode::Element(el) if el.name.local == local => Some(el),
            _ => None,
        })
    }

    pub fn attr(&self, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k.ns.is_none() && k.local == local)
            .map(|(_, v)| v.as_str())
    }

    pub fn attr_ns(&self, ns: &str, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k.ns.as_deref() == Some(ns) && k.local == local)
            .map(|(_, v)| v.as_str())
    }

    pub fn set_attr(&mut self, local: &str, value: impl Into<String>) {
        self.attrs.insert(
            QName {
                ns: None,
                local: local.to_string(),
            },
            value.into(),
        );
    }

    pub fn remove_attr(&mut self, local: &str) {
        let key = self
            .attrs
            .keys()
            .find(|k| k.ns.is_none() && k.local == local)
            .cloned();
        if let Some(key) = key {
            self.attrs.remove(&key);
        }
    }

    pub fn text(&self) -> Option<&str> {
        self.children.iter().find_map(|child| match child {
            XmlNode::Text(t) => Some(t.as_str()),
            _ => None,
        })
    }

    pub fn to_xml_string(&self) -> String {
        let mut namespaces = BTreeSet::new();
        collect_namespaces(self, &mut namespaces);

        let default_ns = self.name.ns.clone();
        let mut prefixes: BTreeMap<String, String> = BTreeMap::new();
        let mut counter = 1u32;
        for ns in namespaces {
            if Some(&ns) == default_ns.as_ref() {
                continue;
            }
            let prefix = match ns.as_str() {
                "http://www.w3.org/XML/1998/namespace" => "xml".to_string(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => {
                    "r".to_string()
                }
                "http://schemas.openxmlformats.org/package/2006/relationships" => "rel".to_string(),
                _ => {
                    let p = format!("ns{counter}");
                    counter += 1;
                    p
                }
            };
            prefixes.insert(ns, prefix);
        }

        let mut out = String::new();
        out.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        out.push('\n');

        write_element(&mut out, self, &default_ns, &prefixes, true);
        out.push('\n');
        out
    }
}

fn build_node(node: Node<'_, '_>, preserve_space: bool) -> Option<XmlNode> {
    match node.node_type() {
        roxmltree::NodeType::Element => Some(XmlNode::Element(build_element(node, preserve_space))),
        roxmltree::NodeType::Text => {
            let text = node.text().unwrap_or_default().replace("\r\n", "\n");
            if preserve_space || !text.trim().is_empty() {
                Some(XmlNode::Text(text))
            } else {
                None
            }
        }
        _ => None,
    }
}

fn build_element(node: Node<'_, '_>, preserve_space: bool) -> XmlElement {
    let name = QName {
        ns: node.tag_name().namespace().map(|s| s.to_string()),
        local: node.tag_name().name().to_string(),
    };

    let mut attrs = BTreeMap::new();
    let mut xml_space_preserve = preserve_space;
    for attr in node.attributes() {
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
            && attr.value() == "preserve"
        {
            xml_space_preserve = true;
        }

        attrs.insert(qname, attr.value().to_string());
    }

    let mut children = Vec::new();
    for child in node.children() {
        if let Some(built) = build_node(child, xml_space_preserve) {
            children.push(built);
        }
    }

    XmlElement {
        name,
        attrs,
        children,
    }
}

fn collect_namespaces(el: &XmlElement, out: &mut BTreeSet<String>) {
    if let Some(ns) = &el.name.ns {
        out.insert(ns.clone());
    }
    for (k, _) in &el.attrs {
        if let Some(ns) = &k.ns {
            out.insert(ns.clone());
        }
    }
    for child in &el.children {
        if let XmlNode::Element(child_el) = child {
            collect_namespaces(child_el, out);
        }
    }
}

fn write_element(
    out: &mut String,
    el: &XmlElement,
    default_ns: &Option<String>,
    prefixes: &BTreeMap<String, String>,
    is_root: bool,
) {
    out.push('<');
    write_qname(out, &el.name, default_ns, prefixes);

    if is_root {
        if let Some(ns) = default_ns {
            out.push_str(r#" xmlns=""#);
            escape_attr(out, ns);
            out.push('"');
        }
        for (ns, prefix) in prefixes {
            out.push_str(r#" xmlns:"#);
            out.push_str(prefix);
            out.push_str(r#"=""#);
            escape_attr(out, ns);
            out.push('"');
        }
    }

    for (k, v) in &el.attrs {
        out.push(' ');
        write_qname(out, k, default_ns, prefixes);
        out.push_str(r#"=""#);
        escape_attr(out, v);
        out.push('"');
    }

    if el.children.is_empty() {
        out.push_str("/>");
        return;
    }

    out.push('>');
    for child in &el.children {
        match child {
            XmlNode::Element(child_el) => write_element(out, child_el, default_ns, prefixes, false),
            XmlNode::Text(text) => escape_text(out, text),
        }
    }
    out.push_str("</");
    write_qname(out, &el.name, default_ns, prefixes);
    out.push('>');
}

fn write_qname(
    out: &mut String,
    name: &QName,
    default_ns: &Option<String>,
    prefixes: &BTreeMap<String, String>,
) {
    match (&name.ns, default_ns) {
        (Some(ns), Some(default)) if ns == default => out.push_str(&name.local),
        (Some(ns), _) => {
            if let Some(prefix) = prefixes.get(ns) {
                out.push_str(prefix);
                out.push(':');
            }
            out.push_str(&name.local);
        }
        (None, _) => out.push_str(&name.local),
    }
}

fn escape_attr(out: &mut String, value: &str) {
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

fn escape_text(out: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
}
