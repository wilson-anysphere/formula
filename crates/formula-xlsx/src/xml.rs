use std::collections::{BTreeMap, BTreeSet};

use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::Reader;
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

/// Detect the namespace prefix used for SpreadsheetML elements in a worksheet XML document.
///
/// This scans the XML until the first `<worksheet>` start/empty tag and returns its prefix, if any
/// (e.g. `Some("x")` for `<x:worksheet>`). For unprefixed worksheets (`<worksheet>`), this returns
/// `None`.
pub(crate) fn worksheet_spreadsheetml_prefix(xml: &str) -> Result<Option<String>, quick_xml::Error> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"worksheet" => {
                let name = e.name();
                let name = name.as_ref();
                let prefix = name
                    .iter()
                    .rposition(|b| *b == b':')
                    .map(|idx| &name[..idx])
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                return Ok(prefix);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

pub(crate) fn prefixed_tag(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) => format!("{prefix}:{local}"),
        None => local.to_string(),
    }
}

pub(crate) const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const OFFICE_RELATIONSHIPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct WorkbookXmlNamespaces {
    /// Prefix bound to the SpreadsheetML namespace. If SpreadsheetML is the default namespace,
    /// this will be `None`.
    pub spreadsheetml_prefix: Option<String>,
    /// Whether SpreadsheetML is set as the default namespace (`xmlns="…/main"`).
    pub spreadsheetml_is_default: bool,
    /// Prefix bound to the officeDocument relationships namespace
    /// (`http://schemas.openxmlformats.org/officeDocument/2006/relationships`).
    pub office_relationships_prefix: Option<String>,
}

pub(crate) fn workbook_xml_namespaces_from_workbook_start(
    e: &BytesStart<'_>,
) -> Result<WorkbookXmlNamespaces, quick_xml::Error> {
    let mut spreadsheetml_is_default = false;
    let mut spreadsheetml_prefix_decl: Option<String> = None;
    let mut rels_prefix_decl: Option<String> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let value = attr.value.as_ref();

        if key == b"xmlns" && value == SPREADSHEETML_NS.as_bytes() {
            spreadsheetml_is_default = true;
            continue;
        }

        if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            if value == SPREADSHEETML_NS.as_bytes() {
                spreadsheetml_prefix_decl = Some(String::from_utf8_lossy(prefix).into_owned());
            } else if value == OFFICE_RELATIONSHIPS_NS.as_bytes() {
                rels_prefix_decl = Some(String::from_utf8_lossy(prefix).into_owned());
            }
        }
    }

    let name = e.name();
    let name = name.as_ref();
    let element_prefix = name
        .iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
        .map(|bytes| String::from_utf8_lossy(bytes).into_owned());

    // Prefer the prefix used by the `<workbook>` element itself when present.
    //
    // Some producers declare multiple prefixes bound to the SpreadsheetML namespace
    // (`xmlns:x="…/main" xmlns:y="…/main"`). When we synthesize new SpreadsheetML elements (e.g.
    // `<calcPr/>`, `<workbookPr/>`) we want to follow the file's established style by using the
    // same prefix as `<workbook>` (if it's prefixed), rather than whichever `xmlns:*` declaration
    // happened to appear last.
    let spreadsheetml_prefix = if element_prefix.is_some() {
        element_prefix
    } else if spreadsheetml_is_default {
        None
    } else {
        spreadsheetml_prefix_decl
    };

    Ok(WorkbookXmlNamespaces {
        spreadsheetml_prefix,
        spreadsheetml_is_default,
        office_relationships_prefix: rels_prefix_decl,
    })
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
