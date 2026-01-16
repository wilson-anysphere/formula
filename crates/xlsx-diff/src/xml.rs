use std::collections::BTreeMap;
use std::fmt;

use anyhow::{Context, Result};
use roxmltree::{Document, Node};

use crate::Severity;

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_DOCUMENT_REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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
        let text = crate::decode_xml_bytes(bytes)
            .with_context(|| format!("decode xml bytes for {part_name}"))?;
        let doc =
            Document::parse(text.as_ref()).with_context(|| format!("parse xml for {part_name}"))?;
        let root = doc.root_element();
        // `diff_archives` passes normalized OPC part names, but `NormalizedXml::parse` is public and
        // can be used directly by tests/consumers. Normalize here so our ordering rules are applied
        // consistently even if callers pass leading slashes, backslashes, or `..` segments.
        let normalized_part_name = crate::normalize_opc_part_name(part_name);

        Ok(Self {
            root: build_node(root, false, &normalized_part_name),
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
            let raw = node.text().unwrap_or_default();
            if !preserve_space && raw.trim().is_empty() {
                XmlNode::Text(String::new())
            } else {
                // Do not trim text: whitespace can be significant (e.g. shared strings).
                let text = if raw.contains("\r\n") {
                    raw.replace("\r\n", "\n")
                } else {
                    raw.to_string()
                };
                XmlNode::Text(text)
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

    // Worksheet cell storage order is not semantically meaningful. Normalize the most
    // common containers so diffs focus on actual content changes rather than writer
    // iteration order.
    let is_worksheet_part = part_name.starts_with("xl/worksheets/sheet") && part_name.ends_with(".xml");

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "sheetData"
    {
        children.sort_by_key(sheetdata_child_sort_key);
    }

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "cols"
    {
        children.sort_by_key(cols_child_sort_key);
    }

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "row"
    {
        children.sort_by_key(row_child_sort_key);
    }

    // Worksheet containers whose child ordering is effectively a set/map. Excel and other
    // writers frequently rewrite these in different orders.
    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "mergeCells"
    {
        sort_selected_children(children, merge_cells_child_sort_key);
    }

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "hyperlinks"
    {
        sort_selected_children(children, hyperlinks_child_sort_key);
    }

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "dataValidations"
    {
        sort_selected_children(children, data_validations_child_sort_key);
    }

    if is_worksheet_part
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "conditionalFormatting"
    {
        sort_selected_children(children, conditional_formatting_child_sort_key);
    }

    if part_name == "xl/workbook.xml"
        && parent.ns.as_deref() == Some(SPREADSHEETML_NS)
        && parent.local == "definedNames"
    {
        children.sort_by_key(defined_names_child_sort_key);
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

fn sheetdata_child_sort_key(node: &XmlNode) -> (u8, u32, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "row" => (0, row_index(el), String::new()),
        XmlNode::Element(el) => (1, u32::MAX, el.name.local.clone()),
        XmlNode::Text(_) => (2, u32::MAX, String::new()),
    }
}

fn row_child_sort_key(node: &XmlNode) -> (u8, u32, u32, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "c" => {
            let (row, col) = cell_ref_key(el);
            (0, row, col, String::new())
        }
        XmlNode::Element(el) => (1, u32::MAX, u32::MAX, el.name.local.clone()),
        XmlNode::Text(_) => (2, u32::MAX, u32::MAX, String::new()),
    }
}

fn cols_child_sort_key(node: &XmlNode) -> (u8, u32, u32, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "col" => {
            let min = attr_value(el, "min")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(u32::MAX);
            let max = attr_value(el, "max")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(u32::MAX);
            (0, min, max, String::new())
        }
        XmlNode::Element(el) => (1, u32::MAX, u32::MAX, el.name.local.clone()),
        XmlNode::Text(_) => (2, u32::MAX, u32::MAX, String::new()),
    }
}

fn sort_selected_children<K, F>(children: &mut Vec<XmlNode>, mut key_for: F)
where
    K: Ord,
    F: FnMut(&XmlNode) -> Option<K>,
{
    let mut positions = Vec::new();
    let mut keyed = Vec::new();
    for idx in 0..children.len() {
        let key = key_for(&children[idx]);
        if let Some(key) = key {
            positions.push(idx);
            let node = std::mem::replace(&mut children[idx], XmlNode::Text(String::new()));
            keyed.push((key, node));
        }
    }

    if keyed.len() >= 2 {
        keyed.sort_by(|a, b| a.0.cmp(&b.0));
    }

    for (pos, (_, node)) in positions.into_iter().zip(keyed.into_iter()) {
        children[pos] = node;
    }
}

fn merge_cells_child_sort_key(node: &XmlNode) -> Option<String> {
    match node {
        XmlNode::Element(el) if el.name.local == "mergeCell" => Some(
            attr_value(el, "ref")
                .unwrap_or_default()
                .to_string(),
        ),
        _ => None,
    }
}

fn hyperlinks_child_sort_key(node: &XmlNode) -> Option<String> {
    match node {
        XmlNode::Element(el) if el.name.local == "hyperlink" => {
            let key = attr_value(el, "ref")
                .or_else(|| attr_value_ns(el, Some(OFFICE_DOCUMENT_REL_NS), "id"))
                .unwrap_or_default()
                .to_string();
            Some(key)
        }
        _ => None,
    }
}

fn data_validations_child_sort_key(node: &XmlNode) -> Option<(String, String)> {
    match node {
        XmlNode::Element(el) if el.name.local == "dataValidation" => Some((
            attr_value(el, "sqref").unwrap_or_default().to_string(),
            attr_value(el, "type").unwrap_or_default().to_string(),
        )),
        _ => None,
    }
}

fn conditional_formatting_child_sort_key(node: &XmlNode) -> Option<(u32, String)> {
    match node {
        XmlNode::Element(el) if el.name.local == "cfRule" => {
            let priority = attr_value(el, "priority")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(u32::MAX);
            let ty = attr_value(el, "type").unwrap_or_default().to_string();
            Some((priority, ty))
        }
        _ => None,
    }
}

fn defined_names_child_sort_key(node: &XmlNode) -> (u8, String, u32, String) {
    match node {
        XmlNode::Element(el) if el.name.local == "definedName" => {
            let name = attr_value(el, "name").unwrap_or_default().to_string();
            let local_sheet_id = attr_value(el, "localSheetId")
                .and_then(|v| v.parse::<u32>().ok())
                .unwrap_or(u32::MAX);
            (0, name, local_sheet_id, String::new())
        }
        XmlNode::Element(el) => (1, el.name.local.clone(), u32::MAX, String::new()),
        XmlNode::Text(_) => (2, String::new(), u32::MAX, String::new()),
    }
}

fn row_index(el: &XmlElement) -> u32 {
    attr_value(el, "r")
        .and_then(|v| v.parse::<u32>().ok())
        .unwrap_or(u32::MAX)
}

fn cell_ref_key(el: &XmlElement) -> (u32, u32) {
    attr_value(el, "r")
        .and_then(parse_a1_reference)
        .unwrap_or((u32::MAX, u32::MAX))
}

fn attr_value<'a>(el: &'a XmlElement, local: &str) -> Option<&'a str> {
    el.attrs
        .iter()
        .find(|(k, _)| k.local == local)
        .map(|(_, v)| v.as_str())
}

fn attr_value_ns<'a>(el: &'a XmlElement, ns: Option<&str>, local: &str) -> Option<&'a str> {
    el.attrs
        .iter()
        .find(|(k, _)| k.local == local && k.ns.as_deref() == ns)
        .map(|(_, v)| v.as_str())
}

fn parse_a1_reference(reference: &str) -> Option<(u32, u32)> {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut in_row = false;

    for mut b in reference.bytes() {
        if b == b'$' {
            continue;
        }
        if !in_row {
            if b.is_ascii_alphabetic() {
                b = b.to_ascii_uppercase();
                col = col
                    .checked_mul(26)?
                    .checked_add((b - b'A' + 1) as u32)?;
                continue;
            }
            if b.is_ascii_digit() {
                in_row = true;
            } else {
                return None;
            }
        }

        if !b.is_ascii_digit() {
            return None;
        }
        row = row.checked_mul(10)?.checked_add((b - b'0') as u32)?;
    }

    if col == 0 || row == 0 {
        return None;
    }
    Some((row, col))
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

#[derive(Debug, Clone)]
struct ChildMatchKey {
    map_key: String,
    attr: &'static str,
    value: String,
}

#[derive(Debug, Clone, Copy)]
enum KeyedChildrenKind {
    Relationships,
    ContentTypes,
    SheetData,
    Row,
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

    if let Some(kind) = keyed_children_kind(&expected.name) {
        if diff_children_keyed(expected, actual, severity, &path, diffs, kind) {
            return;
        }
    }

    diff_children_indexed(expected, actual, severity, &path, diffs);
}

fn keyed_children_kind(name: &QName) -> Option<KeyedChildrenKind> {
    match (name.ns.as_deref(), name.local.as_str()) {
        (Some("http://schemas.openxmlformats.org/package/2006/relationships"), "Relationships") => {
            Some(KeyedChildrenKind::Relationships)
        }
        (Some("http://schemas.openxmlformats.org/package/2006/content-types"), "Types") => {
            Some(KeyedChildrenKind::ContentTypes)
        }
        (Some(SPREADSHEETML_NS), "sheetData") => Some(KeyedChildrenKind::SheetData),
        (Some(SPREADSHEETML_NS), "row") => Some(KeyedChildrenKind::Row),
        _ => None,
    }
}

fn diff_children_keyed(
    expected: &XmlElement,
    actual: &XmlElement,
    severity: Severity,
    path: &str,
    diffs: &mut Vec<XmlDiff>,
    kind: KeyedChildrenKind,
) -> bool {
    let expected_map = match build_keyed_children_map(&expected.children, kind) {
        Some(map) => map,
        None => return false,
    };
    let actual_map = match build_keyed_children_map(&actual.children, kind) {
        Some(map) => map,
        None => return false,
    };

    for (key, (expected_child, expected_key)) in &expected_map {
        match actual_map.get(key) {
            Some((actual_child, _actual_key)) => {
                let child_path = keyed_child_path(path, expected_child, expected_key);
                diff_node(expected_child, actual_child, severity, child_path, diffs);
            }
            None => {
                let child_path = keyed_child_path(path, expected_child, expected_key);
                diffs.push(XmlDiff {
                    severity,
                    path: child_path,
                    kind: "child_missing".to_string(),
                    expected: Some(node_summary(expected_child)),
                    actual: None,
                });
            }
        }
    }

    for (key, (actual_child, actual_key)) in &actual_map {
        if expected_map.contains_key(key) {
            continue;
        }
        let child_path = keyed_child_path(path, actual_child, actual_key);
        diffs.push(XmlDiff {
            severity,
            path: child_path,
            kind: "child_added".to_string(),
            expected: None,
            actual: Some(node_summary(actual_child)),
        });
    }

    true
}

fn build_keyed_children_map<'a>(
    children: &'a [XmlNode],
    kind: KeyedChildrenKind,
) -> Option<BTreeMap<String, (&'a XmlNode, ChildMatchKey)>> {
    let mut map: BTreeMap<String, (&'a XmlNode, ChildMatchKey)> = BTreeMap::new();
    for child in children {
        let key = child_match_key(child, kind)?;
        if map.insert(key.map_key.clone(), (child, key)).is_some() {
            // Duplicate keys - fall back to indexed comparison.
            return None;
        }
    }
    Some(map)
}

fn child_match_key(node: &XmlNode, kind: KeyedChildrenKind) -> Option<ChildMatchKey> {
    match (kind, node) {
        (KeyedChildrenKind::Relationships, XmlNode::Element(el))
            if el.name.local == "Relationship" =>
        {
            let id = attr_value(el, "Id")?;
            Some(ChildMatchKey {
                map_key: id.to_string(),
                attr: "Id",
                value: id.to_string(),
            })
        }
        (KeyedChildrenKind::ContentTypes, XmlNode::Element(el)) if el.name.local == "Default" => {
            let ext = attr_value(el, "Extension")?;
            Some(ChildMatchKey {
                map_key: format!("Default:{ext}"),
                attr: "Extension",
                value: ext.to_string(),
            })
        }
        (KeyedChildrenKind::ContentTypes, XmlNode::Element(el)) if el.name.local == "Override" => {
            let part = attr_value(el, "PartName")?;
            Some(ChildMatchKey {
                map_key: format!("Override:{part}"),
                attr: "PartName",
                value: part.to_string(),
            })
        }
        (KeyedChildrenKind::SheetData, XmlNode::Element(el)) if el.name.local == "row" => {
            let r = attr_value(el, "r")?;
            Some(ChildMatchKey {
                map_key: r.to_string(),
                attr: "r",
                value: r.to_string(),
            })
        }
        (KeyedChildrenKind::Row, XmlNode::Element(el)) if el.name.local == "c" => {
            let r = attr_value(el, "r")?;
            Some(ChildMatchKey {
                map_key: r.to_string(),
                attr: "r",
                value: r.to_string(),
            })
        }
        _ => None,
    }
}

fn diff_children_indexed(
    expected: &XmlElement,
    actual: &XmlElement,
    severity: Severity,
    path: &str,
    diffs: &mut Vec<XmlDiff>,
) {
    let max_len = expected.children.len().max(actual.children.len());
    for idx in 0..max_len {
        let expected_child = expected.children.get(idx);
        let actual_child = actual.children.get(idx);
        match (expected_child, actual_child) {
            (Some(a), Some(b)) => {
                let child_path = indexed_child_path(path, a, idx + 1);
                diff_node(a, b, severity, child_path, diffs);
            }
            (Some(a), None) => diffs.push(XmlDiff {
                severity,
                path: indexed_child_path(path, a, idx + 1),
                kind: "child_missing".to_string(),
                expected: Some(node_summary(a)),
                actual: None,
            }),
            (None, Some(b)) => diffs.push(XmlDiff {
                severity,
                path: indexed_child_path(path, b, idx + 1),
                kind: "child_added".to_string(),
                expected: None,
                actual: Some(node_summary(b)),
            }),
            (None, None) => {}
        }
    }
}

fn indexed_child_path(base: &str, child: &XmlNode, index: usize) -> String {
    format!(
        "{}/{}[{}]",
        base.trim_end_matches('/'),
        child_name(child),
        index
    )
}

fn keyed_child_path(base: &str, child: &XmlNode, key: &ChildMatchKey) -> String {
    format!(
        "{}/{}[@{}=\"{}\"]",
        base.trim_end_matches('/'),
        child_name(child),
        key.attr,
        escape_path_value(&key.value)
    )
}

fn escape_path_value(value: &str) -> String {
    if !value.contains('"') {
        return value.to_string();
    }
    value.replace('"', "\\\"")
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
    if value.len() <= MAX {
        return value.to_string();
    }

    // `str` indices are in bytes; slicing at an arbitrary byte offset can panic if the
    // cut falls in the middle of a multi-byte UTF-8 sequence. Excel XML frequently
    // contains non-ASCII text (shared strings, sheet names, comments, etc.), so we
    // must truncate on a valid UTF-8 boundary.
    let mut end = 0usize;
    for (idx, _) in value.char_indices() {
        if idx > MAX {
            break;
        }
        end = idx;
    }

    format!("{}…", &value[..end])
}
