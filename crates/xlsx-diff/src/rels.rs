use std::collections::{BTreeMap, BTreeSet};
use std::fmt;

use anyhow::{Context, Result};
use roxmltree::Document;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) enum RelationshipTargetMode {
    Internal,
    External,
}

impl RelationshipTargetMode {
    pub(crate) fn from_attribute(value: Option<&str>) -> Self {
        match value {
            Some(mode) if mode.eq_ignore_ascii_case("External") => Self::External,
            _ => Self::Internal,
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Internal => "Internal",
            Self::External => "External",
        }
    }
}

impl fmt::Display for RelationshipTargetMode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct RelationshipSemanticKey {
    pub(crate) ty: String,
    /// Relationship target normalized to a stable semantic form:
    /// - internal targets are resolved relative to the `.rels` part and normalized as an OPC part name
    /// - external targets are normalized to forward slashes but otherwise left untouched
    pub(crate) resolved_target: String,
    pub(crate) target_mode: RelationshipTargetMode,
}

impl RelationshipSemanticKey {
    pub(crate) fn to_diff_path(&self) -> String {
        format!(
            "/Relationship[@Type=\"{}\"][@ResolvedTarget=\"{}\"][@TargetMode=\"{}\"]",
            escape_path_value(&self.ty),
            escape_path_value(&self.resolved_target),
            self.target_mode
        )
    }
}

#[derive(Debug, Clone)]
pub(crate) struct RelationshipIdChange {
    pub(crate) key: RelationshipSemanticKey,
    pub(crate) expected_id: String,
    pub(crate) actual_id: String,
}

pub(crate) fn relationship_semantic_id_map(
    rels_part: &str,
    bytes: &[u8],
) -> Result<BTreeMap<RelationshipSemanticKey, String>> {
    let text =
        std::str::from_utf8(bytes).with_context(|| format!("part {rels_part} is not valid UTF-8"))?;
    let doc = Document::parse(text).with_context(|| format!("parse xml for {rels_part}"))?;

    let mut map: BTreeMap<RelationshipSemanticKey, String> = BTreeMap::new();
    let mut ambiguous: BTreeSet<RelationshipSemanticKey> = BTreeSet::new();

    for node in doc
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "Relationship")
    {
        let Some(id) = node.attribute("Id") else {
            continue;
        };
        let ty = node.attribute("Type").unwrap_or_default();
        let target = node.attribute("Target").unwrap_or_default();
        let mode = RelationshipTargetMode::from_attribute(node.attribute("TargetMode"));
        let resolved_target = match mode {
            RelationshipTargetMode::External => target.replace('\\', "/"),
            RelationshipTargetMode::Internal => super::resolve_relationship_target(rels_part, target),
        };

        let key = RelationshipSemanticKey {
            ty: ty.to_string(),
            resolved_target,
            target_mode: mode,
        };

        if ambiguous.contains(&key) {
            continue;
        }

        if map.insert(key.clone(), id.to_string()).is_some() {
            // Duplicate semantic keys - treat as ambiguous and drop from the map. This avoids
            // misclassifying add/remove churn as a pure Id renumbering.
            map.remove(&key);
            ambiguous.insert(key);
        }
    }

    Ok(map)
}

pub(crate) fn detect_relationship_id_changes(
    rels_part: &str,
    expected: &[u8],
    actual: &[u8],
) -> Result<Vec<RelationshipIdChange>> {
    let expected_map = relationship_semantic_id_map(rels_part, expected)?;
    let actual_map = relationship_semantic_id_map(rels_part, actual)?;

    let mut changes = Vec::new();
    for (key, expected_id) in &expected_map {
        let Some(actual_id) = actual_map.get(key) else {
            continue;
        };
        if expected_id != actual_id {
            changes.push(RelationshipIdChange {
                key: key.clone(),
                expected_id: expected_id.clone(),
                actual_id: actual_id.clone(),
            });
        }
    }

    Ok(changes)
}

fn escape_path_value(value: &str) -> String {
    value.replace('"', "\\\"")
}

