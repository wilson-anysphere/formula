use std::collections::HashSet;

use formula_model::CfRule;

/// Normalize `cfRule/@priority` values for serialization.
///
/// Excel expects conditional formatting `priority` attributes to be positive and unique
/// (across the worksheet). When priorities are missing (`u32::MAX`), zero, or duplicated we
/// rewrite the priorities to `1..=n` in the current rule order.
///
/// When all existing priorities are already valid (>0) and unique, they are preserved verbatim.
pub(crate) fn normalize_cf_priorities(rules: &[CfRule]) -> Vec<u32> {
    if rules.is_empty() {
        return Vec::new();
    }

    let mut seen: HashSet<u32> = HashSet::with_capacity(rules.len());
    let mut all_valid_unique = true;

    for rule in rules {
        let p = rule.priority;
        // `u32::MAX` is used throughout the codebase as "unset".
        if p == 0 || p == u32::MAX || !seen.insert(p) {
            all_valid_unique = false;
            break;
        }
    }

    if all_valid_unique {
        return rules.iter().map(|r| r.priority).collect();
    }

    // Rewrite priorities to a stable 1..=n sequence in rule order.
    (1..=rules.len() as u32).collect()
}

/// Patch the `priority` attribute on a raw `<cfRule>` XML fragment.
///
/// This is a best-effort helper used when round-tripping unsupported conditional formatting rule
/// kinds, where we preserve the original rule XML as a string.
///
/// - If `priority="..."` is present and parses to `priority`, returns `raw_xml` unchanged.
/// - If `priority` is missing or differs, updates/inserts it in the start tag.
pub(crate) fn patch_cf_rule_priority(raw_xml: &str, priority: u32) -> String {
    let Some(tag_start) = raw_xml.find('<') else {
        return raw_xml.to_string();
    };

    let bytes = raw_xml.as_bytes();

    let Some(tag_end) = find_tag_end(bytes, tag_start) else {
        return raw_xml.to_string();
    };

    // First, try to locate an existing priority attribute in the start tag.
    if let Some((value_start, value_end)) = find_attr_value_range(bytes, tag_start, tag_end, b"priority") {
        let existing_str = &raw_xml[value_start..value_end];
        if existing_str.parse::<u32>().ok() == Some(priority) {
            return raw_xml.to_string();
        }

        let mut out = String::with_capacity(raw_xml.len() + 12);
        out.push_str(&raw_xml[..value_start]);
        out.push_str(&priority.to_string());
        out.push_str(&raw_xml[value_end..]);
        return out;
    }

    // Otherwise insert a new attribute.
    let insert_pos = priority_insert_pos(bytes, tag_start, tag_end);
    let mut out = String::with_capacity(raw_xml.len() + 24);
    out.push_str(&raw_xml[..insert_pos]);
    out.push_str(r#" priority=""#);
    out.push_str(&priority.to_string());
    out.push('"');
    out.push_str(&raw_xml[insert_pos..]);
    out
}

fn find_tag_end(bytes: &[u8], tag_start: usize) -> Option<usize> {
    let mut quote: Option<u8> = None;
    for (i, &b) in bytes.iter().enumerate().skip(tag_start) {
        match quote {
            Some(q) => {
                if b == q {
                    quote = None;
                }
            }
            None => match b {
                b'\'' | b'"' => quote = Some(b),
                b'>' => return Some(i),
                _ => {}
            },
        }
    }
    None
}

fn find_attr_value_range(
    bytes: &[u8],
    tag_start: usize,
    tag_end: usize,
    attr_name: &[u8],
) -> Option<(usize, usize)> {
    // Parse a subset of XML attributes in the start tag.
    let mut i = tag_start + 1;

    // Skip element name.
    while i < tag_end && !is_space(bytes[i]) && bytes[i] != b'/' {
        i += 1;
    }

    while i < tag_end {
        // Skip whitespace.
        while i < tag_end && is_space(bytes[i]) {
            i += 1;
        }
        if i >= tag_end || bytes[i] == b'/' {
            break;
        }

        // Parse attribute name.
        let name_start = i;
        while i < tag_end && !is_space(bytes[i]) && bytes[i] != b'=' && bytes[i] != b'/' {
            i += 1;
        }
        let name_end = i;

        // Skip whitespace before '='.
        while i < tag_end && is_space(bytes[i]) {
            i += 1;
        }
        if i >= tag_end || bytes[i] != b'=' {
            // Attribute without value.
            continue;
        }
        i += 1;

        // Skip whitespace after '='.
        while i < tag_end && is_space(bytes[i]) {
            i += 1;
        }
        if i >= tag_end {
            break;
        }

        // Parse attribute value.
        let (value_start, value_end) = match bytes[i] {
            b'\'' | b'"' => {
                let q = bytes[i];
                i += 1;
                let value_start = i;
                while i < tag_end && bytes[i] != q {
                    i += 1;
                }
                let value_end = i;
                if i < tag_end && bytes[i] == q {
                    i += 1;
                }
                (value_start, value_end)
            }
            _ => {
                let value_start = i;
                while i < tag_end && !is_space(bytes[i]) && bytes[i] != b'/' {
                    i += 1;
                }
                let value_end = i;
                (value_start, value_end)
            }
        };

        if bytes[name_start..name_end] == *attr_name {
            return Some((value_start, value_end));
        }
    }

    None
}

fn priority_insert_pos(bytes: &[u8], tag_start: usize, tag_end: usize) -> usize {
    // Insert before `>` or before the `/` in a self-closing tag.
    let mut i = tag_end;
    while i > tag_start && is_space(bytes[i - 1]) {
        i -= 1;
    }

    if i > tag_start && bytes[i - 1] == b'/' {
        i - 1
    } else {
        tag_end
    }
}

fn is_space(b: u8) -> bool {
    matches!(b, b' ' | b'\n' | b'\r' | b'\t')
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{CfRuleKind, CfRuleSchema};

    fn rule(priority: u32) -> CfRule {
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority,
            applies_to: Vec::new(),
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: Vec::new(),
        }
    }

    #[test]
    fn valid_unique_priorities_are_preserved() {
        let rules = vec![rule(10), rule(1), rule(5)];
        assert_eq!(normalize_cf_priorities(&rules), vec![10, 1, 5]);
    }

    #[test]
    fn duplicate_priorities_are_rewritten_to_1_to_n() {
        let rules = vec![rule(1), rule(1), rule(2)];
        assert_eq!(normalize_cf_priorities(&rules), vec![1, 2, 3]);
    }

    #[test]
    fn unset_priorities_are_rewritten_to_1_to_n() {
        let rules = vec![rule(u32::MAX), rule(2)];
        assert_eq!(normalize_cf_priorities(&rules), vec![1, 2]);
    }

    #[test]
    fn patch_cf_rule_priority_inserts_when_missing() {
        let raw = r#"<cfRule type="expression"><formula>A1&gt;0</formula></cfRule>"#;
        let patched = patch_cf_rule_priority(raw, 3);
        assert!(
            patched.starts_with(r#"<cfRule type="expression" priority="3">"#),
            "got: {patched}"
        );
    }

    #[test]
    fn patch_cf_rule_priority_rewrites_when_present() {
        let raw = r#"<cfRule type="expression" priority="99"><formula>A1&gt;0</formula></cfRule>"#;
        let patched = patch_cf_rule_priority(raw, 3);
        assert!(patched.contains(r#"priority="3""#), "got: {patched}");
        assert!(!patched.contains(r#"priority="99""#), "got: {patched}");
    }

    #[test]
    fn patch_cf_rule_priority_is_idempotent_when_numeric_value_matches() {
        // Preserve the original formatting (leading zeros) when the numeric value matches.
        let raw = r#"<cfRule type="expression" priority="001"><formula>A1&gt;0</formula></cfRule>"#;
        let patched = patch_cf_rule_priority(raw, 1);
        assert_eq!(patched, raw);
    }

    #[test]
    fn patch_cf_rule_priority_handles_self_closing_tag() {
        let raw = r#"<cfRule type="uniqueValues"/>"#;
        let patched = patch_cf_rule_priority(raw, 7);
        assert_eq!(patched, r#"<cfRule type="uniqueValues" priority="7"/>"#);
    }
}
