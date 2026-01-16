use std::collections::HashSet;

use formula_model::CfRule;

/// Ensure that every conditional formatting rule has an `id`.
///
/// Excel uses `cfRule/@id` as the linking key between a base SpreadsheetML rule
/// (`<conditionalFormatting><cfRule .../></conditionalFormatting>`) and its optional x14
/// extension (`<x14:conditionalFormattings>...<x14:cfRule .../></x14:conditionalFormattings>`).
/// When emitting x14 rules, writers must therefore guarantee that `CfRule.id` is present.
///
/// This helper **never modifies** existing ids. For rules with a missing/empty `id`, it derives a
/// deterministic GUID-like string from the caller-provided `seed` and the rule's index in the
/// provided slice. Callers should supply a `seed` that incorporates the worksheet identity (e.g.
/// stable worksheet id / sheet index) so that ids are stable and extremely unlikely to collide
/// across different sheets.
///
/// # Collision behavior
///
/// The generated ids are 128-bit and deterministic. Collisions are therefore *extremely* unlikely,
/// but still possible if:
/// - callers reuse the same `seed` for different sheets, and/or
/// - ids are truncated/rewritten externally.
///
/// To be defensive, this function checks for collisions within `rules` (case-insensitive) and
/// retries with an incrementing counter until it finds an unused id.
pub fn ensure_rule_ids(rules: &mut [CfRule], seed: u128) {
    fn normalize_ascii_uppercase(s: &str) -> String {
        // Fast path: if there are no ASCII lowercase letters, preserve the string verbatim.
        if s.as_bytes().iter().all(|b| !b.is_ascii_lowercase()) {
            return s.to_string();
        }
        let mut owned = s.to_string();
        owned.make_ascii_uppercase();
        owned
    }

    // Use a normalized (uppercase) set for collision checks.
    let mut seen: HashSet<String> = rules
        .iter()
        .filter_map(|r| r.id.as_deref())
        .filter(|s| !s.is_empty())
        .map(normalize_ascii_uppercase)
        .collect();

    for (idx, rule) in rules.iter_mut().enumerate() {
        let needs_id = rule.id.as_deref().map(|s| s.is_empty()).unwrap_or(true);
        if !needs_id {
            continue;
        }

        let mut attempt: u32 = 0;
        loop {
            let id = rule_id_for_index_with_attempt(seed, idx, attempt);
            // `rule_id_for_index_with_attempt` returns an uppercase GUID; avoid allocating another
            // uppercased copy just to insert into the already-normalized set.
            if seen.insert(id.clone()) {
                rule.id = Some(id);
                break;
            }
            attempt = attempt.saturating_add(1);
        }
    }
}

/// Deterministically generate an Excel-style GUID for a rule index.
///
/// This is a pure helper that does **not** consult other rules for collision avoidance. Most
/// callers should prefer [`ensure_rule_ids`] instead.
pub fn rule_id_for_index(seed: u128, index: usize) -> String {
    rule_id_for_index_with_attempt(seed, index, 0)
}

fn rule_id_for_index_with_attempt(seed: u128, index: usize, attempt: u32) -> String {
    let value = deterministic_u128(seed, index, attempt);
    format_excel_guid(value)
}

fn deterministic_u128(seed: u128, index: usize, attempt: u32) -> u128 {
    // SplitMix64-based expansion from (seed, index, attempt) -> 128 bits.
    //
    // This is not intended to be cryptographically secure; it is only used to generate stable,
    // GUID-shaped identifiers with excellent distribution properties.
    let seed_lo = seed as u64;
    let seed_hi = (seed >> 64) as u64;

    let mut state = seed_lo
        ^ seed_hi
        ^ (index as u64).wrapping_mul(0x9E37_79B9_7F4A_7C15)
        ^ (attempt as u64).wrapping_mul(0xD2B7_4407_B1CE_6E93);

    let hi = splitmix64(&mut state);
    let lo = splitmix64(&mut state);

    ((hi as u128) << 64) | (lo as u128)
}

fn splitmix64(state: &mut u64) -> u64 {
    *state = state.wrapping_add(0x9E37_79B9_7F4A_7C15);
    let mut z = *state;
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58_476D_1CE4_E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D0_49BB_1331_11EB);
    z ^ (z >> 31)
}

/// Format a `u128` as an Excel-style GUID: `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.
fn format_excel_guid(value: u128) -> String {
    let hex = format!("{value:032X}");
    // Safe: `hex` is ASCII-only.
    format!(
        "{{{}-{}-{}-{}-{}}}",
        &hex[0..8],
        &hex[8..12],
        &hex[12..16],
        &hex[16..20],
        &hex[20..32]
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{parse_range_a1, CfRuleKind, CfRuleSchema};

    fn rule(id: Option<&str>) -> CfRule {
        CfRule {
            schema: CfRuleSchema::X14,
            id: id.map(|s| s.to_string()),
            priority: 1,
            applies_to: vec![parse_range_a1("A1:A1").unwrap()],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: vec![],
        }
    }

    fn is_excel_guid(s: &str) -> bool {
        if s.len() != 38 {
            return false;
        }
        let bytes = s.as_bytes();
        if bytes[0] != b'{' || bytes[37] != b'}' {
            return false;
        }
        if bytes[9] != b'-' || bytes[14] != b'-' || bytes[19] != b'-' || bytes[24] != b'-' {
            return false;
        }
        for (i, &b) in bytes.iter().enumerate() {
            if matches!(i, 0 | 9 | 14 | 19 | 24 | 37) {
                continue;
            }
            let is_hex_upper = (b'0'..=b'9').contains(&b) || (b'A'..=b'F').contains(&b);
            if !is_hex_upper {
                return false;
            }
        }
        true
    }

    #[test]
    fn rule_id_for_index_is_excel_guid_format() {
        let id = rule_id_for_index(123, 0);
        assert!(is_excel_guid(&id), "expected Excel GUID format, got {id}");
    }

    #[test]
    fn ensure_rule_ids_is_deterministic_and_preserves_existing() {
        let seed = 0x0123_4567_89AB_CDEF_FEDC_BA98_7654_3210u128;

        let mut rules1 = vec![rule(None), rule(None), rule(Some("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"))];
        let mut rules2 = rules1.clone();

        ensure_rule_ids(&mut rules1, seed);
        ensure_rule_ids(&mut rules2, seed);

        assert_eq!(rules1, rules2, "same inputs must produce identical ids");

        for r in &rules1 {
            let id = r.id.as_deref().unwrap();
            assert!(is_excel_guid(id), "expected Excel GUID format, got {id}");
        }

        assert_ne!(rules1[0].id, rules1[1].id, "ids should be distinct");
        assert_eq!(
            rules1[2].id.as_deref(),
            Some("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"),
            "existing ids must be preserved verbatim"
        );

        // Re-running should not change ids.
        let snapshot = rules1.clone();
        ensure_rule_ids(&mut rules1, seed);
        assert_eq!(rules1, snapshot, "ensure_rule_ids must be idempotent");
    }

    #[test]
    fn ensure_rule_ids_avoids_collisions() {
        let seed = 42u128;
        let collision = rule_id_for_index(seed, 0);
        // Force a collision by pre-seeding another rule with the id we would generate.
        let mut rules = vec![rule(None), rule(Some(&collision))];
        ensure_rule_ids(&mut rules, seed);
        assert_ne!(rules[0].id.as_deref(), Some(collision.as_str()));
        assert!(is_excel_guid(rules[0].id.as_deref().unwrap()));
        assert!(is_excel_guid(rules[1].id.as_deref().unwrap()));
    }
}

