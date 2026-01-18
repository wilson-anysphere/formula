use std::borrow::Cow;

pub fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

pub fn resolve_target(source_part: &str, target: &str) -> String {
    resolve_target_inner(source_part, target)
}

/// Resolve a relationship `Target` URI to a list of normalized OPC part name candidates.
///
/// Some producers percent-encode relationship targets (e.g. `sheet%201.xml`) while storing ZIP
/// entries with unescaped names (`sheet 1.xml`), and vice versa. Callers that need to open ZIP
/// parts should try each candidate in order.
pub fn resolve_target_candidates(source_part: &str, target: &str) -> Vec<String> {
    let raw = resolve_target_inner(source_part, target);

    // Fast path: if the target path has no percent escapes *and* no characters we can reasonably
    // percent-encode (today: spaces), there is no alternate candidate.
    let source_part_seps = normalize_separators(source_part);
    let target_seps = normalize_separators(target);
    let stripped = strip_uri_suffixes(target_seps.as_ref());
    if !stripped.contains('%') && !stripped.contains(' ') {
        return vec![raw];
    }

    let mut candidates = vec![raw.clone()];

    // If the relationship target is percent-encoded but the stored ZIP entry name is unescaped,
    // include the best-effort percent-decoded candidate.
    if let Cow::Owned(decoded) = percent_decode_best_effort(stripped) {
        let decoded_resolved =
            resolve_target_from_stripped(source_part_seps.as_ref(), decoded.as_str());
        if decoded_resolved != raw {
            candidates.push(decoded_resolved);
        }
    }

    // If the relationship target contains unescaped spaces (producer bug) but the stored ZIP entry
    // name is percent-encoded, include an encoded candidate.
    if let Cow::Owned(encoded) = percent_encode_best_effort(stripped) {
        let encoded_resolved =
            resolve_target_from_stripped(source_part_seps.as_ref(), encoded.as_str());
        if !candidates.iter().any(|c| c == &encoded_resolved) {
            candidates.push(encoded_resolved);
        }
    }

    candidates
}

fn strip_uri_suffixes(target: &str) -> &str {
    let target = target.trim();
    let target = target.split_once('#').map(|(t, _)| t).unwrap_or(target);
    target.split_once('?').map(|(t, _)| t).unwrap_or(target)
}

fn resolve_target_inner(source_part: &str, target: &str) -> String {
    let source_part = normalize_separators(source_part);
    let target = normalize_separators(target);

    // Relationship targets are URIs. For in-package parts, the `#fragment` and `?query` portions
    // are not part of the OPC part name and must be ignored when mapping to ZIP entry names.
    //
    // Some producers (and some Excel-generated parts) include fragments on image relationships,
    // e.g. `../media/image1.png#something`. Excel itself treats this as a reference to the same
    // underlying image part.
    let target_path = strip_uri_suffixes(target.as_ref());
    resolve_target_from_stripped(source_part.as_ref(), target_path)
}

fn resolve_target_from_stripped(source_part: &str, target: &str) -> String {
    if target.is_empty() {
        // A target of just `#fragment` refers to the source part itself.
        return normalize(source_part);
    }
    if let Some(target) = target.strip_prefix('/') {
        return normalize(target);
    }

    let base_dir = source_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    normalize(&format!("{base_dir}/{target}"))
}

fn normalize_separators(path: &str) -> Cow<'_, str> {
    // Be resilient to invalid/unescaped Windows-style path separators.
    if path.contains('\\') {
        Cow::Owned(path.replace('\\', "/"))
    } else {
        Cow::Borrowed(path)
    }
}

fn percent_decode_best_effort(input: &str) -> Cow<'_, str> {
    let bytes = input.as_bytes();
    let Some(first_pct) = bytes.iter().position(|b| *b == b'%') else {
        return Cow::Borrowed(input);
    };

    // Only allocate if there is at least one valid percent escape to decode.
    let mut has_valid_escape = false;
    let mut i = first_pct;
    while i + 2 < bytes.len() {
        if bytes[i] == b'%' && is_hex_digit(bytes[i + 1]) && is_hex_digit(bytes[i + 2]) {
            has_valid_escape = true;
            break;
        }
        i += 1;
    }
    if !has_valid_escape {
        return Cow::Borrowed(input);
    }

    let mut out = Vec::new();
    if out.try_reserve_exact(bytes.len()).is_err() {
        return Cow::Borrowed(input);
    }
    out.extend_from_slice(&bytes[..first_pct]);
    let mut i = first_pct;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() && is_hex_digit(bytes[i + 1]) && is_hex_digit(bytes[i + 2]) {
            let hi = hex_value(bytes[i + 1]);
            let lo = hex_value(bytes[i + 2]);
            out.push((hi << 4) | lo);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }

    match String::from_utf8(out) {
        Ok(s) => Cow::Owned(s),
        Err(_) => Cow::Borrowed(input),
    }
}

fn percent_encode_best_effort(input: &str) -> Cow<'_, str> {
    // Relationship targets are URIs, so spaces should be percent-encoded as `%20`. Some producers
    // emit literal spaces; tolerate those by generating an alternate candidate.
    if !input.as_bytes().contains(&b' ') {
        return Cow::Borrowed(input);
    }

    let mut out = String::new();
    if out.try_reserve(input.len().saturating_add(2)).is_err() {
        return Cow::Borrowed(input);
    }
    for ch in input.chars() {
        if ch == ' ' {
            out.push_str("%20");
        } else {
            out.push(ch);
        }
    }
    Cow::Owned(out)
}

fn is_hex_digit(b: u8) -> bool {
    matches!(b, b'0'..=b'9' | b'a'..=b'f' | b'A'..=b'F')
}

fn hex_value(b: u8) -> u8 {
    match b {
        b'0'..=b'9' => b - b'0',
        b'a'..=b'f' => 10 + (b - b'a'),
        b'A'..=b'F' => 10 + (b - b'A'),
        _ => 0,
    }
}

fn normalize(path: &str) -> String {
    let mut out: Vec<&str> = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            other => out.push(other),
        }
    }
    out.join("/")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rels_for_part_in_root() {
        assert_eq!(rels_for_part("workbook.xml"), "_rels/workbook.xml.rels");
    }

    #[test]
    fn rels_for_part_in_subdir() {
        assert_eq!(rels_for_part("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
    }

    #[test]
    fn resolve_target_relative_to_source_dir() {
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "../media/image1.png"),
            "xl/media/image1.png"
        );
    }

    #[test]
    fn resolve_target_strips_fragments() {
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets/sheet1.xml#rId1"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn resolve_target_strips_query_strings() {
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets/sheet1.xml?foo=bar"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets/sheet1.xml?foo=bar#rId1"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn resolve_target_normalizes_backslashes() {
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "..\\media\\image1.png"),
            "xl/media/image1.png"
        );
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets\\sheet1.xml#rId1"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn resolve_target_hash_only_refs_source_part() {
        assert_eq!(resolve_target("xl/workbook.xml", "#rId1"), "xl/workbook.xml");
    }

    #[test]
    fn resolve_target_absolute_paths_are_normalized() {
        assert_eq!(
            resolve_target("xl/workbook.xml", "/xl/../docProps/core.xml"),
            "docProps/core.xml"
        );
    }

    #[test]
    fn resolve_target_handles_dot_segments() {
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "./../worksheets/./sheet2.xml"),
            "xl/worksheets/sheet2.xml"
        );
    }

    #[test]
    fn resolve_target_candidates_returns_percent_decoded_variant() {
        let targets = resolve_target_candidates("xl/workbook.xml", "worksheets/sheet%201.xml");
        assert_eq!(
            targets,
            vec![
                "xl/worksheets/sheet%201.xml".to_string(),
                "xl/worksheets/sheet 1.xml".to_string()
            ]
        );
    }

    #[test]
    fn resolve_target_candidates_returns_percent_encoded_variant() {
        let targets = resolve_target_candidates("xl/workbook.xml", "worksheets/sheet 1.xml");
        assert_eq!(
            targets,
            vec![
                "xl/worksheets/sheet 1.xml".to_string(),
                "xl/worksheets/sheet%201.xml".to_string()
            ]
        );
    }
}
