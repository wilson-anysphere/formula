//! ASCII-only case-insensitive helpers for protocol-ish strings.
//!
//! These helpers intentionally operate on raw bytes and use ASCII-only comparisons:
//! - ASCII letters compare case-insensitively
//! - non-ASCII bytes must match exactly
//!
//! This is appropriate for OOXML part names, relationship targets, and similar internal paths.

use std::borrow::Cow;

#[inline]
pub fn starts_with_ignore_case(s: &str, prefix: &str) -> bool {
    let s = s.as_bytes();
    let prefix = prefix.as_bytes();
    s.get(..prefix.len())
        .is_some_and(|p| p.eq_ignore_ascii_case(prefix))
}

#[inline]
pub fn ends_with_ignore_case(s: &str, suffix: &str) -> bool {
    let s = s.as_bytes();
    let suffix = suffix.as_bytes();
    if s.len() < suffix.len() {
        return false;
    }
    let start = s.len() - suffix.len();
    s.get(start..)
        .is_some_and(|tail| tail.eq_ignore_ascii_case(suffix))
}

#[inline]
pub fn strip_prefix_ignore_case<'a>(s: &'a str, prefix: &str) -> Option<&'a str> {
    starts_with_ignore_case(s, prefix).then(|| s.get(prefix.len()..)).flatten()
}

#[inline]
pub fn strip_suffix_ignore_case<'a>(s: &'a str, suffix: &str) -> Option<&'a str> {
    if !ends_with_ignore_case(s, suffix) {
        return None;
    }
    let end = s.len().checked_sub(suffix.len())?;
    s.get(..end)
}

#[inline]
pub fn contains_ignore_case(haystack: &str, needle: &str) -> bool {
    let haystack = haystack.as_bytes();
    let needle = needle.as_bytes();
    if needle.is_empty() {
        return true;
    }
    if haystack.len() < needle.len() {
        return false;
    }
    haystack
        .windows(needle.len())
        .any(|w| w.eq_ignore_ascii_case(needle))
}

#[inline]
pub fn rfind_ignore_case(haystack: &str, needle: &str) -> Option<usize> {
    let haystack = haystack.as_bytes();
    let needle = needle.as_bytes();
    if needle.is_empty() {
        return Some(haystack.len());
    }
    if haystack.len() < needle.len() {
        return None;
    }
    haystack
        .windows(needle.len())
        .rposition(|w| w.eq_ignore_ascii_case(needle))
}

pub fn normalize_extension_ascii_lowercase(ext: &str) -> Cow<'_, str> {
    if ext.as_bytes().iter().all(|b| !b.is_ascii_uppercase()) {
        return Cow::Borrowed(ext);
    }
    Cow::Owned(ext.to_ascii_lowercase())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn starts_and_ends_ignore_ascii_case() {
        assert!(starts_with_ignore_case("xl/workbook.xml", "XL/"));
        assert!(ends_with_ignore_case("xl/workbook.XML", ".xml"));
        assert!(!starts_with_ignore_case("xl/workbook.xml", "xl/workbook.xmlx"));
        assert!(!ends_with_ignore_case("xl/workbook.xml", ".xmlx"));
    }

    #[test]
    fn strip_prefix_suffix_ignore_ascii_case() {
        assert_eq!(
            strip_prefix_ignore_case("xl/workbook.xml", "XL/"),
            Some("workbook.xml")
        );
        assert_eq!(
            strip_suffix_ignore_case("xl/workbook.XML", ".xml"),
            Some("xl/workbook")
        );
        assert_eq!(strip_prefix_ignore_case("xl/workbook.xml", "nope"), None);
        assert_eq!(strip_suffix_ignore_case("xl/workbook.xml", "nope"), None);
    }

    #[test]
    fn contains_and_rfind_ignore_ascii_case() {
        assert!(contains_ignore_case("xl/_rels/workbook.xml.rels", "/_rels/"));
        assert_eq!(
            rfind_ignore_case("xl/_rels/workbook.xml.rels", "/_RELS/"),
            Some(2)
        );
        assert_eq!(rfind_ignore_case("abc", ""), Some(3));
    }

    #[test]
    fn does_not_casefold_unicode() {
        // Non-ASCII bytes must match exactly.
        assert!(!starts_with_ignore_case("école", "É"));
        assert!(!contains_ignore_case("école", "É"));
    }
}

