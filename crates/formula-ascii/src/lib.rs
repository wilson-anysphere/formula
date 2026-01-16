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
    s[s.len() - suffix.len()..].eq_ignore_ascii_case(suffix)
}

#[inline]
pub fn strip_prefix_ignore_case<'a>(s: &'a str, prefix: &str) -> Option<&'a str> {
    starts_with_ignore_case(s, prefix).then(|| &s[prefix.len()..])
}

#[inline]
pub fn strip_suffix_ignore_case<'a>(s: &'a str, suffix: &str) -> Option<&'a str> {
    ends_with_ignore_case(s, suffix).then(|| &s[..s.len() - suffix.len()])
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
    for start in 0..=haystack.len() - needle.len() {
        if haystack[start..start + needle.len()].eq_ignore_ascii_case(needle) {
            return true;
        }
    }
    false
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
    for start in (0..=haystack.len() - needle.len()).rev() {
        if haystack[start..start + needle.len()].eq_ignore_ascii_case(needle) {
            return Some(start);
        }
    }
    None
}

pub fn normalize_extension_ascii_lowercase(ext: &str) -> Cow<'_, str> {
    if ext.as_bytes().iter().all(|b| !b.is_ascii_uppercase()) {
        return Cow::Borrowed(ext);
    }
    Cow::Owned(ext.to_ascii_lowercase())
}

