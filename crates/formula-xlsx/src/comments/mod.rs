use std::collections::BTreeMap;

use formula_model::Comment;

use crate::XlsxPackage;

// These modules are crate-visible so the higher-fidelity `XlsxDocument` read/write pipeline can
// reuse the existing comment XML parsers/writers without exposing them as part of the public API.
pub(crate) mod legacy;
pub(crate) mod import;
pub(crate) mod persons;
pub(crate) mod threaded;

pub use legacy::parse_vml_drawing_cells;

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CommentParts {
    pub comments: Vec<Comment>,
    pub preserved: BTreeMap<String, Vec<u8>>,
}

impl XlsxPackage {
    /// Extract all comment-related parts from the package, parsing any known XML
    /// formats (legacy notes + modern threaded comments) and preserving all
    /// comment-related parts byte-for-byte.
    pub fn comment_parts(&self) -> CommentParts {
        extract_comment_parts(self)
    }

    /// Apply an updated set of comment parts back onto the package.
    ///
    /// This writes any regenerated XML (`comments*.xml`, `threadedComments*.xml`)
    /// while keeping all other comment-related parts preserved verbatim.
    pub fn write_comment_parts(&mut self, parts: &CommentParts) {
        for (path, bytes) in render_comment_parts(parts) {
            self.set_part(path, bytes);
        }
    }
}

pub fn extract_comment_parts(pkg: &XlsxPackage) -> CommentParts {
    let mut parts = CommentParts::default();

    let mut persons = BTreeMap::<String, String>::new();

    for (path, bytes) in pkg.parts() {
        if !is_comment_related_part(path) {
            continue;
        }

        parts.preserved.insert(path.to_string(), bytes.to_vec());

        if is_persons_xml(path) {
            if let Ok(parsed) = persons::parse_persons_xml(bytes) {
                persons.extend(parsed);
            }
        }
    }

    for (path, bytes) in &parts.preserved {
        if is_legacy_comments_xml(path) {
            if let Ok(mut comments) = legacy::parse_comments_xml(bytes) {
                parts.comments.append(&mut comments);
            }
        } else if is_threaded_comments_xml(path) {
            if let Ok(mut comments) = threaded::parse_threaded_comments_xml(bytes, &persons) {
                parts.comments.append(&mut comments);
            }
        }
    }

    parts
}

pub fn apply_comment_parts(pkg: &mut XlsxPackage, parts: &CommentParts) {
    for (path, bytes) in &parts.preserved {
        pkg.set_part(path.clone(), bytes.clone());
    }
}

pub fn render_comment_parts(parts: &CommentParts) -> BTreeMap<String, Vec<u8>> {
    let mut out = parts.preserved.clone();

    let legacy_target = parts
        .preserved
        .keys()
        .find(|path| is_legacy_comments_xml(path))
        .cloned()
        .unwrap_or_else(|| "xl/comments1.xml".to_string());

    let legacy_comments = parts
        .comments
        .iter()
        .filter(|comment| comment.kind == formula_model::CommentKind::Note)
        .cloned()
        .collect::<Vec<_>>();
    if !legacy_comments.is_empty() {
        out.insert(legacy_target, legacy::write_comments_xml(&legacy_comments));
    }

    let threaded_target = parts
        .preserved
        .keys()
        .find(|path| is_threaded_comments_xml(path))
        .cloned()
        .unwrap_or_else(|| "xl/threadedComments/threadedComments1.xml".to_string());

    let threaded_comments = parts
        .comments
        .iter()
        .filter(|comment| comment.kind == formula_model::CommentKind::Threaded)
        .cloned()
        .collect::<Vec<_>>();
    if !threaded_comments.is_empty() {
        out.insert(
            threaded_target,
            threaded::write_threaded_comments_xml(&threaded_comments),
        );
    }

    out
}

fn normalize_part_path_for_match(path: &str) -> String {
    // XLSX part names are case-sensitive in the OPC spec, but in practice we want to be tolerant to
    // common producer mistakes (Windows separators, leading slashes, ASCII casing differences).
    let path = path.trim_start_matches(|c| c == '/' || c == '\\');
    if path.contains('\\') {
        path.replace('\\', "/")
    } else {
        path.to_string()
    }
}

fn is_comment_related_part(path: &str) -> bool {
    let path = normalize_part_path_for_match(path);
    crate::ascii::starts_with_ignore_case(&path, "xl/comments")
        || crate::ascii::starts_with_ignore_case(&path, "xl/threadedComments/")
        || crate::ascii::starts_with_ignore_case(&path, "xl/persons/")
        || crate::ascii::starts_with_ignore_case(&path, "xl/drawings/vmlDrawing")
        || crate::ascii::starts_with_ignore_case(&path, "xl/drawings/commentsDrawing")
        || crate::ascii::contains_ignore_case(&path, "commentsExt")
}

fn is_legacy_comments_xml(path: &str) -> bool {
    let path = normalize_part_path_for_match(path);
    if !crate::ascii::starts_with_ignore_case(&path, "xl/comments")
        || !crate::ascii::ends_with_ignore_case(&path, ".xml")
    {
        return false;
    }
    if crate::ascii::contains_ignore_case(&path, "threaded") {
        return false;
    }
    // Exclude other comment-related parts like `xl/commentsExt*.xml`.
    let suffix = &path["xl/comments".len()..];
    if suffix.starts_with('.') {
        return true;
    }
    if suffix
        .as_bytes()
        .first()
        .is_some_and(|b| b.is_ascii_digit())
    {
        return true;
    }
    if suffix.as_bytes().first() == Some(&b'%') {
        if let Some(decoded) = decode_percent_byte(suffix) {
            return decoded == b'.' || decoded.is_ascii_digit();
        }
    }
    false
}

fn is_threaded_comments_xml(path: &str) -> bool {
    let path = normalize_part_path_for_match(path);
    crate::ascii::starts_with_ignore_case(&path, "xl/threadedcomments/")
        && crate::ascii::ends_with_ignore_case(&path, ".xml")
}

fn is_persons_xml(path: &str) -> bool {
    let path = normalize_part_path_for_match(path);
    crate::ascii::starts_with_ignore_case(&path, "xl/persons/")
        && crate::ascii::ends_with_ignore_case(&path, ".xml")
}

fn decode_percent_byte(s: &str) -> Option<u8> {
    // Percent-encoded URIs are common in relationships, but some producers may also percent-encode
    // the underlying ZIP entry names. We only need the first decoded byte for comment part
    // classification.
    let bytes = s.as_bytes();
    if bytes.len() < 3 || bytes[0] != b'%' {
        return None;
    }
    let hi = hex_value(bytes[1])?;
    let lo = hex_value(bytes[2])?;
    Some((hi << 4) | lo)
}

fn hex_value(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}
