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

fn is_comment_related_part(path: &str) -> bool {
    path.starts_with("xl/comments")
        || path.starts_with("xl/threadedComments/")
        || path.starts_with("xl/persons/")
        || path.starts_with("xl/drawings/vmlDrawing")
        || path.starts_with("xl/drawings/commentsDrawing")
        || path.contains("commentsExt")
}

fn is_legacy_comments_xml(path: &str) -> bool {
    path.starts_with("xl/comments") && path.ends_with(".xml") && !path.contains("threaded")
}

fn is_threaded_comments_xml(path: &str) -> bool {
    path.contains("threadedComments") && path.ends_with(".xml")
}

fn is_persons_xml(path: &str) -> bool {
    path.starts_with("xl/persons/") && path.ends_with(".xml")
}
