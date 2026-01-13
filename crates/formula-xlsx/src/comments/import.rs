use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};

use formula_model::Worksheet;

/// Worksheet relationship type URI for legacy note comments (`xl/comments*.xml`).
const REL_TYPE_LEGACY_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

/// Worksheet relationship type URI for threaded comments (`xl/threadedComments/*.xml`).
const REL_TYPE_THREADED_COMMENTS: &str =
    "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";

fn is_threaded_comment_rel_type(type_uri: &str) -> bool {
    // Excel has emitted a few variants over time; accept the canonical URI and tolerate
    // other future variants that contain "threadedComment".
    type_uri == REL_TYPE_THREADED_COMMENTS || type_uri.contains("threadedComment")
}

fn is_person_rel_type(type_uri: &str) -> bool {
    // Common form: `http://schemas.microsoft.com/office/2017/10/relationships/person`.
    // Be tolerant to future versions by matching the path suffix.
    type_uri.ends_with("/person") || type_uri.contains("relationships/person")
}

/// Collect a threaded-comment personId -> displayName mapping.
///
/// This supports both:
/// - Explicit `xl/persons/*.xml` parts (passed in via `person_part_names`)
/// - Workbook-level relationships of type `.../relationships/person`
///
/// Best-effort: parse failures are ignored.
pub(crate) fn collect_persons<'a, F, I>(
    workbook_part: &str,
    workbook_rels_xml: &[u8],
    person_part_names: I,
    mut get_part: F,
) -> BTreeMap<String, String>
where
    F: FnMut(&str) -> Option<Cow<'a, [u8]>>,
    I: IntoIterator<Item = String>,
{
    let mut targets: BTreeSet<String> = BTreeSet::new();

    // 1) Any explicit `xl/persons/*.xml` parts.
    targets.extend(person_part_names);

    // 2) Any persons parts referenced via workbook relationships.
    if let Ok(rels) = crate::openxml::parse_relationships(workbook_rels_xml) {
        for rel in rels {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            if !is_person_rel_type(&rel.type_uri) {
                continue;
            }

            let target = crate::path::resolve_target(workbook_part, &rel.target);
            if !target.is_empty() {
                targets.insert(target);
            }
        }
    }

    let mut persons = BTreeMap::<String, String>::new();
    for target in targets {
        let Some(bytes) = get_part(&target) else {
            continue;
        };
        if let Ok(parsed) = super::persons::parse_persons_xml(bytes.as_ref()) {
            persons.extend(parsed);
        }
    }

    persons
}

/// Import legacy + threaded comments for a single worksheet into the in-memory model.
///
/// Best-effort: missing relationships/parts or parse errors are ignored.
pub(crate) fn import_sheet_comments<'a, F>(
    worksheet: &mut Worksheet,
    worksheet_part: &str,
    worksheet_rels_xml: Option<&[u8]>,
    persons: &BTreeMap<String, String>,
    mut get_part: F,
) where
    F: FnMut(&str) -> Option<Cow<'a, [u8]>>,
{
    let Some(rels_xml) = worksheet_rels_xml else {
        return;
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return,
    };

    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        let is_legacy = rel.type_uri == REL_TYPE_LEGACY_COMMENTS;
        let is_threaded = is_threaded_comment_rel_type(&rel.type_uri);
        if !is_legacy && !is_threaded {
            continue;
        }

        let target = crate::path::resolve_target(worksheet_part, &rel.target);
        if target.is_empty() {
            continue;
        }
        let Some(bytes) = get_part(&target) else {
            continue;
        };

        let parsed = if is_legacy {
            super::legacy::parse_comments_xml(bytes.as_ref()).ok()
        } else {
            super::threaded::parse_threaded_comments_xml(bytes.as_ref(), persons).ok()
        };
        let Some(comments) = parsed else {
            continue;
        };

        for comment in comments {
            // Use the worksheet helper so merged-cell anchoring semantics apply.
            let _ = worksheet.add_comment(comment.cell_ref, comment);
        }
    }
}

