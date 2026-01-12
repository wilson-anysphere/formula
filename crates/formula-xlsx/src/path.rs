use std::borrow::Cow;

pub fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

pub fn resolve_target(source_part: &str, target: &str) -> String {
    // Be resilient to invalid/unescaped Windows-style path separators.
    let source_part: Cow<'_, str> = if source_part.contains('\\') {
        Cow::Owned(source_part.replace('\\', "/"))
    } else {
        Cow::Borrowed(source_part)
    };

    // Be resilient to invalid/unescaped Windows-style path separators.
    let target: Cow<'_, str> = if target.contains('\\') {
        Cow::Owned(target.replace('\\', "/"))
    } else {
        Cow::Borrowed(target)
    };

    // Relationship targets are URIs. For in-package parts, the `#fragment` and `?query` portions
    // are not part of the OPC part name and must be ignored when mapping to ZIP entry names.
    //
    // Some producers (and some Excel-generated parts) include fragments on image relationships,
    // e.g. `../media/image1.png#something`. Excel itself treats this as a reference to the same
    // underlying image part.
    let target = strip_uri_suffixes(target.as_ref());
    if target.is_empty() {
        // A target of just `#fragment` refers to the source part itself.
        return normalize(source_part.as_ref());
    }
    if let Some(target) = target.strip_prefix('/') {
        return normalize(target);
    }

    let base_dir = source_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    normalize(&format!("{base_dir}/{target}"))
}

fn strip_uri_suffixes(target: &str) -> &str {
    let target = target.trim();
    let target = target.split_once('#').map(|(t, _)| t).unwrap_or(target);
    target.split_once('?').map(|(t, _)| t).unwrap_or(target)
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
}
