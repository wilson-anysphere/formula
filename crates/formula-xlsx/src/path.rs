pub fn rels_for_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file_name)) => format!("{dir}/_rels/{file_name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

pub fn resolve_target(source_part: &str, target: &str) -> String {
    // Relationship targets are URIs; some producers include a URI fragment (e.g. `../media/img.png#id`).
    // OPC part names do not include fragments, so strip them before resolving.
    let target = target.split('#').next().unwrap_or(target);
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
