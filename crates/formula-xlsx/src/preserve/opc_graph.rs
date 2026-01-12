use std::collections::{BTreeMap, HashSet, VecDeque};

use crate::openxml::parse_relationships;
use crate::path::{rels_for_part, resolve_target};
use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

/// Collect a transitive closure of OPC parts reachable from `root_parts` via `.rels` files.
///
/// The closure includes:
/// - Each root part itself (when present in the package).
/// - The part's corresponding `.rels` (when present).
/// - All internal relationship targets that exist in the package, recursively.
///
/// This helper is used for best-effort preservation, so it intentionally ignores:
/// - Missing `.rels` parts
/// - Malformed `.rels` XML
/// - `TargetMode="External"` relationships
pub(crate) fn collect_transitive_related_parts(
    pkg: &XlsxPackage,
    root_parts: impl IntoIterator<Item = String>,
) -> Result<BTreeMap<String, Vec<u8>>, ChartExtractionError> {
    let mut out: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut visited: HashSet<String> = HashSet::new();
    let mut queue: VecDeque<String> = root_parts.into_iter().collect();

    while let Some(part_name) = queue.pop_front() {
        if !visited.insert(part_name.clone()) {
            continue;
        }

        let Some(part_bytes) = pkg.part(&part_name) else {
            continue;
        };
        out.insert(part_name.clone(), part_bytes.to_vec());

        let rels_part_name = rels_for_part(&part_name);
        let Some(rels_bytes) = pkg.part(&rels_part_name) else {
            continue;
        };
        out.insert(rels_part_name.clone(), rels_bytes.to_vec());

        let relationships = match parse_relationships(rels_bytes) {
            Ok(rels) => rels,
            Err(_) => continue,
        };

        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }

            let target = strip_fragment(&rel.target);
            if target.is_empty() {
                continue;
            }
            // Be resilient to invalid/unescaped Windows-style path separators.
            let target: std::borrow::Cow<'_, str> = if target.contains('\\') {
                std::borrow::Cow::Owned(target.replace('\\', "/"))
            } else {
                std::borrow::Cow::Borrowed(target)
            };
            let target = target.as_ref();
            let target = target.strip_prefix("./").unwrap_or(target);
            let target_part = resolve_target(&part_name, target);
            if pkg.part(&target_part).is_some() {
                queue.push_back(target_part);
                continue;
            }

            // Some rich-data producers emit relationship targets like `media/image1.png` (relative to
            // `xl/`) instead of the more common `../media/image1.png` (relative to `xl/richData/`).
            // Be defensive and try a best-effort fallback for RichData parts.
            if part_name.starts_with("xl/richData/") {
                if target.starts_with("media/") {
                    let alt = format!("xl/{target}");
                    if pkg.part(&alt).is_some() {
                        queue.push_back(alt);
                    }
                } else if target.starts_with("xl/") && pkg.part(target).is_some() {
                    queue.push_back(target.to_string());
                }
            }
        }
    }

    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::collections::BTreeSet;
    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    fn keys(map: BTreeMap<String, Vec<u8>>) -> BTreeSet<String> {
        map.into_keys().collect()
    }

    #[test]
    fn missing_rels_part_is_ignored() {
        let pkg = build_package(&[("xl/worksheets/sheet1.xml", br#"<worksheet/>"#)]);

        let parts = collect_transitive_related_parts(&pkg, ["xl/worksheets/sheet1.xml".to_string()])
            .expect("traverse");

        assert_eq!(
            keys(parts),
            BTreeSet::from(["xl/worksheets/sheet1.xml".to_string()])
        );
    }

    #[test]
    fn malformed_rels_xml_is_ignored() {
        // Rel targets a valid part, but the rels XML is malformed, so traversal should not follow it.
        let pkg = build_package(&[
            ("xl/drawings/drawing1.xml", br#"<wsDr/>"#),
            ("xl/drawings/_rels/drawing1.xml.rels", br#"<Relationships><Relationship"#),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let parts = collect_transitive_related_parts(&pkg, ["xl/drawings/drawing1.xml".to_string()])
            .expect("traverse");

        assert_eq!(
            keys(parts),
            BTreeSet::from([
                "xl/drawings/drawing1.xml".to_string(),
                "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            ])
        );
    }

    #[test]
    fn external_relationship_is_ignored() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/worksheets/sheet1.xml", br#"<worksheet/>"#),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels),
        ]);

        let parts = collect_transitive_related_parts(&pkg, ["xl/worksheets/sheet1.xml".to_string()])
            .expect("traverse");

        assert_eq!(
            keys(parts),
            BTreeSet::from([
                "xl/worksheets/sheet1.xml".to_string(),
                "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
            ])
        );
    }

    #[test]
    fn rich_data_media_relative_targets_with_backslashes_are_handled() {
        // Some producers emit `Target="media\\image1.png"` (relative to `xl/`) instead of the more
        // canonical `../media/image1.png` form. Ensure the preservation traversal still follows the
        // relationship when the target uses backslashes.
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media\image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/richData/richValueRel.xml", br#"<rvRel/>"#),
            ("xl/richData/_rels/richValueRel.xml.rels", rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let parts =
            collect_transitive_related_parts(&pkg, ["xl/richData/richValueRel.xml".to_string()])
                .expect("traverse");

        assert_eq!(
            keys(parts),
            BTreeSet::from([
                "xl/media/image1.png".to_string(),
                "xl/richData/richValueRel.xml".to_string(),
                "xl/richData/_rels/richValueRel.xml.rels".to_string(),
            ])
        );
    }

    #[test]
    fn missing_target_is_ignored() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/drawings/drawing1.xml", br#"<wsDr/>"#),
            ("xl/drawings/_rels/drawing1.xml.rels", rels),
        ]);

        let parts = collect_transitive_related_parts(&pkg, ["xl/drawings/drawing1.xml".to_string()])
            .expect("traverse");

        assert_eq!(
            keys(parts),
            BTreeSet::from([
                "xl/drawings/drawing1.xml".to_string(),
                "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            ])
        );
    }

    #[test]
    fn cycle_is_handled() {
        let a_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="b.xml"/>
</Relationships>"#;

        let b_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="a.xml"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/a.xml", br#"<a/>"#),
            ("xl/_rels/a.xml.rels", a_rels),
            ("xl/b.xml", br#"<b/>"#),
            ("xl/_rels/b.xml.rels", b_rels),
        ]);

        let parts = collect_transitive_related_parts(&pkg, ["xl/a.xml".to_string()]).expect("traverse");
        assert_eq!(
            keys(parts),
            BTreeSet::from([
                "xl/a.xml".to_string(),
                "xl/_rels/a.xml.rels".to_string(),
                "xl/b.xml".to_string(),
                "xl/_rels/b.xml.rels".to_string(),
            ])
        );
    }
}
