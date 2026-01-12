use std::collections::BTreeMap;

use crate::openxml;
use crate::path::resolve_target;
use crate::{XlsxError, XlsxPackage};

const RICH_VALUE_REL_RELS_PREFIX: &str = "xl/richData/_rels/richValueRel";
const RICH_VALUE_REL_RELS_SUFFIX: &str = ".xml.rels";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

impl XlsxPackage {
    /// Extract images referenced via `xl/richData/_rels/richValueRel*.xml.rels`.
    ///
    /// See [`extract_rich_value_rel_images`] for details.
    pub fn extract_rich_data_images(&self) -> Result<BTreeMap<String, Vec<u8>>, XlsxError> {
        extract_rich_value_rel_images(self)
    }
}

/// Extract image bytes referenced via richData relationships.
///
/// Excel stores "cell images" (and other rich data types) using `xl/richData/*` parts and
/// relationship parts under `xl/richData/_rels/richValueRel*.xml.rels`.
///
/// Some producers may emit:
/// - `TargetMode="External"` relationship targets pointing at URLs (not package parts)
/// - absolute targets (e.g. `/xl/media/image1.png`)
///
/// This helper:
/// - skips external relationships
/// - resolves absolute targets to their normalized in-package part name (no leading `/`)
/// - resolves relative targets against the source part directory (`xl/richData/…`, not `_rels/`)
pub(crate) fn extract_rich_value_rel_images(
    pkg: &XlsxPackage,
) -> Result<BTreeMap<String, Vec<u8>>, XlsxError> {
    let mut out: BTreeMap<String, Vec<u8>> = BTreeMap::new();

    for (rels_part, rels_bytes) in pkg.parts() {
        if !rels_part.starts_with(RICH_VALUE_REL_RELS_PREFIX)
            || !rels_part.ends_with(RICH_VALUE_REL_RELS_SUFFIX)
        {
            continue;
        }

        let Some(source_part) = source_part_from_rels_part(rels_part) else {
            continue;
        };

        let relationships = openxml::parse_relationships(rels_bytes)?;
        for rel in relationships {
            if rel.type_uri != REL_TYPE_IMAGE {
                continue;
            }

            let Some(target_part) = resolve_internal_relationship_target(&source_part, &rel) else {
                continue;
            };

            if out.contains_key(&target_part) {
                continue;
            }

            let Some(bytes) = pkg.part(&target_part) else {
                continue;
            };
            out.insert(target_part, bytes.to_vec());
        }
    }

    Ok(out)
}

fn resolve_internal_relationship_target(
    source_part: &str,
    rel: &openxml::Relationship,
) -> Option<String> {
    if rel
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
    {
        return None;
    }

    Some(resolve_target(source_part, &rel.target))
}

fn source_part_from_rels_part(rels_part: &str) -> Option<String> {
    // Root relationships.
    if rels_part == "_rels/.rels" {
        return Some(String::new());
    }

    if let Some(rels_file) = rels_part.strip_prefix("_rels/") {
        let rels_file = rels_file.strip_suffix(".rels")?;
        return Some(rels_file.to_string());
    }

    let (dir, rels_file) = rels_part.rsplit_once("/_rels/")?;
    let rels_file = rels_file.strip_suffix(".rels")?;

    if dir.is_empty() {
        Some(rels_file.to_string())
    } else {
        Some(format!("{dir}/{rels_file}"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn rich_value_rel_relationships_skip_external_targets_and_resolve_internal() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/image.png" TargetMode="External"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/richData/_rels/richValueRel.xml.rels", rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let extracted = extract_rich_value_rel_images(&pkg).expect("extract images");
        assert_eq!(extracted.len(), 1, "expected external relationship to be skipped");
        assert_eq!(
            extracted.get("xl/media/image1.png").map(Vec::as_slice),
            Some(b"png-bytes".as_slice())
        );
    }

    #[test]
    fn rich_value_rel_relationships_resolve_absolute_targets() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/xl/media/image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/richData/_rels/richValueRel.xml.rels", rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let extracted = extract_rich_value_rel_images(&pkg).expect("extract images");
        assert_eq!(extracted.len(), 1);
        assert_eq!(
            extracted.get("xl/media/image1.png").map(Vec::as_slice),
            Some(b"png-bytes".as_slice())
        );
    }
}
