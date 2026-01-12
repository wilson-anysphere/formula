use crate::{path, XlsxError, XlsxPackage};

/// Discover Rich Data part names referenced by `xl/metadata.xml`.
///
/// Excel stores rich-data payloads under `xl/richData/*`, but the relationship `Type` values
/// used to reference them are not stable across versions. To locate these parts without
/// hard-coding type URIs, this helper inspects `xl/_rels/metadata.xml.rels` and returns the
/// subset of relationship targets that resolve into `xl/richData/` and exist in the package.
///
/// If `xl/metadata.xml` is absent, this returns an empty vector.
pub fn discover_rich_data_part_names(pkg: &XlsxPackage) -> Result<Vec<String>, XlsxError> {
    if pkg.part("xl/metadata.xml").is_none() {
        return Ok(Vec::new());
    }

    let rels_name = crate::openxml::rels_part_name("xl/metadata.xml");
    let Some(rels_bytes) = pkg.part(&rels_name) else {
        return Ok(Vec::new());
    };

    let relationships = crate::openxml::parse_relationships(rels_bytes)?;
    let mut out = Vec::new();
    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        // Relationship targets are URIs and may include fragments; resolve to an OPC part name.
        let target = path::resolve_target("xl/metadata.xml", &rel.target);
        if !target.starts_with("xl/richData/") {
            continue;
        }
        if pkg.part(&target).is_none() {
            continue;
        }
        out.push(target);
    }

    out.sort();
    out.dedup();
    Ok(out)
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
    fn returns_empty_when_metadata_is_absent() {
        let pkg = build_package(&[]);
        let discovered = discover_rich_data_part_names(&pkg).expect("discover rich data");
        assert!(discovered.is_empty());
    }

    #[test]
    fn discovers_rich_data_parts_from_metadata_relationships() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:rich-data" Target="richData/rd2.xml#frag"/>
  <Relationship Id="rId2" Type="urn:example:rich-data" Target="richData/rd1.xml#frag"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/metadata.xml", br#"<metadata/>"#),
            ("xl/_rels/metadata.xml.rels", rels),
            ("xl/richData/rd1.xml", br#"<rd/>"#),
            ("xl/richData/rd2.xml", br#"<rd/>"#),
        ]);

        let discovered = discover_rich_data_part_names(&pkg).expect("discover rich data");
        assert_eq!(
            discovered,
            vec!["xl/richData/rd1.xml".to_string(), "xl/richData/rd2.xml".to_string()]
        );
    }

    #[test]
    fn ignores_missing_referenced_parts() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:rich-data" Target="richData/present.xml#frag"/>
  <Relationship Id="rId2" Type="urn:example:rich-data" Target="richData/missing.xml#frag"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/metadata.xml", br#"<metadata/>"#),
            ("xl/_rels/metadata.xml.rels", rels),
            ("xl/richData/present.xml", br#"<rd/>"#),
        ]);

        let discovered = discover_rich_data_part_names(&pkg).expect("discover rich data");
        assert_eq!(discovered, vec!["xl/richData/present.xml".to_string()]);
    }

    #[test]
    fn ignores_external_relationships_even_when_target_exists() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:rich-data" Target="richData/rd1.xml" TargetMode="External"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/metadata.xml", br#"<metadata/>"#),
            ("xl/_rels/metadata.xml.rels", rels),
            ("xl/richData/rd1.xml", br#"<rd/>"#),
        ]);

        let discovered = discover_rich_data_part_names(&pkg).expect("discover rich data");
        assert!(discovered.is_empty());
    }
}
