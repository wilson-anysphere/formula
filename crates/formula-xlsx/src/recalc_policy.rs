use std::collections::BTreeMap;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::xml::workbook_xml_namespaces_from_workbook_start;

/// Policy describing how writers should ensure Excel recalculates formulas after edits.
///
/// Excel workbooks may contain both cached `<v>` values and an optional `xl/calcChain.xml`. If we
/// edit formulas without updating the calc chain, Excel can open the file with stale calculation
/// state. The safest approach is to drop the calc chain and request a full calculation on load.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecalcPolicy {
    /// When a formula-affecting edit occurs, set `<calcPr fullCalcOnLoad="1"/>` in `xl/workbook.xml`.
    pub force_full_calc_on_formula_change: bool,
    /// When a formula-affecting edit occurs, remove `xl/calcChain.xml` and the associated metadata
    /// entries from the package (`[Content_Types].xml` + `xl/_rels/workbook.xml.rels`).
    pub drop_calc_chain_on_formula_change: bool,
    /// When a formula-affecting edit occurs, clear cached `<v>` values for edited formula cells.
    ///
    /// This can help avoid Excel briefly displaying stale values before recalculation on open.
    /// Default: `false`.
    pub clear_cached_values_on_formula_change: bool,
}

impl RecalcPolicy {
    /// Preserve existing calculation metadata (do nothing).
    pub const PRESERVE: Self = Self {
        force_full_calc_on_formula_change: false,
        drop_calc_chain_on_formula_change: false,
        clear_cached_values_on_formula_change: false,
    };
}

impl Default for RecalcPolicy {
    fn default() -> Self {
        Self {
            force_full_calc_on_formula_change: true,
            drop_calc_chain_on_formula_change: true,
            clear_cached_values_on_formula_change: false,
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub(crate) enum RecalcPolicyError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
}

pub(crate) fn apply_recalc_policy_to_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    policy: RecalcPolicy,
) -> Result<(), RecalcPolicyError> {
    if !policy.force_full_calc_on_formula_change && !policy.drop_calc_chain_on_formula_change {
        return Ok(());
    }

    fn part_key_variants(
        parts: &BTreeMap<String, Vec<u8>>,
        canonical: &str,
    ) -> Result<Vec<String>, RecalcPolicyError> {
        let mut out: Vec<String> = Vec::new();
        out.try_reserve(2)
            .map_err(|_| RecalcPolicyError::AllocationFailure("recalc_policy part_key_variants"))?;

        if let Some((key, _)) = parts.get_key_value(canonical) {
            out.push(key.clone());
        }

        // ZIP entry names should not start with `/`, but tolerate producers that include it by
        // matching the existing key rather than allocating `format!("/{canonical}")`.
        for key in parts.keys() {
            let Some(stripped) = key.strip_prefix('/') else {
                continue;
            };
            if stripped == canonical {
                out.push(key.clone());
                break;
            }
        }

        Ok(out)
    }

    if policy.force_full_calc_on_formula_change {
        // ZIP entry names should not start with `/`, but tolerate producers that include it by
        // patching both the canonical and `/`-prefixed variants when present.
        for key in part_key_variants(parts, "xl/workbook.xml")? {
            if let Some(workbook_xml) = parts.get(&key).cloned() {
                let updated = workbook_xml_force_full_calc_on_load(&workbook_xml)?;
                parts.insert(key, updated);
            }
        }
    }

    if policy.drop_calc_chain_on_formula_change {
        for key in part_key_variants(parts, "xl/calcChain.xml")? {
            parts.remove(&key);
        }

        for key in part_key_variants(parts, "xl/_rels/workbook.xml.rels")? {
            if let Some(rels_xml) = parts.get(&key).cloned() {
                let updated = workbook_rels_remove_calc_chain(&rels_xml)?;
                parts.insert(key, updated);
            }
        }

        for key in part_key_variants(parts, "[Content_Types].xml")? {
            if let Some(content_types_xml) = parts.get(&key).cloned() {
                let updated = content_types_remove_calc_chain(&content_types_xml)?;
                parts.insert(key, updated);
            }
        }
    }

    Ok(())
}

pub(crate) fn workbook_xml_force_full_calc_on_load(
    workbook_xml: &[u8],
) -> Result<Vec<u8>, RecalcPolicyError> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(workbook_xml.len() + 64)
        .map_err(|_| RecalcPolicyError::AllocationFailure("workbook_xml_force_full_calc_on_load"))?;
    let mut writer = Writer::new(out);

    let mut buf = Vec::new();
    let mut saw_calc_pr = false;
    let mut in_workbook = false;
    let mut workbook_ns: Option<crate::xml::WorkbookXmlNamespaces> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if e.local_name().as_ref() == b"workbook" => {
                in_workbook = true;
                workbook_ns.get_or_insert(workbook_xml_namespaces_from_workbook_start(e)?);
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                // Some producers emit an entirely empty workbook root (`<workbook/>`). Expand the
                // element so we can safely insert `<calcPr fullCalcOnLoad="1"/>` while preserving
                // the original workbook qualified name + attributes (including namespace decls).
                workbook_ns.get_or_insert(workbook_xml_namespaces_from_workbook_start(e)?);

                let tag = workbook_ns
                    .as_ref()
                    .map(|ns| crate::xml::prefixed_tag(ns.spreadsheetml_prefix.as_deref(), "calcPr"))
                    .unwrap_or_else(|| "calcPr".to_string());
                let mut calc_pr = BytesStart::new(tag.as_str());
                calc_pr.push_attribute(("fullCalcOnLoad", "1"));

                let workbook_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(Event::Start(e.to_owned()))?;
                writer.write_event(Event::Empty(calc_pr.into_owned()))?;
                writer.write_event(Event::End(BytesEnd::new(workbook_tag.as_str())))?;
                saw_calc_pr = true;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"calcPr" => {
                saw_calc_pr = true;
                writer.write_event(Event::Empty(patched_calc_pr(e)?))?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"calcPr" => {
                saw_calc_pr = true;
                writer.write_event(Event::Start(patched_calc_pr(e)?))?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"workbook" => {
                if in_workbook && !saw_calc_pr {
                    let tag = workbook_ns
                        .as_ref()
                        .map(|ns| crate::xml::prefixed_tag(ns.spreadsheetml_prefix.as_deref(), "calcPr"))
                        .unwrap_or_else(|| "calcPr".to_string());
                    let mut calc_pr = BytesStart::new(tag.as_str());
                    calc_pr.push_attribute(("fullCalcOnLoad", "1"));
                    writer.write_event(Event::Empty(calc_pr.into_owned()))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patched_calc_pr(e: &BytesStart<'_>) -> Result<BytesStart<'static>, RecalcPolicyError> {
    let name = e.name();
    let name = std::str::from_utf8(name.as_ref()).unwrap_or("calcPr");
    let mut calc_pr = BytesStart::new(name);
    for attr in e.attributes() {
        let attr = attr?;
        let key = crate::openxml::local_name(attr.key.as_ref());
        if key.eq_ignore_ascii_case(b"fullCalcOnLoad") {
            continue;
        }
        calc_pr.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }
    calc_pr.push_attribute(("fullCalcOnLoad", "1"));
    Ok(calc_pr.into_owned())
}

pub(crate) fn workbook_rels_remove_calc_chain(rels_xml: &[u8]) -> Result<Vec<u8>, RecalcPolicyError> {
    const CALC_CHAIN_REL_TYPE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(rels_xml.len())
        .map_err(|_| RecalcPolicyError::AllocationFailure("workbook_rels_remove_calc_chain"))?;
    let mut writer = Writer::new(out);

    let mut buf = Vec::new();
    let mut skipping = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e)
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if relationship_is_calc_chain(e, CALC_CHAIN_REL_TYPE)? {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if !relationship_is_calc_chain(e, CALC_CHAIN_REL_TYPE)? {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e)
                if skipping && crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                skipping = false;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn relationship_is_calc_chain(e: &BytesStart<'_>, expected_type: &str) -> Result<bool, RecalcPolicyError> {
    let mut rel_type: Option<String> = None;
    let mut target: Option<String> = None;
    for attr in e.attributes() {
        let attr = attr?;
        let value = attr.unescape_value()?.into_owned();
        let key = crate::openxml::local_name(attr.key.as_ref());
        if key.eq_ignore_ascii_case(b"Type") {
            rel_type = Some(value);
        } else if key.eq_ignore_ascii_case(b"Target") {
            target = Some(value);
        }
    }

    if rel_type.as_deref().map(str::trim) == Some(expected_type) {
        return Ok(true);
    }

    // Relationship targets are URIs; some producers include a fragment (e.g. `calcChain.xml#foo`).
    // OPC part names do not include fragments, so strip them before matching.
    Ok(target.as_deref().is_some_and(|t| {
        let t = t.trim();
        let base = t.split_once('#').map(|(base, _)| base).unwrap_or(t);
        let base = base.trim();

        // `xl/_rels/workbook.xml.rels` targets are resolved relative to `xl/workbook.xml`.
        //
        // Match on the resolved part name so we remove calcChain relationships even when the
        // producer uses absolute paths or includes `..` segments. This also avoids false positives
        // for unrelated targets like `worksheets/calcChain.xml`.
        let resolved = crate::path::resolve_target("xl/workbook.xml", base);

        // Some producers incorrectly emit root-relative targets without a leading slash
        // (`xl/calcChain.xml` instead of `/xl/calcChain.xml`). Treat that as calcChain too.
        resolved == "xl/calcChain.xml"
            || base
                .strip_prefix('/')
                .unwrap_or(base)
                .trim()
                .eq("xl/calcChain.xml")
    }))
}

pub(crate) fn content_types_remove_calc_chain(ct_xml: &[u8]) -> Result<Vec<u8>, RecalcPolicyError> {
    let mut reader = Reader::from_reader(ct_xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(ct_xml.len())
        .map_err(|_| RecalcPolicyError::AllocationFailure("content_types_remove_calc_chain"))?;
    let mut writer = Writer::new(out);

    let mut buf = Vec::new();
    let mut skipping = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            // `[Content_Types].xml` can be either prefix-free (`<Override>`) or namespace-prefixed
            // (`<ct:Override>`). Match on local name so we can remove `xl/calcChain.xml` overrides
            // in both forms, while preserving the original prefixes when writing other events.
            Event::Start(ref e)
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                if override_part_name_is_calc_chain(e)? {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                if !override_part_name_is_calc_chain(e)? {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e)
                if skipping && crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                skipping = false;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn override_part_name_is_calc_chain(e: &BytesStart<'_>) -> Result<bool, RecalcPolicyError> {
    for attr in e.attributes() {
        let attr = attr?;
        let key = crate::openxml::local_name(attr.key.as_ref());
        if key.eq_ignore_ascii_case(b"PartName") {
            let value = attr.unescape_value()?;
            let value = value.as_ref().trim();
            let base = value
                .split_once('#')
                .map(|(base, _)| base)
                .unwrap_or(value);
            let normalized = base
                .trim()
                .strip_prefix('/')
                .unwrap_or(base)
                .trim();
            return Ok(normalized == "xl/calcChain.xml");
        }
    }
    Ok(false)
}

#[cfg(test)]
mod tests {
    use std::collections::{BTreeMap, BTreeSet};

    use super::*;

    fn content_types_override_part_names(xml: &[u8]) -> BTreeSet<String> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        let mut parts = BTreeSet::new();

        loop {
            match reader
                .read_event_into(&mut buf)
                .expect("read content types xml event")
            {
                Event::Eof => break,
                Event::Start(ref e) | Event::Empty(ref e)
                    if crate::openxml::local_name(e.name().as_ref())
                        .eq_ignore_ascii_case(b"Override") =>
                {
                    for attr in e.attributes() {
                        let attr = attr.expect("read Override attribute");
                        let key = crate::openxml::local_name(attr.key.as_ref());
                        if !key.eq_ignore_ascii_case(b"PartName") {
                            continue;
                        }
                        let part_name = attr
                            .unescape_value()
                            .expect("unescape PartName value")
                            .into_owned();
                        parts.insert(part_name.to_string());
                    }
                }
                _ => {}
            }
            buf.clear();
        }

        parts
    }

    #[test]
    fn apply_recalc_policy_to_parts_tolerates_slash_prefixed_part_names() {
        // ZIP entry names in valid XLSX packages should not start with `/`, but some producers
        // may include it. Recalc policy should still patch workbook / rels / content types and
        // remove calcChain in that case.
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <calcPr fullCalcOnLoad="0" calcId="171027"/>
</workbook>
"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

        let content_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>
"#;

        let mut parts = BTreeMap::new();
        parts.insert("/xl/workbook.xml".to_string(), workbook_xml.to_vec());
        parts.insert("/xl/_rels/workbook.xml.rels".to_string(), rels_xml.to_vec());
        parts.insert("/[Content_Types].xml".to_string(), content_types_xml.to_vec());
        parts.insert("/xl/calcChain.xml".to_string(), b"<calcChain/>".to_vec());

        apply_recalc_policy_to_parts(&mut parts, RecalcPolicy::default())
            .expect("apply recalc policy to slashed parts");

        assert!(
            !parts.contains_key("/xl/calcChain.xml"),
            "calcChain part should be removed"
        );
        assert!(
            !parts.contains_key("xl/calcChain.xml"),
            "should not create canonical calcChain key while removing"
        );

        assert!(
            !parts.contains_key("xl/workbook.xml"),
            "expected workbook to be patched in-place (not duplicated without slash)"
        );
        let updated_workbook =
            std::str::from_utf8(parts.get("/xl/workbook.xml").expect("workbook present"))
                .expect("utf8 workbook");
        assert!(
            updated_workbook.contains(r#"fullCalcOnLoad="1""#),
            "expected workbook calcPr fullCalcOnLoad=1, got: {updated_workbook}"
        );

        assert!(
            !parts.contains_key("xl/_rels/workbook.xml.rels"),
            "expected rels to be patched in-place (not duplicated without slash)"
        );
        let updated_rels = std::str::from_utf8(
            parts
                .get("/xl/_rels/workbook.xml.rels")
                .expect("workbook rels present"),
        )
        .expect("utf8 rels");
        assert!(
            !updated_rels.contains("calcChain.xml"),
            "expected calcChain relationship to be removed, got: {updated_rels}"
        );
        assert!(
            updated_rels.contains("metadata.xml"),
            "expected metadata relationship to be preserved, got: {updated_rels}"
        );

        assert!(
            !parts.contains_key("[Content_Types].xml"),
            "expected content types to be patched in-place (not duplicated without slash)"
        );
        let updated_ct =
            std::str::from_utf8(parts.get("/[Content_Types].xml").expect("ct present"))
                .expect("utf8 content types");
        assert!(
            !updated_ct.contains("calcChain.xml"),
            "expected calcChain override to be removed, got: {updated_ct}"
        );
        assert!(
            updated_ct.contains("/xl/metadata.xml"),
            "expected metadata override to be preserved, got: {updated_ct}"
        );
    }

    #[test]
    fn content_types_remove_calc_chain_preserves_metadata_and_richdata_overrides() {
        // Regression test: removing calcChain overrides must not discard richData / metadata parts
        // used for embedded images.
        let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></Override>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
  <Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>"#;

        let updated = content_types_remove_calc_chain(content_types_xml.as_bytes())
            .expect("remove calc chain from content types");
        let updated_overrides = content_types_override_part_names(&updated);

        assert!(
            !updated_overrides.contains("/xl/calcChain.xml"),
            "calcChain override should be removed"
        );

        let expected: BTreeSet<String> = [
            "/xl/metadata.xml",
            "/xl/richData/rdrichvalue.xml",
            "/xl/richData/rdrichvaluestructure.xml",
            "/xl/richData/rdRichValueTypes.xml",
            "/xl/richData/richValueRel.xml",
        ]
        .into_iter()
        .map(String::from)
        .collect();

        assert_eq!(updated_overrides, expected);
    }

    #[test]
    fn content_types_remove_calc_chain_removes_prefixed_override() {
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</ct:Types>
"#;

        let updated =
            content_types_remove_calc_chain(input.as_bytes()).expect("rewrite content types");
        let updated = std::str::from_utf8(&updated).expect("utf8 updated content types");

        // Calc chain override removed.
        assert!(!updated.contains("calcChain.xml"));

        // Preserve original prefixes for unrelated elements.
        assert!(updated.contains("<ct:Types"));
    }

    #[test]
    fn content_types_remove_calc_chain_removes_prefixed_non_empty_override() {
        // Some producers emit overrides with explicit end tags (not self-closing). Ensure we still
        // remove calcChain overrides in the prefixed form, while preserving other overrides.
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></ct:Override>
  <ct:Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</ct:Types>
"#;

        let updated =
            content_types_remove_calc_chain(input.as_bytes()).expect("rewrite content types");
        let updated = std::str::from_utf8(&updated).expect("utf8 updated content types");

        assert!(
            !updated.contains("calcChain.xml"),
            "calcChain override should be removed, got: {updated}"
        );
        assert!(
            updated.contains(r#"PartName="/xl/metadata.xml""#),
            "metadata override should be preserved, got: {updated}"
        );
        assert!(
            updated.contains("<ct:Types"),
            "expected root prefix to be preserved, got: {updated}"
        );
        assert!(
            updated.contains("<ct:Override"),
            "expected Override prefix to be preserved, got: {updated}"
        );
    }

    #[test]
    fn content_types_remove_calc_chain_removes_override_with_fragment_partname() {
        // Relationship targets can have fragments; PartName shouldn't, but be permissive and remove
        // calcChain overrides even if a fragment is present.
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml#foo" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>"#;

        let updated =
            content_types_remove_calc_chain(input.as_bytes()).expect("rewrite content types");
        let updated = std::str::from_utf8(&updated).expect("utf8 updated content types");

        assert!(
            !updated.contains("calcChain.xml"),
            "calcChain override should be removed, got: {updated}"
        );
        assert!(
            updated.contains("/xl/metadata.xml"),
            "metadata override should be preserved, got: {updated}"
        );
    }

    #[test]
    fn content_types_remove_calc_chain_preserves_similarly_named_overrides() {
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/mycalcChain.xml" ContentType="application/xml"/>
</Types>"#;

        let updated = content_types_remove_calc_chain(input.as_bytes())
            .expect("remove calc chain from content types");
        let updated_overrides = content_types_override_part_names(&updated);

        assert!(
            !updated_overrides.contains("/xl/calcChain.xml"),
            "calcChain override should be removed"
        );
        assert!(
            updated_overrides.contains("/xl/mycalcChain.xml"),
            "expected similarly-named override to be preserved"
        );
    }

    #[test]
    fn content_types_remove_calc_chain_preserves_other_calcchain_named_parts() {
        // Be strict about removing only `/xl/calcChain.xml`; other parts might legitimately use the
        // same filename.
        let input = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/docProps/calcChain.xml" ContentType="application/xml"/>
</Types>"#;

        let updated = content_types_remove_calc_chain(input.as_bytes())
            .expect("remove calc chain from content types");
        let updated_overrides = content_types_override_part_names(&updated);

        assert!(
            !updated_overrides.contains("/xl/calcChain.xml"),
            "calcChain override should be removed"
        );
        assert!(
            updated_overrides.contains("/docProps/calcChain.xml"),
            "expected non-xl calcChain override to be preserved"
        );
    }

    #[test]
    fn workbook_xml_force_full_calc_on_load_patches_prefixed_calc_pr() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:calcPr fullCalcOnLoad="0" calcId="171027"/>
</x:workbook>
"#;

        let updated =
            workbook_xml_force_full_calc_on_load(workbook_xml.as_bytes()).expect("patch workbook");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook");

        assert!(updated.contains("<x:calcPr"), "expected calcPr prefix to be preserved");
        assert!(
            updated.contains(r#"fullCalcOnLoad="1""#),
            "expected fullCalcOnLoad to be set to 1"
        );
        assert!(
            updated.contains(r#"calcId="171027""#),
            "expected other calcPr attributes to be preserved"
        );
    }

    #[test]
    fn workbook_xml_force_full_calc_on_load_inserts_prefixed_calc_pr_when_missing() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</x:workbook>
"#;

        let updated =
            workbook_xml_force_full_calc_on_load(workbook_xml.as_bytes()).expect("patch workbook");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook");

        assert!(
            updated.contains("<x:workbook"),
            "expected workbook prefix to be preserved"
        );
        assert!(
            updated.contains("<x:calcPr"),
            "expected inserted calcPr to use workbook prefix"
        );
        assert!(
            updated.contains(r#"fullCalcOnLoad="1""#),
            "expected inserted calcPr fullCalcOnLoad=1"
        );
    }

    #[test]
    fn workbook_xml_force_full_calc_on_load_prefers_workbook_element_prefix_when_multiple_prefixes_bind_spreadsheetml(
    ) {
        // Some producers declare multiple prefixes for SpreadsheetML. When we need to insert a
        // missing `<calcPr/>`, we should use the prefix that matches the `<workbook>` element.
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:y="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
</x:workbook>
"#;

        let updated =
            workbook_xml_force_full_calc_on_load(workbook_xml.as_bytes()).expect("patch workbook");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook");

        assert!(
            updated.contains("<x:calcPr"),
            "expected inserted calcPr to use the workbook element prefix, got: {updated}"
        );
        assert!(
            !updated.contains("<y:calcPr"),
            "expected calcPr not to use an unrelated SpreadsheetML prefix, got: {updated}"
        );
        assert!(updated.contains(r#"fullCalcOnLoad="1""#));
    }

    #[test]
    fn workbook_rels_remove_calc_chain_preserves_other_relationships_and_prefixes() {
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <r:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <r:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <r:Relationship Id="rId3" Type="http://example.com/keep" Target="xl/calcChain.xml"/>
  <r:Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</r:Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook rels");

        assert!(
            updated.contains(r#"Id="rId1""#),
            "expected worksheet relationship to be preserved, got: {updated}"
        );
        assert!(
            !updated.contains("calcChain.xml"),
            "expected calc chain relationship to be removed, got: {updated}"
        );
        assert!(
            updated.contains("metadata.xml"),
            "expected metadata relationship to be preserved, got: {updated}"
        );
        assert!(
            updated.contains(r#"Id="rId9""#),
            "expected metadata relationship to be preserved, got: {updated}"
        );
        assert!(
            updated.contains("<r:Relationships"),
            "expected root element prefix to be preserved, got: {updated}"
        );
        assert!(
            updated.contains("<r:Relationship"),
            "expected Relationship element prefix to be preserved, got: {updated}"
        );
    }

    #[test]
    fn workbook_rels_remove_calc_chain_removes_by_target_even_if_type_is_unexpected() {
        // Some producers may emit a nonstandard relationship type for calcChain, but still target
        // `calcChain.xml`. The recalc policy should be tolerant and remove it based on the target.
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://example.com/not-calc-chain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook rels");

        assert!(
            !updated.contains("calcChain.xml"),
            "expected calc chain relationship to be removed, got: {updated}"
        );
        assert!(
            updated.contains("metadata.xml"),
            "expected metadata relationship to be preserved, got: {updated}"
        );
    }

    #[test]
    fn workbook_rels_remove_calc_chain_removes_relationship_target_with_fragment() {
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://example.com/not-calc-chain" Target="calcChain.xml#foo"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook rels");

        assert!(
            !updated.contains("calcChain.xml"),
            "expected calc chain relationship to be removed, got: {updated}"
        );
        assert!(
            updated.contains("metadata.xml"),
            "expected metadata relationship to be preserved, got: {updated}"
        );
    }

    #[test]
    fn workbook_rels_remove_calc_chain_preserves_similarly_named_targets() {
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://example.com/not-calc-chain" Target="mycalcChain.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");

        let rels = crate::openxml::parse_relationships(&updated)
            .expect("parse updated workbook relationships");
        let targets: BTreeSet<String> = rels.into_iter().map(|r| r.target).collect();

        assert!(
            !targets.contains("calcChain.xml"),
            "expected calcChain.xml relationship to be removed"
        );
        assert!(
            targets.contains("mycalcChain.xml"),
            "expected similarly-named relationship target to be preserved"
        );
        assert!(
            targets.contains("metadata.xml"),
            "expected metadata relationship to be preserved"
        );
    }

    #[test]
    fn workbook_rels_remove_calc_chain_preserves_calcchain_named_targets_in_other_dirs() {
        // Remove calcChain relationships based on their resolved target, not just the filename.
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/keep" Target="worksheets/calcChain.xml"/>
  <Relationship Id="rId2" Type="http://example.com/not-calc-chain" Target="calcChain.xml"/>
</Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");

        let rels = crate::openxml::parse_relationships(&updated)
            .expect("parse updated workbook relationships");
        let targets: BTreeSet<String> = rels.into_iter().map(|r| r.target).collect();

        assert!(
            !targets.contains("calcChain.xml"),
            "expected calcChain.xml relationship to be removed"
        );
        assert!(
            targets.contains("worksheets/calcChain.xml"),
            "expected worksheets/calcChain.xml relationship to be preserved"
        );
    }

    #[test]
    fn workbook_xml_force_full_calc_on_load_expands_prefixed_self_closing_workbook_root() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"#;

        let updated =
            workbook_xml_force_full_calc_on_load(workbook_xml.as_bytes()).expect("patch workbook");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook");

        assert!(
            updated.contains("<x:workbook"),
            "expected workbook prefix to be preserved"
        );
        assert!(
            updated.contains("</x:workbook>"),
            "expected workbook root to be expanded (not self-closing)"
        );
        assert!(
            updated.contains("<x:calcPr"),
            "expected inserted calcPr to use workbook prefix"
        );
        assert!(
            updated.contains(r#"fullCalcOnLoad="1""#),
            "expected inserted calcPr fullCalcOnLoad=1"
        );

        roxmltree::Document::parse(updated).expect("updated workbook xml should parse");
    }

    #[test]
    fn workbook_xml_force_full_calc_on_load_expands_default_ns_self_closing_workbook_root() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"#;

        let updated =
            workbook_xml_force_full_calc_on_load(workbook_xml.as_bytes()).expect("patch workbook");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook");

        assert!(updated.contains("<workbook"));
        assert!(
            updated.contains("</workbook>"),
            "expected workbook root to be expanded (not self-closing)"
        );
        assert!(
            updated.contains("<calcPr"),
            "expected inserted calcPr to be unprefixed in default-namespace workbooks"
        );
        assert!(updated.contains(r#"fullCalcOnLoad="1""#));

        roxmltree::Document::parse(updated).expect("updated workbook xml should parse");
    }

    #[test]
    fn workbook_rels_remove_calc_chain_removes_non_empty_relationship_element() {
        // Some producers emit relationships with explicit end tags (not self-closing). Ensure we
        // still remove calcChain and preserve other relationships.
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <r:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"></r:Relationship>
  <r:Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"></r:Relationship>
</r:Relationships>
"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes())
            .expect("remove calc chain relationship from workbook.xml.rels");
        let updated = std::str::from_utf8(&updated).expect("utf8 workbook rels");

        assert!(
            !updated.contains("calcChain.xml"),
            "expected calc chain relationship to be removed, got: {updated}"
        );
        assert!(
            updated.contains(r#"Id="rId9""#) && updated.contains(r#"Target="metadata.xml""#),
            "expected metadata relationship to be preserved, got: {updated}"
        );
        assert!(
            updated.contains("<r:Relationship"),
            "expected relationship elements to keep their prefix, got: {updated}"
        );
    }

    #[test]
    fn workbook_rels_remove_calc_chain_preserves_richdata_relationships() {
        // Regression test: Excel embedded images-in-cells rely on Office richData relationships.
        // Our recalc-policy rewrite must only remove the calcChain relationship.
        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="sheetMetadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdRichValue.xml"/>
  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdRichValueStructure.xml"/>
  <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes" Target="richData/rdRichValueTypes.xml"/>
</Relationships>"#;

        let updated = workbook_rels_remove_calc_chain(rels_xml.as_bytes()).expect("rewrite rels");
        let updated = String::from_utf8(updated).expect("xml utf-8");

        assert!(
            !updated.contains("relationships/calcChain"),
            "expected calcChain relationship to be removed, got:\n{updated}"
        );
        assert!(
            !updated.contains("calcChain.xml"),
            "expected calcChain Target to be removed, got:\n{updated}"
        );

        for rel_type in [
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
            "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel",
            "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue",
            "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure",
            "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes",
        ] {
            assert!(
                updated.contains(rel_type),
                "expected relationship Type to be preserved: {rel_type}, got:\n{updated}"
            );
        }
    }
}
