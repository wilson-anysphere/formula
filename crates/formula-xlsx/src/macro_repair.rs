use std::collections::BTreeMap;

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};

use crate::openxml::local_name;
use crate::XlsxError;

fn select_part_key(parts: &BTreeMap<String, Vec<u8>>, name: &str) -> Option<String> {
    if parts.contains_key(name) {
        return Some(name.to_string());
    }

    if let Some(stripped) = name.strip_prefix('/').or_else(|| name.strip_prefix('\\')) {
        if parts.contains_key(stripped) {
            return Some(stripped.to_string());
        }
    } else {
        // Some producers incorrectly store OPC part names with a leading `/` in the ZIP.
        // Preserve exact names for round-trip, but make lookups resilient.
        let with_slash = format!("/{name}");
        if parts.contains_key(&with_slash) {
            return Some(with_slash);
        }
    }

    parts
        .keys()
        .find(|key| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
        .cloned()
}

fn part_exists(parts: &BTreeMap<String, Vec<u8>>, name: &str) -> bool {
    select_part_key(parts, name).is_some()
}

fn relationship_id_number(id: &str) -> Option<u32> {
    let id = id.trim();
    let bytes = id.as_bytes();
    if bytes.len() < 4 {
        return None;
    }
    if !(matches!(bytes[0], b'r' | b'R')
        && matches!(bytes[1], b'i' | b'I')
        && matches!(bytes[2], b'd' | b'D'))
    {
        return None;
    }
    let digits = &id[3..];
    if digits.is_empty() || !digits.chars().all(|ch| ch.is_ascii_digit()) {
        return None;
    }
    digits.parse::<u32>().ok()
}

fn prefixed_tag(container_name: &[u8], local: &str) -> String {
    match container_name.iter().position(|&b| b == b':') {
        Some(idx) => {
            let prefix = std::str::from_utf8(&container_name[..idx]).unwrap_or_default();
            format!("{prefix}:{local}")
        }
        None => local.to_string(),
    }
}

pub(crate) fn ensure_xlsm_content_types(
    parts: &mut BTreeMap<String, Vec<u8>>,
) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let Some(ct_key) = select_part_key(parts, ct_name) else {
        // We don't attempt to synthesize a full content types file; macro
        // preservation in this minimal crate assumes an existing workbook.
        return Ok(());
    };
    let Some(existing) = parts.get(&ct_key).cloned() else {
        debug_assert!(false, "content types key resolved but missing");
        return Err(XlsxError::MissingPart(ct_key));
    };

    // Only repair to XLSM if the package actually contains the VBA project payload.
    if !part_exists(parts, "xl/vbaProject.bin") {
        return Ok(());
    }

    const WORKBOOK_PART_NAME: &str = "/xl/workbook.xml";
    // When a workbook contains `xl/vbaProject.bin`, Excel expects the main workbook content type
    // to be one of the macro-enabled types:
    // - `.xlsm`  → `sheet.macroEnabled`
    // - `.xltm`  → `template.macroEnabled`
    // - `.xlam`  → `addin.macroEnabled`
    //
    // This helper is invoked during `XlsxPackage::write_to(...)` to repair the package before
    // writing. Importantly: callers may have already set a more specific macro-enabled workbook
    // kind (e.g. `.xltm` / `.xlam`). Do not blindly rewrite those to `.xlsm`.
    const WORKBOOK_MACRO_CONTENT_TYPE: &str =
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    const WORKBOOK_TEMPLATE_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
    const WORKBOOK_MACRO_TEMPLATE_CONTENT_TYPE: &str =
        "application/vnd.ms-excel.template.macroEnabled.main+xml";
    const WORKBOOK_ADDIN_MACRO_CONTENT_TYPE: &str =
        "application/vnd.ms-excel.addin.macroEnabled.main+xml";
    const VBA_PART_NAME: &str = "/xl/vbaProject.bin";
    const VBA_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProject";
    const VBA_SIGNATURE_PART_NAME: &str = "/xl/vbaProjectSignature.bin";
    const VBA_SIGNATURE_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaProjectSignature";
    const VBA_DATA_PART_NAME: &str = "/xl/vbaData.xml";
    const VBA_DATA_CONTENT_TYPE: &str = "application/vnd.ms-office.vbaData+xml";

    let needs_signature = part_exists(parts, "xl/vbaProjectSignature.bin");
    let needs_vba_data = part_exists(parts, "xl/vbaData.xml");

    let mut reader = XmlReader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let Some(cap) = existing.len().checked_add(256) else {
        return Err(XlsxError::AllocationFailure("macro_repair [Content_Types].xml output"));
    };
    if out.try_reserve(cap).is_err() {
        return Err(XlsxError::AllocationFailure("macro_repair [Content_Types].xml output"));
    }
    let mut writer = XmlWriter::new(out);
    let mut buf = Vec::new();

    let mut override_tag_name: Option<String> = None;
    let mut default_tag_name: Option<String> = None;

    let mut has_workbook_override = false;
    let mut has_vba_override = false;
    let mut has_vba_signature_override = false;
    let mut has_vba_data_override = false;
    let mut changed = false;

    fn is_macro_enabled_workbook_content_type(content_type: &str) -> bool {
        // Preserve any macro-enabled workbook kind (`.xlsm`, `.xltm`, `.xlam`) rather than
        // unconditionally forcing `.xlsm`. This allows callers to patch `[Content_Types].xml`
        // for templates/add-ins and still rely on macro repair to inject the other required
        // VBA overrides.
        let content_type = content_type.trim();
        content_type.eq_ignore_ascii_case(WORKBOOK_MACRO_CONTENT_TYPE)
            || content_type.eq_ignore_ascii_case(WORKBOOK_MACRO_TEMPLATE_CONTENT_TYPE)
            || content_type.eq_ignore_ascii_case(WORKBOOK_ADDIN_MACRO_CONTENT_TYPE)
    }

    fn part_name_matches(candidate: &str, expected: &str) -> bool {
        crate::zip_util::zip_part_names_equivalent(candidate.trim(), expected.trim())
    }

    fn handle_override(
        e: BytesStart<'_>,
        is_start: bool,
        needs_signature: bool,
        needs_vba_data: bool,
        writer: &mut XmlWriter<Vec<u8>>,
        changed: &mut bool,
        has_workbook_override: &mut bool,
        has_vba_override: &mut bool,
        has_vba_signature_override: &mut bool,
        has_vba_data_override: &mut bool,
    ) -> Result<(), XlsxError> {
        let mut part_name = None;
        let mut content_type = None;
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            match local_name(attr.key.as_ref()) {
                b"PartName" => part_name = Some(attr.unescape_value()?.into_owned()),
                b"ContentType" => content_type = Some(attr.unescape_value()?.into_owned()),
                _ => {}
            }
        }

        let desired_content_type = match part_name.as_deref() {
            Some(part) if part_name_matches(part, WORKBOOK_PART_NAME) => {
                *has_workbook_override = true;
                // Preserve any macro-enabled workbook kind (`.xlsm`, `.xltm`, `.xlam`) rather than
                // unconditionally forcing `.xlsm`. This allows callers to patch
                // `[Content_Types].xml` for templates/add-ins and still rely on macro repair to
                // inject the other required VBA overrides.
                match content_type.as_deref() {
                    Some(existing) if is_macro_enabled_workbook_content_type(existing) => None,
                    Some(existing)
                        if existing
                            .trim()
                            .eq_ignore_ascii_case(WORKBOOK_TEMPLATE_CONTENT_TYPE) =>
                    {
                        Some(WORKBOOK_MACRO_TEMPLATE_CONTENT_TYPE)
                    }
                    _ => Some(WORKBOOK_MACRO_CONTENT_TYPE),
                }
            }
            Some(part) if part_name_matches(part, VBA_PART_NAME) => {
                *has_vba_override = true;
                Some(VBA_CONTENT_TYPE)
            }
            Some(part) if needs_signature && part_name_matches(part, VBA_SIGNATURE_PART_NAME) => {
                *has_vba_signature_override = true;
                Some(VBA_SIGNATURE_CONTENT_TYPE)
            }
            Some(part) if needs_vba_data && part_name_matches(part, VBA_DATA_PART_NAME) => {
                *has_vba_data_override = true;
                Some(VBA_DATA_CONTENT_TYPE)
            }
            _ => None,
        };

        if let Some(desired_content_type) = desired_content_type {
            if content_type.as_deref() != Some(desired_content_type) {
                *changed = true;

                let tag_name = e.name();
                let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Override");
                let mut patched = BytesStart::new(tag_name);
                let mut saw_content_type = false;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"ContentType") {
                        saw_content_type = true;
                        patched.push_attribute((attr.key.as_ref(), desired_content_type.as_bytes()));
                    } else {
                        patched.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                    }
                }
                if !saw_content_type {
                    patched.push_attribute(("ContentType", desired_content_type));
                }

                if is_start {
                    writer.write_event(Event::Start(patched.into_owned()))?;
                } else {
                    writer.write_event(Event::Empty(patched.into_owned()))?;
                }
            } else if is_start {
                writer.write_event(Event::Start(e))?;
            } else {
                writer.write_event(Event::Empty(e))?;
            }
        } else if is_start {
            writer.write_event(Event::Start(e))?;
        } else {
            writer.write_event(Event::Empty(e))?;
        }

        Ok(())
    }

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                writer.write_event(Event::Empty(e))?;
            }
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") => {
                if override_tag_name.is_none() {
                    override_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                handle_override(
                    e,
                    true,
                    needs_signature,
                    needs_vba_data,
                    &mut writer,
                    &mut changed,
                    &mut has_workbook_override,
                    &mut has_vba_override,
                    &mut has_vba_signature_override,
                    &mut has_vba_data_override,
                )?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") => {
                if override_tag_name.is_none() {
                    override_tag_name = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                handle_override(
                    e,
                    false,
                    needs_signature,
                    needs_vba_data,
                    &mut writer,
                    &mut changed,
                    &mut has_workbook_override,
                    &mut has_vba_override,
                    &mut has_vba_signature_override,
                    &mut has_vba_data_override,
                )?;
            }
            Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                // Prefer using an existing `<Override>` prefix when available. If the file has no
                // overrides (or is missing them due to producer bugs), fall back to the prefix
                // used by `<Default>` entries so we don't inject un-namespaced overrides into a
                // prefix-only content types document.
                let override_tag_name = override_tag_name
                    .clone()
                    .or_else(|| {
                        default_tag_name
                            .as_ref()
                            .map(|tag| prefixed_tag(tag.as_bytes(), "Override"))
                    })
                    .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Override"));

                if !has_workbook_override {
                    changed = true;
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", WORKBOOK_PART_NAME));
                    override_el.push_attribute(("ContentType", WORKBOOK_MACRO_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }

                if !has_vba_override {
                    changed = true;
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", VBA_PART_NAME));
                    override_el.push_attribute(("ContentType", VBA_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }

                if needs_signature && !has_vba_signature_override {
                    changed = true;
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", VBA_SIGNATURE_PART_NAME));
                    override_el.push_attribute(("ContentType", VBA_SIGNATURE_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }

                if needs_vba_data && !has_vba_data_override {
                    changed = true;
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", VBA_DATA_PART_NAME));
                    override_el.push_attribute(("ContentType", VBA_DATA_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }

                writer.write_event(Event::End(e))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                // Degenerate case: a self-closing `<Types/>` root. Expand it so we can inject
                // the required overrides.
                changed = true;
                let types_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let override_tag_name = override_tag_name
                    .clone()
                    .or_else(|| {
                        default_tag_name
                            .as_ref()
                            .map(|tag| prefixed_tag(tag.as_bytes(), "Override"))
                    })
                    .unwrap_or_else(|| prefixed_tag(types_tag_name.as_bytes(), "Override"));

                writer.write_event(Event::Start(e))?;

                let mut workbook_override = BytesStart::new(override_tag_name.as_str());
                workbook_override.push_attribute(("PartName", WORKBOOK_PART_NAME));
                workbook_override.push_attribute(("ContentType", WORKBOOK_MACRO_CONTENT_TYPE));
                writer.write_event(Event::Empty(workbook_override))?;

                let mut vba_override = BytesStart::new(override_tag_name.as_str());
                vba_override.push_attribute(("PartName", VBA_PART_NAME));
                vba_override.push_attribute(("ContentType", VBA_CONTENT_TYPE));
                writer.write_event(Event::Empty(vba_override))?;

                if needs_signature {
                    let mut sig_override = BytesStart::new(override_tag_name.as_str());
                    sig_override.push_attribute(("PartName", VBA_SIGNATURE_PART_NAME));
                    sig_override.push_attribute(("ContentType", VBA_SIGNATURE_CONTENT_TYPE));
                    writer.write_event(Event::Empty(sig_override))?;
                }

                if needs_vba_data {
                    let mut data_override = BytesStart::new(override_tag_name.as_str());
                    data_override.push_attribute(("PartName", VBA_DATA_PART_NAME));
                    data_override.push_attribute(("ContentType", VBA_DATA_CONTENT_TYPE));
                    writer.write_event(Event::Empty(data_override))?;
                }

                writer.write_event(Event::End(BytesEnd::new(types_tag_name.as_str())))?;
            }
            Event::Eof => break,
            other => writer.write_event(other)?,
        }

        buf.clear();
    }

    if changed {
        parts.insert(ct_key, writer.into_inner());
    }
    Ok(())
}

pub(crate) fn ensure_workbook_rels_has_vba(
    parts: &mut BTreeMap<String, Vec<u8>>,
) -> Result<(), XlsxError> {
    let rels_name = "xl/_rels/workbook.xml.rels";
    let Some(rels_key) = select_part_key(parts, rels_name) else {
        return Ok(());
    };
    let existing = parts
        .get(&rels_key)
        .cloned()
        .ok_or_else(|| {
            debug_assert!(false, "workbook rels key resolved but missing");
            XlsxError::MissingPart(rels_key.clone())
        })?;

    const VBA_REL_TYPE: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    const VBA_TARGET: &str = "vbaProject.bin";

    let mut reader = XmlReader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let Some(cap) = existing.len().checked_add(128) else {
        return Err(XlsxError::AllocationFailure("ensure_workbook_rels_has_vba output"));
    };
    if out.try_reserve(cap).is_err() {
        return Err(XlsxError::AllocationFailure("ensure_workbook_rels_has_vba output"));
    }
    let mut writer = XmlWriter::new(out);
    let mut buf = Vec::new();

    let mut relationship_tag_name: Option<String> = None;

    let mut has_vba_rel = false;
    let mut changed = false;
    let mut max_rid = 0u32;

    fn handle_relationship(
        e: BytesStart<'_>,
        is_start: bool,
        writer: &mut XmlWriter<Vec<u8>>,
        changed: &mut bool,
        has_vba_rel: &mut bool,
        max_rid: &mut u32,
    ) -> Result<(), XlsxError> {
        let mut id = None;
        let mut type_uri = None;
        let mut target = None;
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            match local_name(attr.key.as_ref()) {
                b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                b"Type" => type_uri = Some(attr.unescape_value()?.into_owned()),
                b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                _ => {}
            }
        }

        if let Some(id) = id {
            if let Some(n) = relationship_id_number(&id) {
                *max_rid = (*max_rid).max(n);
            }
        }

        let is_vba = type_uri
            .as_deref()
            .is_some_and(|t| t.trim() == VBA_REL_TYPE);
        if is_vba {
            if target
                .as_deref()
                .is_some_and(|t| crate::zip_util::zip_part_names_equivalent(t.trim(), VBA_TARGET))
            {
                *has_vba_rel = true;
                if is_start {
                    writer.write_event(Event::Start(e))?;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            } else {
                // Relationship exists, but is missing/has an unexpected Target; patch it.
                *has_vba_rel = true;
                *changed = true;

                let tag_name = e.name();
                let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Relationship");
                let mut patched = BytesStart::new(tag_name);
                let mut saw_target = false;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Target") {
                        saw_target = true;
                        patched.push_attribute((attr.key.as_ref(), VBA_TARGET.as_bytes()));
                    } else {
                        patched.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                    }
                }
                if !saw_target {
                    patched.push_attribute(("Target", VBA_TARGET));
                }

                if is_start {
                    writer.write_event(Event::Start(patched.into_owned()))?;
                } else {
                    writer.write_event(Event::Empty(patched.into_owned()))?;
                }
            }
        } else if is_start {
            writer.write_event(Event::Start(e))?;
        } else {
            writer.write_event(Event::Empty(e))?;
        }

        Ok(())
    }

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                // Degenerate case: self-closing `<Relationships/>` root; expand to insert the vba
                // relationship.
                changed = true;
                let relationships_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let relationship_tag_name = relationship_tag_name
                    .clone()
                    .unwrap_or_else(|| prefixed_tag(relationships_tag_name.as_bytes(), "Relationship"));

                writer.write_event(Event::Start(e))?;

                let next_rid = max_rid + 1;
                let rel_id = format!("rId{next_rid}");
                let mut rel = BytesStart::new(relationship_tag_name.as_str());
                rel.push_attribute(("Id", rel_id.as_str()));
                rel.push_attribute(("Type", VBA_REL_TYPE));
                rel.push_attribute(("Target", VBA_TARGET));
                writer.write_event(Event::Empty(rel))?;

                writer.write_event(Event::End(BytesEnd::new(relationships_tag_name.as_str())))?;
            }
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                if relationship_tag_name.is_none() {
                    relationship_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                handle_relationship(
                    e,
                    true,
                    &mut writer,
                    &mut changed,
                    &mut has_vba_rel,
                    &mut max_rid,
                )?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                if relationship_tag_name.is_none() {
                    relationship_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                handle_relationship(
                    e,
                    false,
                    &mut writer,
                    &mut changed,
                    &mut has_vba_rel,
                    &mut max_rid,
                )?;
            }
            Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                if !has_vba_rel {
                    changed = true;
                    let relationship_tag_name = relationship_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Relationship"));
                    let next_rid = max_rid + 1;
                    let rel_id = format!("rId{next_rid}");
                    let mut rel = BytesStart::new(relationship_tag_name.as_str());
                    rel.push_attribute(("Id", rel_id.as_str()));
                    rel.push_attribute(("Type", VBA_REL_TYPE));
                    rel.push_attribute(("Target", VBA_TARGET));
                    writer.write_event(Event::Empty(rel))?;
                }
                writer.write_event(Event::End(e))?;
            }
            Event::Eof => break,
            other => writer.write_event(other)?,
        }

        buf.clear();
    }

    if changed {
        parts.insert(rels_key, writer.into_inner());
    }
    Ok(())
}

pub(crate) fn ensure_vba_project_rels_has_signature(
    parts: &mut BTreeMap<String, Vec<u8>>,
) -> Result<(), XlsxError> {
    if !(part_exists(parts, "xl/vbaProject.bin") && part_exists(parts, "xl/vbaProjectSignature.bin")) {
        return Ok(());
    }

    let rels_name = "xl/_rels/vbaProject.bin.rels";
    const REL_TYPE: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
    const TARGET: &str = "vbaProjectSignature.bin";

    let rels_key = select_part_key(parts, rels_name).unwrap_or_else(|| {
        // Prefer matching the leading-`/` style used by `vbaProject.bin` when synthesizing.
        if select_part_key(parts, "xl/vbaProject.bin")
            .as_deref()
            .is_some_and(|key| key.starts_with('/') || key.starts_with('\\'))
        {
            format!("/{rels_name}")
        } else {
            rels_name.to_string()
        }
    });

    match parts.get(&rels_key).cloned() {
        Some(existing) => {
            let mut reader = XmlReader::from_reader(existing.as_slice());
            reader.config_mut().trim_text(false);
            let mut out = Vec::new();
            let Some(cap) = existing.len().checked_add(128) else {
                return Err(XlsxError::AllocationFailure("ensure_vba_signature_relationship output"));
            };
            if out.try_reserve(cap).is_err() {
                return Err(XlsxError::AllocationFailure("ensure_vba_signature_relationship output"));
            }
            let mut writer = XmlWriter::new(out);
            let mut buf = Vec::new();

            let mut relationship_tag_name: Option<String> = None;

            let mut has_signature_rel = false;
            let mut changed = false;
            let mut max_rid = 0u32;

            fn handle_relationship(
                e: BytesStart<'_>,
                is_start: bool,
                writer: &mut XmlWriter<Vec<u8>>,
                changed: &mut bool,
                has_signature_rel: &mut bool,
                max_rid: &mut u32,
            ) -> Result<(), XlsxError> {
                let mut id = None;
                let mut type_uri = None;
                let mut target = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match local_name(attr.key.as_ref()) {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Type" => type_uri = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }

                if let Some(id) = id {
                    if let Some(n) = relationship_id_number(&id) {
                        *max_rid = (*max_rid).max(n);
                    }
                }

                let is_signature = type_uri
                    .as_deref()
                    .is_some_and(|t| t.trim() == REL_TYPE);
                if is_signature {
                    if target
                        .as_deref()
                        .is_some_and(|t| crate::zip_util::zip_part_names_equivalent(t.trim(), TARGET))
                    {
                        *has_signature_rel = true;
                        if is_start {
                            writer.write_event(Event::Start(e))?;
                        } else {
                            writer.write_event(Event::Empty(e))?;
                        }
                    } else {
                        // Relationship exists, but is missing/has an unexpected Target; patch it.
                        *has_signature_rel = true;
                        *changed = true;

                        let tag_name = e.name();
                        let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Relationship");
                        let mut patched = BytesStart::new(tag_name);
                        let mut saw_target = false;
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr?;
                            if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Target") {
                                saw_target = true;
                                patched.push_attribute((attr.key.as_ref(), TARGET.as_bytes()));
                            } else {
                                patched.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                            }
                        }
                        if !saw_target {
                            patched.push_attribute(("Target", TARGET));
                        }

                        if is_start {
                            writer.write_event(Event::Start(patched.into_owned()))?;
                        } else {
                            writer.write_event(Event::Empty(patched.into_owned()))?;
                        }
                    }
                } else if is_start {
                    writer.write_event(Event::Start(e))?;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }

                Ok(())
            }

            loop {
                let event = reader.read_event_into(&mut buf)?;
                match event {
                    Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                        writer.write_event(Event::Start(e))?;
                    }
                    Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                        // Degenerate case: self-closing `<Relationships/>` root; expand to insert
                        // the signature relationship.
                        changed = true;
                        let relationships_tag_name =
                            String::from_utf8_lossy(e.name().as_ref()).into_owned();
                        let relationship_tag_name = relationship_tag_name.clone().unwrap_or_else(|| {
                            prefixed_tag(relationships_tag_name.as_bytes(), "Relationship")
                        });

                        writer.write_event(Event::Start(e))?;

                        let next_rid = max_rid + 1;
                        let rel_id = format!("rId{next_rid}");
                        let mut rel = BytesStart::new(relationship_tag_name.as_str());
                        rel.push_attribute(("Id", rel_id.as_str()));
                        rel.push_attribute(("Type", REL_TYPE));
                        rel.push_attribute(("Target", TARGET));
                        writer.write_event(Event::Empty(rel))?;

                        writer.write_event(Event::End(BytesEnd::new(relationships_tag_name.as_str())))?;
                    }
                    Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                        if relationship_tag_name.is_none() {
                            relationship_tag_name =
                                Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                        }
                        handle_relationship(
                            e,
                            true,
                            &mut writer,
                            &mut changed,
                            &mut has_signature_rel,
                            &mut max_rid,
                        )?;
                    }
                    Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                        if relationship_tag_name.is_none() {
                            relationship_tag_name =
                                Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                        }
                        handle_relationship(
                            e,
                            false,
                            &mut writer,
                            &mut changed,
                            &mut has_signature_rel,
                            &mut max_rid,
                        )?;
                    }
                    Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                        if !has_signature_rel {
                            changed = true;
                            let relationship_tag_name = relationship_tag_name.clone().unwrap_or_else(|| {
                                prefixed_tag(e.name().as_ref(), "Relationship")
                            });
                            let next_rid = max_rid + 1;
                            let rel_id = format!("rId{next_rid}");
                            let mut rel = BytesStart::new(relationship_tag_name.as_str());
                            rel.push_attribute(("Id", rel_id.as_str()));
                            rel.push_attribute(("Type", REL_TYPE));
                            rel.push_attribute(("Target", TARGET));
                            writer.write_event(Event::Empty(rel))?;
                        }
                        writer.write_event(Event::End(e))?;
                    }
                    Event::Eof => break,
                    other => writer.write_event(other)?,
                }

                buf.clear();
            }
            if changed {
                parts.insert(rels_key, writer.into_inner());
            }
        }
        None => {
            // If the relationship part is missing, synthesize the minimal valid XML.
            // We keep the relationship ID deterministic to keep write output stable.
            let xml = format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{REL_TYPE}" Target="{TARGET}"/>
</Relationships>"#
            );
            parts.insert(rels_key, xml.into_bytes());
        }
    }

    Ok(())
}
