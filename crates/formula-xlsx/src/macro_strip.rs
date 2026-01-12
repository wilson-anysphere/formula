use std::collections::{BTreeMap, BTreeSet, VecDeque};

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};

use crate::package::XlsxError;

const CUSTOM_UI_REL_TYPES: [&str; 2] = [
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
];

const RELATIONSHIPS_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub(crate) fn strip_macros(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<(), XlsxError> {
    let delete_parts = compute_macro_delete_set(parts)?;

    for part in &delete_parts {
        parts.remove(part);
    }

    clean_relationship_parts(parts, &delete_parts)?;
    clean_content_types(parts, &delete_parts)?;

    Ok(())
}

fn compute_macro_delete_set(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<BTreeSet<String>, XlsxError> {
    let mut delete = BTreeSet::new();

    // VBA project payloads.
    delete.insert("xl/vbaProject.bin".to_string());
    delete.insert("xl/vbaData.xml".to_string());
    delete.insert("xl/vbaProjectSignature.bin".to_string());

    // Ribbon customizations.
    for name in parts.keys() {
        if name.starts_with("customUI/") {
            delete.insert(name.clone());
        }
    }

    // ActiveX + legacy form controls.
    for name in parts.keys() {
        if name.starts_with("xl/activeX/")
            || name.starts_with("xl/ctrlProps/")
            || name.starts_with("xl/controls/")
        {
            delete.insert(name.clone());
        }
    }

    // Legacy macro surfaces beyond VBA:
    // - Excel 4.0 macro sheets (XLM) stored under `xl/macrosheets/**`
    // - Dialog sheets stored under `xl/dialogsheets/**`
    for name in parts.keys() {
        if name.starts_with("xl/macrosheets/") || name.starts_with("xl/dialogsheets/") {
            delete.insert(name.clone());
        }
    }

    // Parts referenced by `xl/_rels/vbaProject.bin.rels` (e.g. signature payloads).
    if let Some(rels_bytes) = parts.get("xl/_rels/vbaProject.bin.rels") {
        let targets = parse_internal_relationship_targets(
            rels_bytes,
            "xl/vbaProject.bin",
            "xl/_rels/vbaProject.bin.rels",
            parts,
        )?;
        delete.extend(targets);
    }

    // ActiveX controls embedded into VML drawings can reference OLE/ActiveX binaries via
    // `xl/drawings/_rels/vmlDrawing*.vml.rels`. These VML parts are often shared with legacy
    // comments (ObjectType="Note"), so we cannot delete the whole VML drawing; instead we delete
    // the specific relationship targets used by `<o:OLEObject>` shapes so the cleanup pass can
    // remove only those shapes while preserving comments.
    delete.extend(find_vml_ole_object_targets(parts)?);

    // Build a relationship graph so we can delete any extra parts that are only
    // referenced by macro-related parts (e.g. `xl/embeddings/*` referenced by ActiveX rels).
    let graph = RelationshipGraph::build(parts)?;
    delete_orphan_targets(&graph, &mut delete);

    // If a part is deleted, its relationship part must also be deleted.
    //
    // (Skip `.rels` because relationship parts don't have relationship parts of their own.)
    let rels_to_remove: Vec<String> = delete
        .iter()
        .filter(|name| !name.ends_with(".rels"))
        .map(|name| crate::path::rels_for_part(name))
        .collect();
    delete.extend(rels_to_remove);

    Ok(delete)
}

fn find_vml_ole_object_targets(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<BTreeSet<String>, XlsxError> {
    let mut out = BTreeSet::new();

    for (vml_part, vml_bytes) in parts {
        if !vml_part.ends_with(".vml") {
            continue;
        }

        // Only VML drawings can contain `<o:OLEObject>` control shapes (commentsDrawing* parts are
        // DrawingML XML, not VML).
        if !vml_part.starts_with("xl/drawings/") {
            continue;
        }

        let rel_ids = parse_vml_ole_object_relationship_ids(vml_bytes)?;
        if rel_ids.is_empty() {
            continue;
        }

        let rels_part = crate::path::rels_for_part(vml_part);
        let Some(rels_bytes) = parts.get(&rels_part) else {
            continue;
        };

        out.extend(parse_relationship_targets_for_ids(
            rels_bytes, vml_part, &rel_ids,
        )?);
    }

    Ok(out)
}

fn parse_vml_ole_object_relationship_ids(xml: &[u8]) -> Result<BTreeSet<String>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut namespace_context = NamespaceContext::default();
    let mut ids = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"OLEObject")
                {
                    collect_relationship_id_attrs(e, &namespace_context, &mut ids)?;
                }
                namespace_context.push(changes);
            }
            Event::Empty(ref e) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"OLEObject")
                {
                    collect_relationship_id_attrs(e, &namespace_context, &mut ids)?;
                }
                namespace_context.rollback(changes);
            }
            Event::End(_) => namespace_context.pop(),
            _ => {}
        }
        buf.clear();
    }

    Ok(ids)
}

fn parse_relationship_targets_for_ids(
    xml: &[u8],
    source_part: &str,
    ids: &BTreeSet<String>,
) -> Result<BTreeSet<String>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match crate::openxml::local_name(attr.key.as_ref()) {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        b"TargetMode" => target_mode = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }

                if target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                let Some(id) = id else {
                    continue;
                };
                if !ids.contains(&id) {
                    continue;
                }

                let Some(target) = target else {
                    continue;
                };
                let target = strip_fragment(&target);
                let resolved = resolve_target_for_source(source_part, target);
                // Worksheet OLE objects are stored under `xl/embeddings/` and referenced from
                // `<oleObjects>` in sheet XML (valid in `.xlsx`). For macro stripping we only
                // delete embedding binaries referenced by VML `<o:OLEObject>` control shapes.
                if resolved.starts_with("xl/embeddings/") {
                    out.insert(resolved);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn delete_orphan_targets(graph: &RelationshipGraph, delete: &mut BTreeSet<String>) {
    let mut queue: VecDeque<String> = delete.iter().cloned().collect();
    while let Some(source) = queue.pop_front() {
        let Some(targets) = graph.outgoing.get(&source) else {
            continue;
        };
        for target in targets {
            if delete.contains(target) {
                continue;
            }
            let Some(inbound) = graph.inbound.get(target) else {
                continue;
            };
            // Only delete parts that are referenced *exclusively* by parts we're already deleting.
            if inbound.iter().all(|src| delete.contains(src)) {
                delete.insert(target.clone());
                queue.push_back(target.clone());
            }
        }
    }
}

fn clean_relationship_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    delete_parts: &BTreeSet<String>,
) -> Result<(), XlsxError> {
    let rels_names: Vec<String> = parts
        .keys()
        .filter(|name| name.ends_with(".rels"))
        .cloned()
        .collect();

    for rels_name in rels_names {
        let Some(source_part) = source_part_from_rels_part(&rels_name) else {
            continue;
        };

        // If the relationship source is gone, remove the `.rels` part as well.
        if !source_part.is_empty() && !parts.contains_key(&source_part) {
            parts.remove(&rels_name);
            continue;
        }

        let Some(bytes) = parts.get(&rels_name).cloned() else {
            continue;
        };

        let (updated, removed_ids) =
            strip_deleted_relationships(&rels_name, &source_part, &bytes, delete_parts, parts)?;

        if let Some(updated) = updated {
            parts.insert(rels_name.clone(), updated);
        }

        if !removed_ids.is_empty() {
            strip_source_relationship_references(parts, &source_part, &removed_ids)?;
        }
    }

    Ok(())
}

fn clean_content_types(
    parts: &mut BTreeMap<String, Vec<u8>>,
    delete_parts: &BTreeSet<String>,
) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let Some(existing) = parts.get(ct_name).cloned() else {
        return Ok(());
    };

    if let Some(updated) = strip_content_types(&existing, delete_parts)? {
        parts.insert(ct_name.to_string(), updated);
    }

    Ok(())
}

fn resolve_target_best_effort<F>(
    source_part: &str,
    rels_part: &str,
    target: &str,
    mut is_present: F,
) -> String
where
    F: FnMut(&str) -> bool,
{
    // OPC relationship targets are typically resolved relative to the source part's directory.
    // However, some producers appear to emit paths relative to the `.rels` directory instead
    // (e.g. `../media/*` from a workbook-level part). When the standard resolution doesn't match
    // an existing part, try alternative interpretations so macro stripping doesn't delete shared
    // parts that are still required elsewhere (for example by `xl/cellimages.xml`).
    let direct = resolve_target_for_source(source_part, target);
    if is_present(&direct) {
        return direct;
    }

    let rels_relative = crate::path::resolve_target(rels_part, target);
    if is_present(&rels_relative) {
        return rels_relative;
    }

    if !direct.starts_with("xl/") {
        let xl_prefixed = format!("xl/{direct}");
        if is_present(&xl_prefixed) {
            return xl_prefixed;
        }
    }

    direct
}

fn strip_deleted_relationships(
    rels_part_name: &str,
    source_part: &str,
    xml: &[u8],
    delete_parts: &BTreeSet<String>,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<(Option<Vec<u8>>, BTreeSet<String>), XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

    let mut buf = Vec::new();
    let mut changed = false;
    let mut removed_ids = BTreeSet::new();
    let mut skip_depth = 0usize;

    loop {
        let ev = reader.read_event_into(&mut buf)?;

        if skip_depth > 0 {
            match ev {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => {
                    skip_depth -= 1;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        match ev {
            Event::Eof => break,
            Event::Empty(e) if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" => {
                if should_remove_relationship(rels_part_name, source_part, &e, delete_parts, parts)?
                {
                    changed = true;
                    if let Some(id) = relationship_id(&e)? {
                        removed_ids.insert(id);
                    }
                    buf.clear();
                    continue;
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::Start(e) if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" => {
                if should_remove_relationship(rels_part_name, source_part, &e, delete_parts, parts)?
                {
                    changed = true;
                    if let Some(id) = relationship_id(&e)? {
                        removed_ids.insert(id);
                    }
                    skip_depth = 1;
                    buf.clear();
                    continue;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    let updated = if changed {
        Some(writer.into_inner())
    } else {
        None
    };

    Ok((updated, removed_ids))
}

fn should_remove_relationship(
    rels_part_name: &str,
    source_part: &str,
    e: &BytesStart<'_>,
    delete_parts: &BTreeSet<String>,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<bool, XlsxError> {
    let mut target = None;
    let mut target_mode = None;
    let mut rel_type = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match crate::openxml::local_name(attr.key.as_ref()) {
            b"Target" => target = Some(attr.unescape_value()?.into_owned()),
            b"TargetMode" => target_mode = Some(attr.unescape_value()?.into_owned()),
            b"Type" => rel_type = Some(attr.unescape_value()?.into_owned()),
            _ => {}
        }
    }

    if target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
    {
        return Ok(false);
    }

    if rels_part_name == "_rels/.rels"
        && rel_type
            .as_deref()
            .is_some_and(|ty| CUSTOM_UI_REL_TYPES.iter().any(|known| ty == *known))
    {
        return Ok(true);
    }

    let Some(target) = target else {
        return Ok(false);
    };

    let target = strip_fragment(&target);
    let resolved = resolve_target_best_effort(source_part, rels_part_name, target, |candidate| {
        parts.contains_key(candidate) || delete_parts.contains(candidate)
    });
    Ok(delete_parts.contains(&resolved))
}

fn relationship_id(e: &BytesStart<'_>) -> Result<Option<String>, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn strip_content_types(
    xml: &[u8],
    delete_parts: &BTreeSet<String>,
) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

    let mut buf = Vec::new();
    let mut changed = false;
    let mut skip_depth = 0usize;

    loop {
        let ev = reader.read_event_into(&mut buf)?;

        if skip_depth > 0 {
            match ev {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        match ev {
            Event::Eof => break,
            Event::Empty(e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                if let Some(updated) = patched_override(&e, delete_parts)? {
                    if updated.is_none() {
                        changed = true;
                        buf.clear();
                        continue;
                    }
                    if let Some(updated) = updated {
                        changed = true;
                        writer.write_event(Event::Empty(updated))?;
                        buf.clear();
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::Start(e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                // `<Override>` parts are expected to be empty, but handle the non-empty form just
                // in case by skipping the entire element when needed.
                if let Some(updated) = patched_override(&e, delete_parts)? {
                    if updated.is_none() {
                        changed = true;
                        skip_depth = 1;
                        buf.clear();
                        continue;
                    }
                    if let Some(updated) = updated {
                        changed = true;
                        writer.write_event(Event::Start(updated))?;
                        buf.clear();
                        continue;
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    if changed {
        Ok(Some(writer.into_inner()))
    } else {
        Ok(None)
    }
}

// Returns:
// - Ok(None) -> keep original
// - Ok(Some(None)) -> remove element
// - Ok(Some(Some(updated))) -> replace element
fn patched_override(
    e: &BytesStart<'_>,
    delete_parts: &BTreeSet<String>,
) -> Result<Option<Option<BytesStart<'static>>>, XlsxError> {
    let mut part_name = None;
    let mut content_type = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match crate::openxml::local_name(attr.key.as_ref()) {
            b"PartName" => part_name = Some(attr.unescape_value()?.into_owned()),
            b"ContentType" => content_type = Some(attr.unescape_value()?.into_owned()),
            _ => {}
        }
    }

    let Some(part_name) = part_name else {
        return Ok(None);
    };

    let normalized = part_name.strip_prefix('/').unwrap_or(part_name.as_str());
    if delete_parts.contains(normalized) {
        return Ok(Some(None));
    }

    let Some(content_type) = content_type else {
        return Ok(None);
    };

    if content_type.contains("macroEnabled.main+xml") {
        // Preserve the original element name (including any namespace prefix) so we don't produce
        // namespace-less `<Override>` elements when stripping macros from prefix-only
        // `[Content_Types].xml` documents (e.g. `<ct:Types xmlns:ct=\"...\">`).
        let tag_name = e.name();
        let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Override");
        let mut updated = BytesStart::new(tag_name);
        updated.push_attribute(("PartName", part_name.as_str()));
        updated.push_attribute((
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        ));
        return Ok(Some(Some(updated.into_owned())));
    }

    Ok(None)
}

fn parse_internal_relationship_targets(
    xml: &[u8],
    source_part: &str,
    rels_part: &str,
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<String>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref()) == b"Relationship" =>
            {
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    match crate::openxml::local_name(attr.key.as_ref()) {
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        b"TargetMode" => target_mode = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }

                if target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                let Some(target) = target else {
                    continue;
                };
                let target = strip_fragment(&target);
                out.push(resolve_target_best_effort(source_part, rels_part, target, |candidate| {
                    parts.contains_key(candidate)
                }));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

fn resolve_target_for_source(source_part: &str, target: &str) -> String {
    if source_part.is_empty() {
        crate::path::resolve_target("", target)
    } else {
        crate::path::resolve_target(source_part, target)
    }
}

fn strip_source_relationship_references(
    parts: &mut BTreeMap<String, Vec<u8>>,
    source_part: &str,
    removed_ids: &BTreeSet<String>,
) -> Result<(), XlsxError> {
    if source_part.is_empty()
        || !(source_part.ends_with(".xml") || source_part.ends_with(".vml"))
        || removed_ids.is_empty()
    {
        return Ok(());
    }

    let Some(xml) = parts.get(source_part).cloned() else {
        return Ok(());
    };

    if let Some(updated) = strip_relationship_id_references(&xml, removed_ids)? {
        parts.insert(source_part.to_string(), updated);
    }

    Ok(())
}

fn strip_relationship_id_references(
    xml: &[u8],
    removed_ids: &BTreeSet<String>,
) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));

    let mut buf = Vec::new();
    let mut changed = false;
    let mut skip_depth = 0usize;
    let mut namespace_context = NamespaceContext::default();

    loop {
        let ev = reader.read_event_into(&mut buf)?;

        if skip_depth > 0 {
            match ev {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        match ev {
            Event::Eof => break,
            Event::Start(e) => {
                let changes = namespace_context.apply_namespace_decls(&e)?;
                if element_has_removed_relationship_id(&e, &namespace_context, removed_ids)? {
                    changed = true;
                    namespace_context.rollback(changes);
                    skip_depth = 1;
                    buf.clear();
                    continue;
                }
                namespace_context.push(changes);
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(e) => {
                let changes = namespace_context.apply_namespace_decls(&e)?;
                if element_has_removed_relationship_id(&e, &namespace_context, removed_ids)? {
                    changed = true;
                    namespace_context.rollback(changes);
                    buf.clear();
                    continue;
                }
                namespace_context.rollback(changes);
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(e) => {
                namespace_context.pop();
                writer.write_event(Event::End(e.to_owned()))?;
            }
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    if changed {
        Ok(Some(writer.into_inner()))
    } else {
        Ok(None)
    }
}

fn element_has_removed_relationship_id(
    e: &BytesStart<'_>,
    namespace_context: &NamespaceContext,
    removed_ids: &BTreeSet<String>,
) -> Result<bool, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();

        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }

        let (prefix, local) = split_prefixed_name(key);
        let namespace_uri = prefix.and_then(|p| namespace_context.namespace_for_prefix(p));

        if !is_relationship_id_attribute(namespace_uri, local) {
            continue;
        }
        let value = attr.unescape_value()?;
        if removed_ids.contains(value.as_ref()) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn split_prefixed_name(name: &[u8]) -> (Option<&[u8]>, &[u8]) {
    match name.iter().position(|b| *b == b':') {
        Some(idx) => (Some(&name[..idx]), &name[idx + 1..]),
        None => (None, name),
    }
}

fn is_relationship_id_attribute(namespace_uri: Option<&[u8]>, local_name: &[u8]) -> bool {
    // Be defensive: VML/Office markup commonly uses `o:relid`, but some documents use other
    // prefixes or even no prefix at all. If the local-name is `relid` we treat it as a
    // relationship pointer regardless of namespace.
    if local_name.eq_ignore_ascii_case(b"relid") {
        return true;
    }

    match namespace_uri {
        Some(ns) if ns == RELATIONSHIPS_NS => {
            local_name.eq_ignore_ascii_case(b"id")
                || local_name.eq_ignore_ascii_case(b"embed")
                || local_name.eq_ignore_ascii_case(b"link")
        }
        _ => false,
    }
}

fn source_part_from_rels_part(rels_part: &str) -> Option<String> {
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
        return Some(rels_file.to_string());
    }

    Some(format!("{dir}/{rels_file}"))
}

struct RelationshipGraph {
    outgoing: BTreeMap<String, BTreeSet<String>>,
    inbound: BTreeMap<String, BTreeSet<String>>,
}

#[derive(Debug, Default)]
struct NamespaceContext {
    /// prefix -> namespace URI
    prefixes: BTreeMap<Vec<u8>, Vec<u8>>,
    /// Stack of prefix changes for each started element that was written.
    stack: Vec<Vec<(Vec<u8>, Option<Vec<u8>>)>>,
}

impl NamespaceContext {
    fn apply_namespace_decls(
        &mut self,
        e: &BytesStart<'_>,
    ) -> Result<Vec<(Vec<u8>, Option<Vec<u8>>)>, XlsxError> {
        let mut changes = Vec::new();

        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            let key = attr.key.as_ref();

            // Default namespace (`xmlns="..."`) affects element names, but not attributes.
            if key == b"xmlns" {
                continue;
            }

            let Some(prefix) = key.strip_prefix(b"xmlns:") else {
                continue;
            };

            let uri = attr.unescape_value()?.into_owned().into_bytes();
            let old = self.prefixes.insert(prefix.to_vec(), uri);
            changes.push((prefix.to_vec(), old));
        }

        Ok(changes)
    }

    fn rollback(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
        for (prefix, old) in changes.into_iter().rev() {
            match old {
                Some(uri) => {
                    self.prefixes.insert(prefix, uri);
                }
                None => {
                    self.prefixes.remove(&prefix);
                }
            }
        }
    }

    fn push(&mut self, changes: Vec<(Vec<u8>, Option<Vec<u8>>)>) {
        self.stack.push(changes);
    }

    fn pop(&mut self) {
        if let Some(changes) = self.stack.pop() {
            self.rollback(changes);
        }
    }

    fn namespace_for_prefix(&self, prefix: &[u8]) -> Option<&[u8]> {
        self.prefixes.get(prefix).map(Vec::as_slice)
    }
}

impl RelationshipGraph {
    fn build(parts: &BTreeMap<String, Vec<u8>>) -> Result<Self, XlsxError> {
        let mut outgoing: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
        let mut inbound: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

        for (rels_part, bytes) in parts {
            if !rels_part.ends_with(".rels") {
                continue;
            }

            let Some(source_part) = source_part_from_rels_part(rels_part) else {
                continue;
            };

            // Ignore orphan `.rels` parts; they'll be removed during cleanup.
            if !source_part.is_empty() && !parts.contains_key(&source_part) {
                continue;
            }

            let targets = parse_internal_relationship_targets(bytes, &source_part, rels_part, parts)?;
            for target in targets {
                outgoing
                    .entry(source_part.clone())
                    .or_default()
                    .insert(target.clone());
                inbound
                    .entry(target)
                    .or_default()
                    .insert(source_part.clone());
            }
        }

        Ok(Self { outgoing, inbound })
    }
}

pub fn validate_opc_relationships(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<(), XlsxError> {
    for rels_part in parts.keys().filter(|name| name.ends_with(".rels")) {
        let Some(source_part) = source_part_from_rels_part(rels_part) else {
            continue;
        };

        if !source_part.is_empty() && !parts.contains_key(&source_part) {
            return Err(XlsxError::Invalid(format!(
                "orphan relationship part {rels_part} (missing source {source_part})"
            )));
        }

        let xml = parts
            .get(rels_part)
            .ok_or_else(|| XlsxError::MissingPart(rels_part.to_string()))?;
        let ids = parse_relationship_ids(xml)?;
        let targets = parse_internal_relationship_targets(xml, &source_part, rels_part, parts)?;
        for target in targets {
            if !parts.contains_key(&target) {
                return Err(XlsxError::Invalid(format!(
                    "relationship target {target} referenced from {rels_part} is missing"
                )));
            }
        }

        if !source_part.is_empty()
            && (source_part.ends_with(".xml") || source_part.ends_with(".vml"))
            && parts.contains_key(&source_part)
        {
            let source_xml = parts
                .get(&source_part)
                .ok_or_else(|| XlsxError::MissingPart(source_part.clone()))?;
            let references = parse_relationship_id_references(source_xml)?;
            for id in references {
                if !ids.contains(&id) {
                    return Err(XlsxError::Invalid(format!(
                        "dangling relationship id {id} referenced from {source_part} (missing from {rels_part})"
                    )));
                }
            }
        }
    }

    Ok(())
}

fn parse_relationship_ids(xml: &[u8]) -> Result<BTreeSet<String>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = BTreeSet::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref())
                    .eq_ignore_ascii_case(b"Relationship") =>
            {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                        out.insert(attr.unescape_value()?.into_owned());
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_relationship_id_references(xml: &[u8]) -> Result<BTreeSet<String>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = BTreeSet::new();
    let mut namespace_context = NamespaceContext::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                collect_relationship_id_attrs(e, &namespace_context, &mut out)?;
                namespace_context.push(changes);
            }
            Event::Empty(ref e) => {
                let changes = namespace_context.apply_namespace_decls(e)?;
                collect_relationship_id_attrs(e, &namespace_context, &mut out)?;
                namespace_context.rollback(changes);
            }
            Event::End(_) => namespace_context.pop(),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn collect_relationship_id_attrs(
    e: &BytesStart<'_>,
    namespace_context: &NamespaceContext,
    out: &mut BTreeSet<String>,
) -> Result<(), XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();

        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }

        let (prefix, local) = split_prefixed_name(key);
        let namespace_uri = prefix.and_then(|p| namespace_context.namespace_for_prefix(p));
        if !is_relationship_id_attribute(namespace_uri, local) {
            continue;
        }
        out.insert(attr.unescape_value()?.into_owned());
    }

    Ok(())
}
