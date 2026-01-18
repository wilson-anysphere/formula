use std::collections::{BTreeMap, BTreeSet, VecDeque};

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};

use crate::package::XlsxError;
use crate::WorkbookKind;

type PartNameKey = Vec<u8>;

const CUSTOM_UI_REL_TYPES: [&str; 2] = [
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
];

const RELATIONSHIPS_NS: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn canonical_part_name(name: &str) -> &str {
    name.trim_start_matches(|c| c == '/' || c == '\\')
}

fn part_name_key(name: &str) -> Result<PartNameKey, XlsxError> {
    crate::zip_util::zip_part_name_lookup_key(name)
}

fn get_part<'a>(parts: &'a BTreeMap<String, Vec<u8>>, name: &str) -> Option<&'a [u8]> {
    parts
        .get(name)
        .map(Vec::as_slice)
        .or_else(|| {
            name.strip_prefix('/')
                .or_else(|| name.strip_prefix('\\'))
                .and_then(|name| parts.get(name).map(Vec::as_slice))
        })
        .or_else(|| {
            parts
                .iter()
                .find(|(key, _)| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
                .map(|(_, bytes)| bytes.as_slice())
        })
}

fn get_part_cloned_with_key(
    parts: &BTreeMap<String, Vec<u8>>,
    name: &str,
) -> Option<(String, Vec<u8>)> {
    if let Some(bytes) = parts.get(name) {
        return Some((name.to_string(), bytes.clone()));
    }

    if let Some(stripped) = name.strip_prefix('/').or_else(|| name.strip_prefix('\\')) {
        if let Some(bytes) = parts.get(stripped) {
            return Some((stripped.to_string(), bytes.clone()));
        }
    }

    parts
        .iter()
        .find(|(key, _)| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
        .map(|(key, bytes)| (key.clone(), bytes.clone()))
}

pub(crate) fn strip_macros(parts: &mut BTreeMap<String, Vec<u8>>) -> Result<(), XlsxError> {
    strip_macros_with_kind(parts, WorkbookKind::Workbook)
}

pub(crate) fn strip_macros_with_kind(
    parts: &mut BTreeMap<String, Vec<u8>>,
    target_kind: WorkbookKind,
) -> Result<(), XlsxError> {
    let mut present_parts: BTreeSet<PartNameKey> = BTreeSet::new();
    for name in parts.keys() {
        present_parts.insert(crate::zip_util::zip_part_name_lookup_key(name)?);
    }

    let delete_parts = compute_macro_delete_set(parts, &present_parts)?;

    // Delete any ZIP part whose normalized key is in the delete set. This removes macro surfaces
    // even when a producer used non-canonical naming (case differences, `\` separators, leading
    // separators, percent-encoding).
    let mut to_remove: Vec<String> = Vec::new();
    for name in parts.keys() {
        let key = crate::zip_util::zip_part_name_lookup_key(name)?;
        if delete_parts.contains(&key) {
            to_remove.push(name.clone());
        }
    }
    for name in to_remove {
        parts.remove(&name);
    }

    clean_relationship_parts(parts, &delete_parts, &present_parts)?;
    clean_content_types(parts, &delete_parts, target_kind)?;

    Ok(())
}

fn compute_macro_delete_set(
    parts: &BTreeMap<String, Vec<u8>>,
    present_parts: &BTreeSet<PartNameKey>,
) -> Result<BTreeSet<PartNameKey>, XlsxError> {
    let mut delete: BTreeSet<PartNameKey> = BTreeSet::new();

    // VBA project payloads.
    delete.insert(part_name_key("xl/vbaProject.bin")?);
    delete.insert(part_name_key("xl/vbaData.xml")?);
    delete.insert(part_name_key("xl/vbaProjectSignature.bin")?);

    // Ribbon customizations.
    for name in parts.keys() {
        let key = part_name_key(name)?;
        if key.starts_with(b"customui/") {
            delete.insert(key);
        }
    }

    // ActiveX + legacy form controls.
    for name in parts.keys() {
        let key = part_name_key(name)?;
        if key.starts_with(b"xl/activex/")
            || key.starts_with(b"xl/ctrlprops/")
            || key.starts_with(b"xl/controls/")
        {
            delete.insert(key);
        }
    }

    // Legacy macro surfaces beyond VBA:
    // - Excel 4.0 macro sheets (XLM) stored under `xl/macrosheets/**`
    // - Dialog sheets stored under `xl/dialogsheets/**`
    for name in parts.keys() {
        let key = part_name_key(name)?;
        if key.starts_with(b"xl/macrosheets/") || key.starts_with(b"xl/dialogsheets/") {
            delete.insert(key);
        }
    }

    // Parts referenced by `xl/_rels/vbaProject.bin.rels` (e.g. signature payloads).
    if let Some((_rels_part, rels_bytes)) =
        get_part_cloned_with_key(parts, "xl/_rels/vbaProject.bin.rels")
    {
        let targets = parse_internal_relationship_targets(
            &rels_bytes,
            "xl/vbaProject.bin",
            "xl/_rels/vbaProject.bin.rels",
            present_parts,
            Some(&delete),
        )?;
        delete.extend(targets.into_iter());
    }

    // ActiveX controls embedded into VML drawings can reference OLE/ActiveX binaries via
    // `xl/drawings/_rels/vmlDrawing*.vml.rels`. These VML parts are often shared with legacy
    // comments (ObjectType="Note"), so we cannot delete the whole VML drawing; instead we delete
    // the specific relationship targets used by `<o:OLEObject>` shapes so the cleanup pass can
    // remove only those shapes while preserving comments.
    delete.extend(find_vml_ole_object_targets(parts)?);

    // Build a relationship graph so we can delete any extra parts that are only
    // referenced by macro-related parts (e.g. `xl/embeddings/*` referenced by ActiveX rels).
    let graph = RelationshipGraph::build(parts, present_parts)?;
    delete_orphan_targets(&graph, &mut delete)?;

    // If a part is deleted, its relationship part must also be deleted.
    //
    // (Skip `.rels` because relationship parts don't have relationship parts of their own.)
    let rels_to_remove: Vec<PartNameKey> = delete
        .iter()
        .filter(|name| !name.ends_with(b".rels"))
        .map(|name| rels_for_part_key(name))
        .collect::<Result<_, _>>()?;
    delete.extend(rels_to_remove);

    Ok(delete)
}

fn rels_for_part_key(part: &[u8]) -> Result<PartNameKey, XlsxError> {
    match part.iter().rposition(|b| *b == b'/') {
        Some(idx) => {
            let dir = &part[..idx];
            let file = &part[idx + 1..];
            let mut out = Vec::new();
            out.try_reserve(dir.len() + b"/_rels/".len() + file.len() + b".rels".len())
                .map_err(|_| XlsxError::AllocationFailure("rels_for_part_key"))?;
            out.extend_from_slice(dir);
            out.extend_from_slice(b"/_rels/");
            out.extend_from_slice(file);
            out.extend_from_slice(b".rels");
            Ok(out)
        }
        None => {
            let mut out = Vec::new();
            out.try_reserve(b"_rels/".len() + part.len() + b".rels".len())
                .map_err(|_| XlsxError::AllocationFailure("rels_for_part_key"))?;
            out.extend_from_slice(b"_rels/");
            out.extend_from_slice(part);
            out.extend_from_slice(b".rels");
            Ok(out)
        }
    }
}

fn find_vml_ole_object_targets(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<BTreeSet<PartNameKey>, XlsxError> {
    let mut out: BTreeSet<PartNameKey> = BTreeSet::new();

    for (vml_part, vml_bytes) in parts {
        let vml_key = part_name_key(vml_part)?;
        if !vml_key.ends_with(b".vml") {
            continue;
        }

        // Only VML drawings can contain `<o:OLEObject>` control shapes (commentsDrawing* parts are
        // DrawingML XML, not VML).
        if !vml_key.starts_with(b"xl/drawings/") {
            continue;
        }

        let rel_ids = parse_vml_ole_object_relationship_ids(vml_bytes)?;
        if rel_ids.is_empty() {
            continue;
        }

        let rels_key = rels_for_part_key(&vml_key)?;
        let mut rels_bytes = None;
        for (name, bytes) in parts {
            if part_name_key(name)? == rels_key {
                rels_bytes = Some(bytes.clone());
                break;
            }
        }
        let Some(rels_bytes) = rels_bytes else {
            continue;
        };

        out.extend(parse_relationship_targets_for_ids(
            &rels_bytes,
            vml_part,
            &rel_ids,
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
) -> Result<BTreeSet<PartNameKey>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out: BTreeSet<PartNameKey> = BTreeSet::new();

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
                let resolved_key = part_name_key(&resolved)?;
                // Worksheet OLE objects are stored under `xl/embeddings/` and referenced from
                // `<oleObjects>` in sheet XML (valid in `.xlsx`). For macro stripping we only
                // delete embedding binaries referenced by VML `<o:OLEObject>` control shapes.
                if resolved_key.starts_with(b"xl/embeddings/") {
                    out.insert(resolved_key);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn delete_orphan_targets(
    graph: &RelationshipGraph,
    delete: &mut BTreeSet<PartNameKey>,
) -> Result<(), XlsxError> {
    let mut queue: VecDeque<PartNameKey> = VecDeque::new();
    queue
        .try_reserve(delete.len())
        .map_err(|_| XlsxError::AllocationFailure("delete_orphan_targets queue"))?;
    queue.extend(delete.iter().cloned());
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
    Ok(())
}

fn clean_relationship_parts(
    parts: &mut BTreeMap<String, Vec<u8>>,
    delete_parts: &BTreeSet<PartNameKey>,
    present_parts: &BTreeSet<PartNameKey>,
) -> Result<(), XlsxError> {
    let mut remaining_parts: BTreeSet<PartNameKey> = BTreeSet::new();
    for name in parts.keys() {
        remaining_parts.insert(part_name_key(name)?);
    }

    let mut rels_names: Vec<String> = Vec::new();
    for name in parts.keys() {
        if part_name_key(name)?.ends_with(b".rels") {
            rels_names.push(name.clone());
        }
    }

    for rels_name in rels_names {
        let Some(source_part) = source_part_from_rels_part(&rels_name) else {
            continue;
        };
        let source_part = canonical_part_name(&source_part).to_string();
        let source_key = part_name_key(&source_part)?;

        // If the relationship source is gone, remove the `.rels` part as well.
        if !source_part.is_empty() && !remaining_parts.contains(&source_key) {
            parts.remove(&rels_name);
            continue;
        }

        let Some(bytes) = parts.get(&rels_name).cloned() else {
            continue;
        };

        let (updated, removed_ids) = strip_deleted_relationships(
            &rels_name,
            &source_part,
            &bytes,
            delete_parts,
            present_parts,
        )?;

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
    delete_parts: &BTreeSet<PartNameKey>,
    target_kind: WorkbookKind,
) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let Some((ct_key, existing)) = get_part_cloned_with_key(parts, ct_name) else {
        return Ok(());
    };

    if let Some(updated) = strip_content_types(&existing, delete_parts, target_kind)? {
        parts.insert(ct_key, updated);
    }

    Ok(())
}

fn resolve_target_best_effort_key(
    source_part: &str,
    rels_part: &str,
    target: &str,
    present_parts: &BTreeSet<PartNameKey>,
    delete_parts: Option<&BTreeSet<PartNameKey>>,
) -> Result<PartNameKey, XlsxError> {
    // OPC relationship targets are typically resolved relative to the source part's directory.
    // However, some producers appear to emit paths relative to the `.rels` directory instead
    // (e.g. `../media/*` from a workbook-level part). When the standard resolution doesn't match
    // an existing part, try alternative interpretations so macro stripping doesn't delete shared
    // parts that are still required elsewhere (for example by `xl/cellimages.xml`).
    let direct = resolve_target_for_source(source_part, target);
    let direct_key = part_name_key(&direct)?;
    if present_parts.contains(&direct_key) || delete_parts.is_some_and(|d| d.contains(&direct_key)) {
        return Ok(direct_key);
    }

    let rels_relative = crate::path::resolve_target(rels_part, target);
    let rels_relative_key = part_name_key(&rels_relative)?;
    if present_parts.contains(&rels_relative_key)
        || delete_parts.is_some_and(|d| d.contains(&rels_relative_key))
    {
        return Ok(rels_relative_key);
    }

    if !direct_key.starts_with(b"xl/") {
        let xl_prefixed = format!("xl/{direct}");
        let xl_prefixed_key = part_name_key(&xl_prefixed)?;
        if present_parts.contains(&xl_prefixed_key)
            || delete_parts.is_some_and(|d| d.contains(&xl_prefixed_key))
        {
            return Ok(xl_prefixed_key);
        }
    }

    Ok(direct_key)
}

fn strip_deleted_relationships(
    rels_part_name: &str,
    source_part: &str,
    xml: &[u8],
    delete_parts: &BTreeSet<PartNameKey>,
    present_parts: &BTreeSet<PartNameKey>,
) -> Result<(Option<Vec<u8>>, BTreeSet<String>), XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(xml.len())
        .map_err(|_| XlsxError::AllocationFailure("strip_deleted_relationships output"))?;
    let mut writer = XmlWriter::new(out);

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
                if should_remove_relationship(
                    rels_part_name,
                    source_part,
                    &e,
                    delete_parts,
                    present_parts,
                )? {
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
                if should_remove_relationship(
                    rels_part_name,
                    source_part,
                    &e,
                    delete_parts,
                    present_parts,
                )? {
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
    delete_parts: &BTreeSet<PartNameKey>,
    present_parts: &BTreeSet<PartNameKey>,
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

    if part_name_key(rels_part_name)? == b"_rels/.rels"
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
    let resolved_key = resolve_target_best_effort_key(
        source_part,
        rels_part_name,
        target,
        present_parts,
        Some(delete_parts),
    )?;
    Ok(delete_parts.contains(&resolved_key))
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
    delete_parts: &BTreeSet<PartNameKey>,
    target_kind: WorkbookKind,
) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(xml.len())
        .map_err(|_| XlsxError::AllocationFailure("strip_content_types output"))?;
    let mut writer = XmlWriter::new(out);

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
                if let Some(updated) = patched_override(&e, delete_parts, target_kind)? {
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
                if let Some(updated) = patched_override(&e, delete_parts, target_kind)? {
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
    delete_parts: &BTreeSet<PartNameKey>,
    target_kind: WorkbookKind,
) -> Result<Option<Option<BytesStart<'static>>>, XlsxError> {
    let mut part_name = None;
    let mut content_type = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match crate::openxml::local_name(attr.key.as_ref()) {
            key if key.eq_ignore_ascii_case(b"PartName") => {
                part_name = Some(attr.unescape_value()?.into_owned())
            }
            key if key.eq_ignore_ascii_case(b"ContentType") => {
                content_type = Some(attr.unescape_value()?.into_owned())
            }
            _ => {}
        }
    }

    let Some(part_name) = part_name else {
        return Ok(None);
    };

    let key = part_name_key(part_name.as_str())?;
    if delete_parts.contains(&key) {
        return Ok(Some(None));
    }

    if content_type
        .as_deref()
        .is_some_and(|ty| ty.contains("macroEnabled.main+xml"))
    {
        let workbook_main_type = target_kind
            .macro_free_kind()
            .workbook_content_type();

        // Preserve the original element's qualified name (including any namespace prefix).
        let tag_name = e.name();
        let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Override");
        let mut updated = BytesStart::new(tag_name);

        // Preserve all attributes verbatim (including any prefixes/ordering), except for
        // `ContentType`, which is rewritten to the non-macro workbook content type.
        let mut saw_content_type = false;
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"ContentType") {
                saw_content_type = true;
                updated.push_attribute((attr.key.as_ref(), workbook_main_type.as_bytes()));
            } else {
                updated.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
            }
        }
        if !saw_content_type {
            updated.push_attribute(("ContentType", workbook_main_type));
        }
        return Ok(Some(Some(updated.into_owned())));
    }

    Ok(None)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeSet;

    #[test]
    fn strip_content_types_preserves_prefix_only_override_qname() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml" CustomAttr="yes"/>
  <ct:Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</ct:Types>"#;

        let mut delete_parts = BTreeSet::new();
        delete_parts.insert(part_name_key("xl/vbaProject.bin").unwrap());

        let updated = strip_content_types(xml, &delete_parts, WorkbookKind::Workbook)
            .expect("strip_content_types ok")
            .expect("expected content types update");

        let updated_str = String::from_utf8(updated).expect("utf8 xml");
        roxmltree::Document::parse(&updated_str).expect("valid xml");

        // Ensure we didn't introduce an unprefixed `<Override>` tag.
        assert!(updated_str.contains("<ct:Override"));
        assert!(!updated_str.contains("<Override"));

        // Ensure we downgraded the workbook override.
        assert!(updated_str.contains(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        ));
        assert!(!updated_str.contains("macroEnabled.main+xml"));

        // Ensure extra attributes are preserved.
        assert!(updated_str.contains(r#"CustomAttr="yes""#));

        // Ensure deleted part overrides are removed.
        assert!(!updated_str.contains("vbaProject.bin"));
    }

    #[test]
    fn strip_content_types_handles_non_empty_override_with_prefix() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"></ct:Override>
        </ct:Types>"#;

        let delete_parts = BTreeSet::new();
        let updated = strip_content_types(xml, &delete_parts, WorkbookKind::Workbook)
            .expect("strip_content_types ok")
            .expect("expected content types update");

        let updated_str = String::from_utf8(updated).expect("utf8 xml");
        roxmltree::Document::parse(&updated_str).expect("valid xml");

        // Ensure we didn't introduce mismatched `<Override>` / `</ct:Override>` tags.
        assert!(updated_str.contains("<ct:Override"));
        assert!(updated_str.contains("</ct:Override>"));
        assert!(!updated_str.contains("<Override"));
    }
}

fn parse_internal_relationship_targets(
    xml: &[u8],
    source_part: &str,
    rels_part: &str,
    present_parts: &BTreeSet<PartNameKey>,
    delete_parts: Option<&BTreeSet<PartNameKey>>,
) -> Result<Vec<PartNameKey>, XlsxError> {
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
                out.push(resolve_target_best_effort_key(
                    source_part,
                    rels_part,
                    target,
                    present_parts,
                    delete_parts,
                )?);
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
    let source_key = part_name_key(source_part)?;
    if source_part.is_empty()
        || !(source_key.ends_with(b".xml") || source_key.ends_with(b".vml"))
        || removed_ids.is_empty()
    {
        return Ok(());
    }

    let Some((source_key, xml)) = get_part_cloned_with_key(parts, source_part) else {
        return Ok(());
    };

    if let Some(updated) = strip_relationship_id_references(&xml, removed_ids)? {
        parts.insert(source_key, updated);
    }

    Ok(())
}

fn strip_relationship_id_references(
    xml: &[u8],
    removed_ids: &BTreeSet<String>,
) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    out.try_reserve(xml.len())
        .map_err(|_| XlsxError::AllocationFailure("strip_relationship_id_references output"))?;
    let mut writer = XmlWriter::new(out);

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
    // Normalize separators and strip leading separators for robust parsing.
    let rels_part = rels_part.trim_start_matches(|c| c == '/' || c == '\\');
    let rels_part = if rels_part.contains('\\') {
        std::borrow::Cow::Owned(rels_part.replace('\\', "/"))
    } else {
        std::borrow::Cow::Borrowed(rels_part)
    };
    let rels_part = rels_part.as_ref();

    if rels_part.eq_ignore_ascii_case("_rels/.rels") {
        return Some(String::new());
    }

    if crate::ascii::starts_with_ignore_case(rels_part, "_rels/") {
        let rest = &rels_part["_rels/".len()..];
        if rest.len() < ".rels".len() || !crate::ascii::ends_with_ignore_case(rest, ".rels") {
            return None;
        }
        let base_len = rest.len() - ".rels".len();
        return Some(rest[..base_len].to_string());
    }

    let marker = "/_rels/";
    let idx = crate::ascii::rfind_ignore_case(rels_part, marker)?;
    let dir = &rels_part[..idx];
    let rels_file = &rels_part[idx + marker.len()..];
    if rels_file.len() < ".rels".len() || !crate::ascii::ends_with_ignore_case(rels_file, ".rels") {
        return None;
    }
    let base_len = rels_file.len() - ".rels".len();
    let base = &rels_file[..base_len];

    if dir.is_empty() {
        return Some(base.to_string());
    }

    Some(format!("{dir}/{base}"))
}

struct RelationshipGraph {
    outgoing: BTreeMap<PartNameKey, BTreeSet<PartNameKey>>,
    inbound: BTreeMap<PartNameKey, BTreeSet<PartNameKey>>,
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
    fn build(
        parts: &BTreeMap<String, Vec<u8>>,
        present_parts: &BTreeSet<PartNameKey>,
    ) -> Result<Self, XlsxError> {
        let mut outgoing: BTreeMap<PartNameKey, BTreeSet<PartNameKey>> = BTreeMap::new();
        let mut inbound: BTreeMap<PartNameKey, BTreeSet<PartNameKey>> = BTreeMap::new();

        for (rels_part, bytes) in parts {
            let rels_part_key = part_name_key(rels_part)?;
            if !rels_part_key.ends_with(b".rels") {
                continue;
            }

            let Some(source_part) = source_part_from_rels_part(rels_part) else {
                continue;
            };
            let source_part = canonical_part_name(&source_part).to_string();
            let source_key = part_name_key(&source_part)?;

            // Ignore orphan `.rels` parts; they'll be removed during cleanup.
            if !source_part.is_empty() && !present_parts.contains(&source_key) {
                continue;
            }

            let targets = parse_internal_relationship_targets(
                bytes,
                &source_part,
                rels_part,
                present_parts,
                None,
            )?;
            for target in targets {
                outgoing
                    .entry(source_key.clone())
                    .or_default()
                    .insert(target.clone());
                inbound.entry(target).or_default().insert(source_key.clone());
            }
        }

        Ok(Self { outgoing, inbound })
    }
}

pub fn validate_opc_relationships(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<(), XlsxError> {
    let mut present_parts: BTreeSet<PartNameKey> = BTreeSet::new();
    for name in parts.keys() {
        present_parts.insert(part_name_key(name)?);
    }

    let mut rels_parts: Vec<String> = Vec::new();
    for name in parts.keys() {
        if part_name_key(name)?.ends_with(b".rels") {
            rels_parts.push(name.clone());
        }
    }

    for rels_part in rels_parts {
        let Some(source_part) = source_part_from_rels_part(&rels_part) else {
            continue;
        };
        let source_part = canonical_part_name(&source_part).to_string();
        let source_key = part_name_key(&source_part)?;

        if !source_part.is_empty() && !present_parts.contains(&source_key) {
            return Err(XlsxError::Invalid(format!(
                "orphan relationship part {rels_part} (missing source {source_part})"
            )));
        }

        let xml = parts
            .get(&rels_part)
            .ok_or_else(|| XlsxError::MissingPart(rels_part.to_string()))?;
        let ids = parse_relationship_ids(xml)?;
        let targets = parse_internal_relationship_targets(
            xml,
            &source_part,
            &rels_part,
            &present_parts,
            None,
        )?;
        for target_key in targets {
            if !present_parts.contains(&target_key) {
                let target = String::from_utf8_lossy(&target_key);
                return Err(XlsxError::Invalid(format!(
                    "relationship target {target} referenced from {rels_part} is missing"
                )));
            }
        }

        if !source_part.is_empty()
            && (source_key.ends_with(b".xml") || source_key.ends_with(b".vml"))
            && present_parts.contains(&source_key)
        {
            let source_xml = get_part(parts, &source_part)
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
