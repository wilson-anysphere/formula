use formula_model::{SheetVisibility, TabColor};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use std::collections::{HashMap, HashSet};

use crate::XlsxError;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookSheetInfo {
    pub name: String,
    pub sheet_id: u32,
    pub rel_id: String,
    pub visibility: SheetVisibility,
}

pub fn parse_workbook_sheets(workbook_xml: &str) -> Result<Vec<WorkbookSheetInfo>, XlsxError> {
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) | Event::Start(e) => {
                if e.local_name().as_ref() == b"sheet" {
                    sheets.push(parse_sheet_element(&e)?);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(sheets)
}

fn parse_sheet_element(e: &BytesStart<'_>) -> Result<WorkbookSheetInfo, XlsxError> {
    let mut name: Option<String> = None;
    let mut sheet_id: Option<u32> = None;
    let mut rel_id: Option<String> = None;
    let mut visibility = SheetVisibility::Visible;

    for attr in e.attributes() {
        let attr = attr?;
        let key = attr.key.as_ref();
        match key {
            b"name" => name = Some(attr.unescape_value()?.to_string()),
            b"sheetId" => {
                let v = attr.unescape_value()?;
                sheet_id = Some(v.parse::<u32>().map_err(|_| XlsxError::InvalidSheetId)?);
            }
            b"state" => {
                let v = attr.unescape_value()?;
                visibility = match v.as_ref() {
                    "hidden" => SheetVisibility::Hidden,
                    "veryHidden" => SheetVisibility::VeryHidden,
                    _ => SheetVisibility::Visible,
                };
            }
            _ if crate::openxml::local_name(key) == b"id" => {
                rel_id = Some(attr.unescape_value()?.to_string())
            }
            _ => {}
        }
    }

    Ok(WorkbookSheetInfo {
        name: name.ok_or(XlsxError::MissingAttr("name"))?,
        sheet_id: sheet_id.ok_or(XlsxError::MissingAttr("sheetId"))?,
        rel_id: rel_id.ok_or(XlsxError::MissingAttr("r:id"))?,
        visibility,
    })
}

pub fn write_workbook_sheets(
    workbook_xml: &str,
    sheets: &[WorkbookSheetInfo],
) -> Result<String, XlsxError> {
    let (sheet_tag, rel_id_attr) = detect_workbook_sheet_qnames(workbook_xml)?;
    let sheet_tag = sheet_tag.unwrap_or_else(|| "sheet".to_string());
    // Avoid emitting undeclared namespace prefixes. We will attempt to infer the correct
    // relationships prefix from the workbook root / sheets element; if one can't be found,
    // fall back to an unprefixed `id` attribute.
    let rel_id_attr = rel_id_attr.unwrap_or_else(|| "id".to_string());

    let preserved_attrs = collect_preserved_sheet_attributes(workbook_xml)?;

    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut in_sheets = false;
    let mut skip_depth: usize = 0;
    let mut replaced_sheets = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.local_name().as_ref() == b"sheets" => {
                in_sheets = true;
                replaced_sheets = true;
                writer.write_event(Event::Start(e.to_owned()))?;
                for sheet in sheets {
                    writer.write_event(Event::Empty(build_sheet_element(
                        sheet_tag.as_str(),
                        rel_id_attr.as_str(),
                        sheet,
                        preserved_attrs.for_sheet(sheet),
                    )))?;
                }
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheets" => {
                // Expand a self-closing `<sheets/>` element so that we can insert `<sheet/>`
                // children.
                replaced_sheets = true;
                let sheets_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(Event::Start(e.to_owned()))?;
                for sheet in sheets {
                    writer.write_event(Event::Empty(build_sheet_element(
                        sheet_tag.as_str(),
                        rel_id_attr.as_str(),
                        sheet,
                        preserved_attrs.for_sheet(sheet),
                    )))?;
                }
                writer.write_event(Event::End(BytesEnd::new(sheets_tag.as_str())))?;
            }
            Event::Empty(ref e) if in_sheets && e.local_name().as_ref() == b"sheet" => {}
            Event::Start(ref e) if in_sheets && e.local_name().as_ref() == b"sheet" => {
                skip_depth = 1;
            }
            Event::End(ref e) if in_sheets && skip_depth > 0 => {
                if e.local_name().as_ref() == b"sheet" {
                    skip_depth = 0;
                }
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheets" => {
                in_sheets = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                if skip_depth == 0 {
                    writer.write_event(event.to_owned())?;
                }
            }
        }
        buf.clear();
    }

    if !replaced_sheets {
        return Ok(workbook_xml.to_string());
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn build_sheet_element(
    tag: &str,
    rel_id_attr: &str,
    sheet: &WorkbookSheetInfo,
    preserved_attrs: Option<&[PreservedSheetAttribute]>,
) -> BytesStart<'static> {
    let mut elem = BytesStart::new(tag).into_owned();
    elem.push_attribute(("name", sheet.name.as_str()));
    elem.push_attribute(("sheetId", sheet.sheet_id.to_string().as_str()));
    elem.push_attribute((rel_id_attr, sheet.rel_id.as_str()));
    match sheet.visibility {
        SheetVisibility::Visible => {}
        SheetVisibility::Hidden => elem.push_attribute(("state", "hidden")),
        SheetVisibility::VeryHidden => elem.push_attribute(("state", "veryHidden")),
    }
    if let Some(attrs) = preserved_attrs {
        for attr in attrs {
            elem.push_attribute((attr.key.as_str(), attr.value.as_str()));
        }
    }
    elem
}

#[derive(Debug, Clone)]
struct PreservedSheetAttribute {
    key: String,
    value: String,
}

#[derive(Debug, Default)]
struct PreservedSheetAttributes {
    /// Map of relationship id (`*:id` by local-name) => preserved attributes.
    by_rel_id: HashMap<String, Vec<PreservedSheetAttribute>>,
    /// Map of `sheetId` => preserved attributes (fallback when rel id doesn't match).
    by_sheet_id: HashMap<u32, Vec<PreservedSheetAttribute>>,
}

impl PreservedSheetAttributes {
    fn for_sheet(&self, sheet: &WorkbookSheetInfo) -> Option<&[PreservedSheetAttribute]> {
        self.by_rel_id
            .get(&sheet.rel_id)
            .map(|v| v.as_slice())
            .or_else(|| self.by_sheet_id.get(&sheet.sheet_id).map(|v| v.as_slice()))
    }
}

fn collect_xmlns_prefixes(e: &BytesStart<'_>, out: &mut HashSet<String>) -> Result<(), XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        out.insert(String::from_utf8_lossy(prefix).into_owned());
    }
    Ok(())
}

fn collect_preserved_sheet_attributes(workbook_xml: &str) -> Result<PreservedSheetAttributes, XlsxError> {
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_sheets = false;
    let mut prefixes_in_scope: HashSet<String> = HashSet::new();
    // The `xml` prefix is implicitly declared by the XML Namespaces spec.
    prefixes_in_scope.insert("xml".to_string());

    let mut out = PreservedSheetAttributes::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                collect_xmlns_prefixes(e, &mut prefixes_in_scope)?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheets" => {
                collect_xmlns_prefixes(e, &mut prefixes_in_scope)?;
                in_sheets = true;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheets" => {
                // Self-closing `<sheets/>` has no children, but its namespace declarations are
                // still in-scope for any `<sheet>` elements we insert later.
                collect_xmlns_prefixes(e, &mut prefixes_in_scope)?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheets" => {
                in_sheets = false;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_sheets && e.local_name().as_ref() == b"sheet" =>
            {
                let mut rel_id: Option<String> = None;
                let mut sheet_id: Option<u32> = None;
                let mut preserved: Vec<PreservedSheetAttribute> = Vec::new();

                for attr in e.attributes() {
                    let attr = attr?;
                    let key = attr.key.as_ref();

                    // Record the identifiers used for matching.
                    if key == b"sheetId" {
                        let v = attr.unescape_value()?;
                        sheet_id = Some(v.parse::<u32>().map_err(|_| XlsxError::InvalidSheetId)?);
                        continue;
                    }
                    if crate::openxml::local_name(key) == b"id" {
                        rel_id = Some(attr.unescape_value()?.to_string());
                        continue;
                    }

                    // Managed attributes are always overwritten by the caller-provided sheet list.
                    if matches!(key, b"name" | b"state") {
                        continue;
                    }

                    // Never carry over namespace declarations from the original `<sheet>` element.
                    // We replace `<sheet>` nodes wholesale, so any sheet-scoped declarations would
                    // be dropped unless we re-emit them. For safety we intentionally *do not*
                    // re-emit sheet-scoped `xmlns:*` attributes, and we will also drop any unknown
                    // attributes that rely on a prefix which isn't declared on the workbook root
                    // or `<sheets>` element. This avoids emitting invalid XML with undeclared
                    // prefixes, at the cost of losing those attributes.
                    if key == b"xmlns" || key.starts_with(b"xmlns:") {
                        continue;
                    }

                    let key_str = match std::str::from_utf8(key) {
                        Ok(s) => s,
                        Err(_) => continue,
                    };

                    if let Some((prefix, _)) = key_str.split_once(':') {
                        if !prefixes_in_scope.contains(prefix) {
                            continue;
                        }
                    }

                    preserved.push(PreservedSheetAttribute {
                        key: key_str.to_string(),
                        value: attr.unescape_value()?.to_string(),
                    });
                }

                if preserved.is_empty() {
                    continue;
                }

                if let Some(rel_id) = rel_id.as_ref() {
                    out.by_rel_id.insert(rel_id.clone(), preserved.clone());
                }
                if let Some(sheet_id) = sheet_id {
                    out.by_sheet_id.insert(sheet_id, preserved);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn detect_workbook_sheet_qnames(
    workbook_xml: &str,
) -> Result<(Option<String>, Option<String>), XlsxError> {
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut sheets_prefix: Option<String> = None;
    let mut sheet_tag: Option<String> = None;
    let mut rel_id_attr: Option<String> = None;
    let mut office_relationships_prefixes_root: Vec<String> = Vec::new();
    let mut office_relationships_prefixes_sheets: Vec<String> = Vec::new();
    let mut office_relationships_prefixes_in_scope: Vec<String> = Vec::new();
    let mut in_sheets = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                collect_office_relationships_prefixes_from_xmlns(
                    e,
                    &mut office_relationships_prefixes_root,
                    &mut office_relationships_prefixes_in_scope,
                )?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheets" => {
                if sheets_prefix.is_none() {
                    sheets_prefix = name_prefix(e.name().as_ref());
                }
                collect_office_relationships_prefixes_from_xmlns(
                    e,
                    &mut office_relationships_prefixes_sheets,
                    &mut office_relationships_prefixes_in_scope,
                )?;
                in_sheets = true;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheets" => {
                if sheets_prefix.is_none() {
                    sheets_prefix = name_prefix(e.name().as_ref());
                }
                collect_office_relationships_prefixes_from_xmlns(
                    e,
                    &mut office_relationships_prefixes_sheets,
                    &mut office_relationships_prefixes_in_scope,
                )?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheets" => {
                in_sheets = false;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_sheets && e.local_name().as_ref() == b"sheet" =>
            {
                if sheet_tag.is_none() {
                    sheet_tag = Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                if rel_id_attr.is_none() {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"id" || key.ends_with(b":id") {
                            rel_id_attr = std::str::from_utf8(key).ok().map(|s| s.to_string());
                            break;
                        }
                    }
                }
            }
            _ => {}
        }
        if sheet_tag.is_some() && rel_id_attr.is_some() {
            break;
        }
        buf.clear();
    }

    let sheet_tag =
        sheet_tag.or_else(|| Some(crate::xml::prefixed_tag(sheets_prefix.as_deref(), "sheet")));

    // We must never emit a namespace prefix that isn't declared in the scope we preserve when
    // rewriting `<sheets>`. Namespace declarations that only appeared on the original `<sheet>`
    // elements will be dropped when we replace them with new elements.
    //
    // Therefore:
    // - If we found an existing `*:id` attribute, only keep its prefix if that prefix is declared
    //   on the workbook root or `<sheets>` element.
    // - Otherwise, use an in-scope relationships prefix if one exists, or fall back to an
    //   unprefixed `id` to keep the output namespace-well-formed.
    let fallback_prefix = office_relationships_prefixes_root
        .first()
        .or_else(|| office_relationships_prefixes_sheets.first())
        .map(|s| s.as_str());

    let rel_id_attr = match rel_id_attr {
        Some(found) => match found.split_once(':') {
            // Existing attribute is unprefixed (`id`) => always safe to keep as-is.
            None => Some(found),
            Some((prefix, _)) => {
                if office_relationships_prefixes_in_scope
                    .iter()
                    .any(|p| p.as_str() == prefix)
                {
                    Some(found)
                } else if let Some(fallback_prefix) = fallback_prefix {
                    Some(crate::xml::prefixed_tag(Some(fallback_prefix), "id"))
                } else {
                    Some("id".to_string())
                }
            }
        },
        None => Some(crate::xml::prefixed_tag(fallback_prefix, "id")),
    };

    Ok((sheet_tag, rel_id_attr))
}

fn name_prefix(name: &[u8]) -> Option<String> {
    name.iter()
        .rposition(|b| *b == b':')
        .and_then(|idx| std::str::from_utf8(&name[..idx]).ok())
        .map(|s| s.to_string())
}

fn collect_office_relationships_prefixes_from_xmlns(
    e: &BytesStart<'_>,
    out: &mut Vec<String>,
    in_scope: &mut Vec<String>,
) -> Result<(), XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        if attr.value.as_ref() == crate::xml::OFFICE_RELATIONSHIPS_NS.as_bytes() {
            let prefix = String::from_utf8_lossy(prefix).into_owned();
            if !out.iter().any(|p| p == &prefix) {
                out.push(prefix.clone());
            }
            if !in_scope.iter().any(|p| p == &prefix) {
                in_scope.push(prefix);
            }
        }
    }
    Ok(())
}

pub fn parse_sheet_tab_color(worksheet_xml: &str) -> Result<Option<TabColor>, XlsxError> {
    let mut reader = Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_sheet_pr = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"sheetPr" => {
                in_sheet_pr = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"sheetPr" => {
                in_sheet_pr = false;
            }
            Event::Empty(e) | Event::Start(e)
                if in_sheet_pr && e.local_name().as_ref() == b"tabColor" =>
            {
                return Ok(Some(parse_tab_color_element(&e)?));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn parse_tab_color_element(e: &BytesStart<'_>) -> Result<TabColor, XlsxError> {
    let mut color = TabColor::default();

    for attr in e.attributes() {
        let attr = attr?;
        let value = attr.unescape_value()?.to_string();
        match attr.key.as_ref() {
            b"rgb" => color.rgb = Some(value),
            b"theme" => color.theme = value.parse().ok(),
            b"indexed" => color.indexed = value.parse().ok(),
            b"tint" => color.tint = value.parse().ok(),
            b"auto" => color.auto = value.parse::<u32>().ok().map(|v| v != 0),
            _ => {}
        }
    }

    Ok(color)
}

pub fn write_sheet_tab_color(
    worksheet_xml: &str,
    tab_color: Option<&TabColor>,
) -> Result<String, XlsxError> {
    let has_sheet_pr = worksheet_has_sheet_pr(worksheet_xml)?;
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(worksheet_xml)?;
    let inserted_sheet_pr_tag = crate::xml::prefixed_tag(worksheet_prefix.as_deref(), "sheetPr");
    let inserted_tab_color_tag = crate::xml::prefixed_tag(worksheet_prefix.as_deref(), "tabColor");

    if tab_color.is_none() && !has_sheet_pr {
        return Ok(worksheet_xml.to_string());
    }

    let mut reader = Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut in_sheet_pr = false;
    let mut tab_color_written = false;
    let mut inserted_sheet_pr = false;
    let mut sheet_pr_prefix: Option<String> = None;
    let mut skip_tab_color_depth: usize = 0;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        if skip_tab_color_depth > 0 {
            match event {
                Event::Start(_) => skip_tab_color_depth += 1,
                Event::End(_) => skip_tab_color_depth = skip_tab_color_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            }
            buf.clear();
            continue;
        }

        match event {
            Event::Eof => break,
            Event::Empty(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if let Some(color) = tab_color {
                    let worksheet_tag =
                        String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;
                    writer.write_event(Event::Start(BytesStart::new(
                        inserted_sheet_pr_tag.as_str(),
                    )))?;
                    writer.write_event(Event::Empty(build_tab_color_element(
                        inserted_tab_color_tag.as_str(),
                        color,
                    )))?;
                    writer.write_event(Event::End(BytesEnd::new(
                        inserted_sheet_pr_tag.as_str(),
                    )))?;
                    writer.write_event(Event::End(BytesEnd::new(worksheet_tag.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
                break;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"worksheet" => {
                writer.write_event(Event::Start(e.to_owned()))?;
                if tab_color.is_some() && !has_sheet_pr && !inserted_sheet_pr {
                    writer.write_event(Event::Start(BytesStart::new(
                        inserted_sheet_pr_tag.as_str(),
                    )))?;
                    writer.write_event(Event::Empty(build_tab_color_element(
                        inserted_tab_color_tag.as_str(),
                        tab_color.expect("color present"),
                    )))?;
                    writer.write_event(Event::End(BytesEnd::new(
                        inserted_sheet_pr_tag.as_str(),
                    )))?;
                    inserted_sheet_pr = true;
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetPr" => {
                in_sheet_pr = true;
                tab_color_written = false;
                let name = e.name();
                let name = name.as_ref();
                sheet_pr_prefix = name
                    .iter()
                    .rposition(|b| *b == b':')
                    .map(|idx| &name[..idx])
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetPr" => {
                if let Some(color) = tab_color {
                    let sheet_pr_tag =
                        String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let sheet_pr_prefix = sheet_pr_tag
                        .split_once(':')
                        .map(|(p, _)| p.to_string());
                    let tab_color_tag =
                        crate::xml::prefixed_tag(sheet_pr_prefix.as_deref(), "tabColor");

                    writer.write_event(Event::Start(e.to_owned()))?;
                    writer.write_event(Event::Empty(build_tab_color_element(
                        tab_color_tag.as_str(),
                        color,
                    )))?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_pr_tag.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheetPr" => {
                if in_sheet_pr && tab_color.is_some() && !tab_color_written {
                    let tab_color_tag =
                        crate::xml::prefixed_tag(sheet_pr_prefix.as_deref(), "tabColor");
                    writer.write_event(Event::Empty(build_tab_color_element(
                        tab_color_tag.as_str(),
                        tab_color.expect("color present"),
                    )))?;
                }
                in_sheet_pr = false;
                sheet_pr_prefix = None;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Empty(ref e) if in_sheet_pr && e.local_name().as_ref() == b"tabColor" => {
                if let Some(color) = tab_color {
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Empty(build_tab_color_element(tag.as_str(), color)))?;
                    tab_color_written = true;
                }
            }
            Event::Start(ref e) if in_sheet_pr && e.local_name().as_ref() == b"tabColor" => {
                if let Some(color) = tab_color {
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Empty(build_tab_color_element(tag.as_str(), color)))?;
                    tab_color_written = true;
                }
                // Swallow the original <tabColor> subtree (including its </tabColor> end tag)
                // because we either replaced it with an empty element or removed it.
                skip_tab_color_depth = 1;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn build_tab_color_element(tag: &str, color: &TabColor) -> BytesStart<'static> {
    let mut elem = BytesStart::new(tag).into_owned();
    if let Some(rgb) = &color.rgb {
        elem.push_attribute(("rgb", rgb.as_str()));
    }
    if let Some(theme) = color.theme {
        elem.push_attribute(("theme", theme.to_string().as_str()));
    }
    if let Some(indexed) = color.indexed {
        elem.push_attribute(("indexed", indexed.to_string().as_str()));
    }
    if let Some(tint) = color.tint {
        elem.push_attribute(("tint", tint.to_string().as_str()));
    }
    if let Some(auto) = color.auto {
        elem.push_attribute(("auto", if auto { "1" } else { "0" }));
    }
    elem
}

fn worksheet_has_sheet_pr(worksheet_xml: &str) -> Result<bool, XlsxError> {
    let mut reader = Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_worksheet = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"worksheet" => {
                in_worksheet = true;
            }
            Event::End(e) if e.local_name().as_ref() == b"worksheet" => break,
            Event::Start(e) | Event::Empty(e)
                if in_worksheet && e.local_name().as_ref() == b"sheetPr" =>
            {
                return Ok(true);
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn workbook_sheets_round_trip() {
        let workbook_xml = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" r:id="rId2" state="hidden"/>
    <sheet name="VeryHidden" sheetId="3" r:id="rId3" state="veryHidden"/>
  </sheets>
</workbook>
"#;

        let sheets = parse_workbook_sheets(workbook_xml).unwrap();
        assert_eq!(
            sheets,
            vec![
                WorkbookSheetInfo {
                    name: "Sheet1".to_string(),
                    sheet_id: 1,
                    rel_id: "rId1".to_string(),
                    visibility: SheetVisibility::Visible,
                },
                WorkbookSheetInfo {
                    name: "Hidden".to_string(),
                    sheet_id: 2,
                    rel_id: "rId2".to_string(),
                    visibility: SheetVisibility::Hidden,
                },
                WorkbookSheetInfo {
                    name: "VeryHidden".to_string(),
                    sheet_id: 3,
                    rel_id: "rId3".to_string(),
                    visibility: SheetVisibility::VeryHidden,
                },
            ]
        );

        let mut updated = sheets.clone();
        updated.swap(0, 2);
        updated[0].name = "Renamed".to_string();
        updated[0].visibility = SheetVisibility::Hidden;

        let rewritten = write_workbook_sheets(workbook_xml, &updated).unwrap();
        let reparsed = parse_workbook_sheets(&rewritten).unwrap();
        assert_eq!(reparsed, updated);
    }

    #[test]
    fn write_workbook_sheets_preserves_unknown_sheet_attributes() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1" customAttr="a" xr:uid="u1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2" customAttr="b" xr:uid="u2"/>
  </sheets>
</workbook>
"#;

        let sheets = parse_workbook_sheets(workbook_xml).unwrap();
        let mut updated = sheets.clone();
        updated.swap(0, 1);
        updated[0].name = "Renamed".to_string();

        let rewritten = write_workbook_sheets(workbook_xml, &updated).unwrap();

        let doc = roxmltree::Document::parse(&rewritten)
            .expect("rewritten workbook.xml should be valid XML");

        // Find output sheet elements by relationship id and verify unknown attributes were carried
        // over.
        let rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let xr_ns = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

        let mut seen = 0;
        for sheet in doc.descendants().filter(|n| n.tag_name().name() == "sheet") {
            let rel_id = sheet
                .attribute((rel_ns, "id"))
                .expect("sheet should have r:id");
            match rel_id {
                "rId1" => {
                    assert_eq!(sheet.attribute("customAttr"), Some("a"));
                    assert_eq!(sheet.attribute((xr_ns, "uid")), Some("u1"));
                    seen += 1;
                }
                "rId2" => {
                    assert_eq!(sheet.attribute("customAttr"), Some("b"));
                    assert_eq!(sheet.attribute((xr_ns, "uid")), Some("u2"));
                    seen += 1;
                }
                other => panic!("unexpected rel id: {other}"),
            }
        }
        assert_eq!(seen, 2);
    }

    #[test]
    fn workbook_sheets_rewrite_preserves_prefix_only_namespaces() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
    <x:sheet name="Sheet2" sheetId="2" rel:id="rId2"/>
  </x:sheets>
</x:workbook>
"#;

        let sheets = vec![
            WorkbookSheetInfo {
                name: "Sheet1".to_string(),
                sheet_id: 1,
                rel_id: "rId1".to_string(),
                visibility: SheetVisibility::Visible,
            },
            WorkbookSheetInfo {
                name: "Sheet2".to_string(),
                sheet_id: 2,
                rel_id: "rId2".to_string(),
                visibility: SheetVisibility::Visible,
            },
        ];

        let mut updated = sheets.clone();
        updated.swap(0, 1);
        updated[0].name = "Renamed".to_string();
        updated[0].visibility = SheetVisibility::Hidden;

        let rewritten = write_workbook_sheets(workbook_xml, &updated).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(rewritten.contains("<x:sheet name="));
        assert!(!rewritten.contains("<sheet name="));
        assert!(rewritten.contains("rel:id="));
        assert!(!rewritten.contains(" r:id="));
    }

    #[test]
    fn write_workbook_sheets_expands_self_closing_sheets_prefixed() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</x:workbook>
"#;

        let sheets = vec![WorkbookSheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];

        let rewritten = write_workbook_sheets(workbook_xml, &sheets).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(
            rewritten.contains("<x:sheets") && rewritten.contains("</x:sheets>"),
            "expected output to expand <x:sheets/>, got:\n{rewritten}"
        );
        assert!(rewritten.contains(r#"<x:sheet name=""#));
        assert!(rewritten.contains(r#"name="Sheet1""#));
        assert!(rewritten.contains(r#"rel:id="rId1""#));
        assert!(!rewritten.contains(" r:id="));
    }

    #[test]
    fn write_workbook_sheets_expands_self_closing_sheets_default_ns() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

        let sheets = vec![WorkbookSheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];

        let rewritten = write_workbook_sheets(workbook_xml, &sheets).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(rewritten.contains(r#"<sheet name=""#));
        assert!(rewritten.contains(r#"name="Sheet1""#));
        assert!(rewritten.contains(r#"r:id="rId1""#));
    }

    #[test]
    fn write_workbook_sheets_no_undeclared_relationship_prefix_fallback() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets/>
</x:workbook>
"#;

        let sheets = vec![WorkbookSheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];

        let rewritten = write_workbook_sheets(workbook_xml, &sheets).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(
            rewritten.contains(r#"id="rId1""#),
            "expected output to fall back to an unprefixed id attribute, got:\n{rewritten}"
        );
        assert!(rewritten.contains(r#"<x:sheet name=""#));
        assert!(!rewritten.contains(" r:id="));
        assert!(!rewritten.contains(" rel:id="));
    }

    #[test]
    fn write_workbook_sheets_does_not_emit_sheet_scoped_relationship_prefix() {
        // Relationships prefix is only declared on the original `<sheet>` element, not on the
        // workbook root or `<sheets>`. Since we replace all `<sheet>` elements, we must not keep
        // an `r:id` attribute that would become undeclared after rewriting.
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheets>
    <x:sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>
"#;

        let sheets = vec![WorkbookSheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];

        let rewritten = write_workbook_sheets(workbook_xml, &sheets).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(rewritten.contains(r#"<x:sheet name=""#));
        assert!(rewritten.contains(r#"id="rId1""#));
        assert!(
            !rewritten.contains(" r:id="),
            "should not emit an undeclared r:id prefix, got:\n{rewritten}"
        );
    }

    #[test]
    fn write_workbook_sheets_prefers_existing_rel_id_prefix_when_multiple_prefixes_in_scope() {
        // When multiple prefixes are bound to the relationships namespace, prefer the exact
        // `*:id` qname used by existing <sheet> elements (as long as the prefix is declared in
        // scope).
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>
"#;

        let sheets = vec![WorkbookSheetInfo {
            name: "Sheet1".to_string(),
            sheet_id: 1,
            rel_id: "rId1".to_string(),
            visibility: SheetVisibility::Visible,
        }];

        let rewritten = write_workbook_sheets(workbook_xml, &sheets).unwrap();

        roxmltree::Document::parse(&rewritten).expect("rewritten workbook.xml should be valid XML");
        assert!(
            rewritten.contains(r#"r:id="rId1""#),
            "expected output to keep r:id, got:\n{rewritten}"
        );
        assert!(
            !rewritten.contains(r#"rel:id="rId1""#),
            "should not rewrite to a different relationships prefix, got:\n{rewritten}"
        );
    }

    #[test]
    fn sheet_tab_color_round_trip() {
        let sheet_xml = r#"
 <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr>
    <tabColor rgb="FFFF0000"/>
  </sheetPr>
</worksheet>
"#;

        let color = parse_sheet_tab_color(sheet_xml).unwrap().unwrap();
        assert_eq!(color.rgb.as_deref(), Some("FFFF0000"));

        let new_color = TabColor::rgb("FF00FF00");
        let rewritten = write_sheet_tab_color(sheet_xml, Some(&new_color)).unwrap();
        let reparsed = parse_sheet_tab_color(&rewritten).unwrap().unwrap();
        assert_eq!(reparsed.rgb.as_deref(), Some("FF00FF00"));
    }

    #[test]
    fn remove_sheet_tab_color() {
        let sheet_xml = r#"
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr>
    <tabColor rgb="FFFF0000"/>
  </sheetPr>
</worksheet>
"#;

        let rewritten = write_sheet_tab_color(sheet_xml, None).unwrap();
        assert_eq!(parse_sheet_tab_color(&rewritten).unwrap(), None);
    }
}
