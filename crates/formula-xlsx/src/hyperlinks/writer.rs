use std::collections::{BTreeMap, BTreeSet};

use formula_model::{Hyperlink, HyperlinkTarget};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::XlsxError;

const NS_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_OFFICE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

fn insert_before_tag(name: &[u8]) -> bool {
    matches!(
        name,
        // Elements that come after <hyperlinks> in the SpreadsheetML schema.
        b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"drawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

#[derive(Clone, Debug)]
struct Relationship {
    id: String,
    ty: String,
    target: String,
    target_mode: Option<String>,
}

fn parse_relationships(rels_xml: &str) -> Result<Vec<Relationship>, XlsxError> {
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"Relationship" => {
                let mut id: Option<String> = None;
                let mut ty: Option<String> = None;
                let mut target: Option<String> = None;
                let mut target_mode: Option<String> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let value = attr.unescape_value()?.to_string();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => ty = Some(value),
                        b"Target" => target = Some(value),
                        b"TargetMode" => target_mode = Some(value),
                        _ => {}
                    }
                }

                let Some(id) = id else { continue };
                rels.push(Relationship {
                    id,
                    ty: ty.unwrap_or_default(),
                    target: target.unwrap_or_default(),
                    target_mode,
                });
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}

/// Update the worksheet XML with a `<hyperlinks>` element representing `hyperlinks`.
///
/// If the worksheet already contains `<hyperlinks>`, it is replaced. If it does not
/// and `hyperlinks` is non-empty, the element is inserted before the end of the worksheet
/// (preferably before elements that are required to come after it, e.g. `<printOptions>`,
/// `<pageMargins>`, `<drawing>`, `<tableParts>`, `<extLst>`).
pub fn update_worksheet_xml(sheet_xml: &str, hyperlinks: &[Hyperlink]) -> Result<String, XlsxError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => {
                match event {
                    Event::Start(_) => skip_depth += 1,
                    Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                    Event::Empty(_) => {}
                    _ => {}
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"hyperlinks" => {
                replaced = true;
                if !hyperlinks.is_empty() {
                    write_hyperlinks_block(&mut writer, hyperlinks, worksheet_prefix.as_deref())?;
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"hyperlinks" => {
                replaced = true;
                if !hyperlinks.is_empty() {
                    write_hyperlinks_block(&mut writer, hyperlinks, worksheet_prefix.as_deref())?;
                }
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if !replaced
                    && !hyperlinks.is_empty()
                    && insert_before_tag(e.local_name().as_ref()) =>
            {
                write_hyperlinks_block(&mut writer, hyperlinks, worksheet_prefix.as_deref())?;
                replaced = true;
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !replaced && !hyperlinks.is_empty() {
                    write_hyperlinks_block(&mut writer, hyperlinks, worksheet_prefix.as_deref())?;
                    replaced = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    Ok(String::from_utf8(writer.into_inner())?)
}

fn write_hyperlinks_block<W: std::io::Write>(
    writer: &mut Writer<W>,
    hyperlinks: &[Hyperlink],
    prefix: Option<&str>,
) -> Result<(), XlsxError> {
    let hyperlinks_tag = crate::xml::prefixed_tag(prefix, "hyperlinks");
    let hyperlink_tag = crate::xml::prefixed_tag(prefix, "hyperlink");

    let mut start = BytesStart::new(hyperlinks_tag.as_str());
    // Declare the `r:` prefix locally so we can always emit `r:id`.
    start.push_attribute(("xmlns:r", NS_OFFICE_REL));
    writer.write_event(Event::Start(start))?;

    for link in hyperlinks {
        let mut elem = BytesStart::new(hyperlink_tag.as_str());

        let range = link.range.to_string();
        elem.push_attribute(("ref", range.as_str()));

        match &link.target {
            HyperlinkTarget::Internal { sheet, cell } => {
                let mut location = String::new();
                formula_model::push_sheet_name_a1(&mut location, sheet);
                location.push('!');
                formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut location);
                elem.push_attribute(("location", location.as_str()));
            }
            HyperlinkTarget::ExternalUrl { .. } | HyperlinkTarget::Email { .. } => {
                let rid = link.rel_id.as_deref().ok_or_else(|| {
                    XlsxError::Hyperlink(format!(
                        "external hyperlink {} missing rel_id",
                        link.range
                    ))
                })?;
                elem.push_attribute(("r:id", rid));
            }
        }

        if let Some(display) = &link.display {
            elem.push_attribute(("display", display.as_str()));
        }
        if let Some(tooltip) = &link.tooltip {
            elem.push_attribute(("tooltip", tooltip.as_str()));
        }

        writer.write_event(Event::Empty(elem))?;
    }

    writer.write_event(Event::End(BytesEnd::new(hyperlinks_tag.as_str())))?;
    Ok(())
}

/// Update the worksheet `.rels` XML to reflect the external hyperlinks in `hyperlinks`.
///
/// Returns `None` if the resulting relationships set is empty.
pub fn update_worksheet_relationships(
    rels_xml: Option<&str>,
    hyperlinks: &[Hyperlink],
) -> Result<Option<String>, XlsxError> {
    let mut rels = rels_xml
        .map(parse_relationships)
        .transpose()?
        .unwrap_or_default();

    // Desired hyperlink relationships from the model.
    let mut desired = BTreeMap::<String, Relationship>::new();
    for link in hyperlinks {
        let (rid, target) = match &link.target {
            HyperlinkTarget::ExternalUrl { uri } => (link.rel_id.as_deref(), uri.as_str()),
            HyperlinkTarget::Email { uri } => (link.rel_id.as_deref(), uri.as_str()),
            HyperlinkTarget::Internal { .. } => continue,
        };
        let Some(rid) = rid else {
            return Err(XlsxError::Hyperlink(format!(
                "external hyperlink {} is missing rel_id",
                link.range
            )));
        };
        desired.insert(
            rid.to_string(),
            Relationship {
                id: rid.to_string(),
                ty: REL_TYPE_HYPERLINK.to_string(),
                target: target.to_string(),
                target_mode: Some("External".to_string()),
            },
        );
    }

    let desired_ids: BTreeSet<String> = desired.keys().cloned().collect();
    rels.retain(|r| r.ty != REL_TYPE_HYPERLINK || desired_ids.contains(&r.id));

    for (id, wanted) in desired {
        match rels.iter_mut().find(|r| r.id == id) {
            Some(existing) => {
                existing.ty = wanted.ty;
                existing.target = wanted.target;
                existing.target_mode = wanted.target_mode;
            }
            None => rels.push(wanted),
        }
    }

    if rels.is_empty() {
        return Ok(None);
    }

    rels.sort_by(|a, b| a.id.cmp(&b.id));
    Ok(Some(render_relationships_xml(&rels)?))
}

fn render_relationships_xml(rels: &[Relationship]) -> Result<String, XlsxError> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut root = BytesStart::new("Relationships");
    root.push_attribute(("xmlns", NS_RELATIONSHIPS));
    writer.write_event(Event::Start(root))?;

    for rel in rels {
        let mut elem = BytesStart::new("Relationship");
        elem.push_attribute(("Id", rel.id.as_str()));
        elem.push_attribute(("Type", rel.ty.as_str()));
        elem.push_attribute(("Target", rel.target.as_str()));
        if let Some(mode) = &rel.target_mode {
            elem.push_attribute(("TargetMode", mode.as_str()));
        }
        writer.write_event(Event::Empty(elem))?;
    }

    writer.write_event(Event::End(BytesEnd::new("Relationships")))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_model::{CellRef, Range};

    #[test]
    fn update_inserts_before_table_parts_when_missing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>"#;
        let links = vec![Hyperlink {
            range: Range::from_a1("A1").unwrap(),
            target: HyperlinkTarget::Internal {
                sheet: "Sheet1".to_string(),
                cell: CellRef::new(0, 0),
            },
            display: None,
            tooltip: None,
            rel_id: None,
        }];

        let updated = update_worksheet_xml(xml, &links).unwrap();

        let links_pos = updated.find("<hyperlinks").expect("hyperlinks inserted");
        let table_pos = updated.find("<tableParts").expect("tableParts exists");
        assert!(
            links_pos < table_pos,
            "expected hyperlinks before tableParts, got:\n{updated}"
        );
    }

    #[test]
    fn update_inserts_before_page_margins_when_missing() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#;
        let links = vec![Hyperlink {
            range: Range::from_a1("A1").unwrap(),
            target: HyperlinkTarget::Internal {
                sheet: "Sheet1".to_string(),
                cell: CellRef::new(0, 0),
            },
            display: None,
            tooltip: None,
            rel_id: None,
        }];

        let updated = update_worksheet_xml(xml, &links).unwrap();

        let links_pos = updated.find("<hyperlinks").expect("hyperlinks inserted");
        let margins_pos = updated.find("<pageMargins").expect("pageMargins exists");
        assert!(
            links_pos < margins_pos,
            "expected hyperlinks before pageMargins, got:\n{updated}"
        );
    }
}
