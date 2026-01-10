use std::collections::BTreeMap;

use formula_model::{CellRef, Hyperlink, HyperlinkTarget, Range};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::XlsxError;

const REL_TYPE_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

#[derive(Clone, Debug)]
struct Relationship {
    ty: String,
    target: String,
}

fn parse_relationships(rels_xml: &str) -> Result<BTreeMap<String, Relationship>, XlsxError> {
    let mut reader = Reader::from_str(rels_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut rels = BTreeMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"Relationship" => {
                let mut id: Option<String> = None;
                let mut ty: Option<String> = None;
                let mut target: Option<String> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let value = attr.unescape_value()?.to_string();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(value),
                        b"Type" => ty = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }

                let Some(id) = id else { continue };
                let rel = Relationship {
                    ty: ty.unwrap_or_default(),
                    target: target.unwrap_or_default(),
                };
                rels.insert(id, rel);
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(rels)
}

/// Parse `<hyperlinks>` from a worksheet XML part, using the optional `.rels` part
/// to resolve external targets.
pub fn parse_worksheet_hyperlinks(
    sheet_xml: &str,
    rels_xml: Option<&str>,
) -> Result<Vec<Hyperlink>, XlsxError> {
    let rels = rels_xml
        .map(parse_relationships)
        .transpose()?
        .unwrap_or_default();

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Empty(e) | Event::Start(e) if e.local_name().as_ref() == b"hyperlink" => {
                out.push(parse_hyperlink_element(&e, &rels)?);
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_hyperlink_element(
    e: &BytesStart<'_>,
    rels: &BTreeMap<String, Relationship>,
) -> Result<Hyperlink, XlsxError> {
    let mut reference: Option<String> = None;
    let mut rid: Option<String> = None;
    let mut location: Option<String> = None;
    let mut display: Option<String> = None;
    let mut tooltip: Option<String> = None;

    for attr in e.attributes() {
        let attr = attr?;
        let value = attr.unescape_value()?.to_string();
        match attr.key.as_ref() {
            b"ref" => reference = Some(value),
            b"location" => location = Some(value),
            b"display" => display = Some(value),
            b"tooltip" => tooltip = Some(value),
            b"r:id" => rid = Some(value),
            _ => {}
        }
    }

    let reference = reference.ok_or(XlsxError::MissingAttr("ref"))?;
    let range = parse_range(&reference)?;

    let target = if let Some(location) = location {
        let (sheet, cell) = parse_internal_location(&location)?;
        HyperlinkTarget::Internal { sheet, cell }
    } else if let Some(rid) = &rid {
        let rel = rels.get(rid).ok_or_else(|| {
            XlsxError::Hyperlink(format!("hyperlink references missing relationship {rid}"))
        })?;
        if rel.ty != REL_TYPE_HYPERLINK {
            return Err(XlsxError::Hyperlink(format!(
                "relationship {rid} has unexpected Type {} (expected hyperlink)",
                rel.ty
            )));
        }
        let uri = rel.target.clone();
        if uri.to_ascii_lowercase().starts_with("mailto:") {
            HyperlinkTarget::Email { uri }
        } else {
            HyperlinkTarget::ExternalUrl { uri }
        }
    } else {
        return Err(XlsxError::Hyperlink(format!(
            "hyperlink {reference} missing location or r:id"
        )));
    };

    Ok(Hyperlink {
        range,
        target,
        display,
        tooltip,
        rel_id: rid,
    })
}

fn parse_range(a1: &str) -> Result<Range, XlsxError> {
    let trimmed = a1.trim();
    let (start, end) = match trimmed.split_once(':') {
        Some((a, b)) => (cell_from_a1(a)?, cell_from_a1(b)?),
        None => {
            let c = cell_from_a1(trimmed)?;
            (c, c)
        }
    };
    Ok(Range::new(start, end))
}

fn cell_from_a1(a1: &str) -> Result<CellRef, XlsxError> {
    CellRef::from_a1(a1).map_err(|e| XlsxError::Hyperlink(format!("invalid A1 ref {a1}: {e}")))
}

fn parse_internal_location(location: &str) -> Result<(String, CellRef), XlsxError> {
    let mut loc = location.trim();
    if let Some(rest) = loc.strip_prefix('#') {
        loc = rest;
    }

    let (sheet, cell) = loc.split_once('!').ok_or_else(|| {
        XlsxError::Hyperlink(format!("invalid hyperlink location (missing '!'): {location}"))
    })?;

    let sheet = unquote_sheet_name(sheet.trim());
    let cell_str = cell.trim();
    let cell_str = cell_str
        .split_once(':')
        .map(|(start, _)| start)
        .unwrap_or(cell_str);
    let cell = cell_from_a1(cell_str)?;
    Ok((sheet, cell))
}

fn unquote_sheet_name(name: &str) -> String {
    // Excel quotes sheet names with single quotes; embedded quotes are doubled.
    let mut s = name.trim();
    if s.starts_with('\'') && s.ends_with('\'') && s.len() >= 2 {
        s = &s[1..s.len() - 1];
        return s.replace("''", "'");
    }
    s.to_string()
}

