use std::borrow::Cow;

use formula_model::rich_text::{RichText, RichTextRunStyle, Underline};
use formula_model::Color;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader;
use thiserror::Error;

use super::SharedStrings;

const OOXML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Debug, Error)]
pub enum SharedStringsError {
    #[error("xml parse error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("malformed sharedStrings.xml: {0}")]
    Malformed(&'static str),
}

pub fn parse_shared_strings_xml(xml: &str) -> Result<SharedStrings, SharedStringsError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut items = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"sst" => {
                // Validate namespace but don't fail hard; we care about local names.
                if let Some(ns) = attr_value(&e, b"xmlns")? {
                    if ns != OOXML_NS {
                        // Still proceed; some producers omit/alter namespaces.
                    }
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"si" => {
                items.push(parse_si(&mut reader)?);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(SharedStrings { items })
}

fn parse_si(reader: &mut Reader<&[u8]>) -> Result<RichText, SharedStringsError> {
    let mut buf = Vec::new();
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();
    let mut plain_text: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                let t = read_text(reader, QName(b"t"))?;
                if segments.is_empty() {
                    plain_text.get_or_insert_with(String::new).push_str(&t);
                } else {
                    // Mixed `<t>` and `<r>` is not expected in `si`.
                    return Err(SharedStringsError::Malformed(
                        "shared string item mixes <t> and <r>",
                    ));
                }
            }
            Event::Start(e) if e.local_name().as_ref() == b"r" => {
                if plain_text.is_some() {
                    return Err(SharedStringsError::Malformed(
                        "shared string item mixes <t> and <r>",
                    ));
                }
                segments.push(parse_r(reader)?);
            }
            Event::End(e) if e.local_name().as_ref() == b"si" => break,
            Event::Eof => return Err(SharedStringsError::Malformed("unexpected eof in <si>")),
            _ => {}
        }
        buf.clear();
    }

    if !segments.is_empty() {
        Ok(RichText::from_segments(segments))
    } else {
        Ok(RichText::new(plain_text.unwrap_or_default()))
    }
}

fn parse_r(reader: &mut Reader<&[u8]>) -> Result<(String, RichTextRunStyle), SharedStringsError> {
    let mut buf = Vec::new();
    let mut style = RichTextRunStyle::default();
    let mut text = String::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"rPr" => {
                style = parse_rpr(reader)?;
            }
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                text.push_str(&read_text(reader, QName(b"t"))?);
            }
            Event::End(e) if e.local_name().as_ref() == b"r" => break,
            Event::Eof => return Err(SharedStringsError::Malformed("unexpected eof in <r>")),
            _ => {}
        }
        buf.clear();
    }

    Ok((text, style))
}

fn parse_rpr(reader: &mut Reader<&[u8]>) -> Result<RichTextRunStyle, SharedStringsError> {
    let mut buf = Vec::new();
    let mut style = RichTextRunStyle::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) => parse_rpr_tag(reader, &e, &mut style)?,
            Event::Start(e) => {
                parse_rpr_tag(reader, &e, &mut style)?;
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"rPr" => break,
            Event::Eof => return Err(SharedStringsError::Malformed("unexpected eof in <rPr>")),
            _ => {}
        }
        buf.clear();
    }

    Ok(style)
}

fn parse_rpr_tag(
    _reader: &Reader<&[u8]>,
    e: &quick_xml::events::BytesStart<'_>,
    style: &mut RichTextRunStyle,
) -> Result<(), SharedStringsError> {
    match e.local_name().as_ref() {
        b"b" => style.bold = Some(parse_bool_val(e)?),
        b"i" => style.italic = Some(parse_bool_val(e)?),
        b"u" => {
            let val = attr_value(e, b"val")?;
            if let Some(ul) = Underline::from_ooxml(val.as_deref()) {
                style.underline = Some(ul);
            }
        }
        b"color" => {
            if let Some(rgb) = attr_value(e, b"rgb")? {
                if rgb.len() == 8 {
                    if let Ok(argb) = u32::from_str_radix(&rgb, 16) {
                        style.color = Some(Color::new_argb(argb));
                    }
                }
            }
        }
        b"rFont" | b"name" => {
            if let Some(val) = attr_value(e, b"val")? {
                style.font = Some(val);
            }
        }
        b"sz" => {
            if let Some(val) = attr_value(e, b"val")? {
                if let Some(sz) = parse_size_100pt(&val) {
                    style.size_100pt = Some(sz);
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn parse_bool_val(e: &quick_xml::events::BytesStart<'_>) -> Result<bool, SharedStringsError> {
    let Some(val) = attr_value(e, b"val")? else {
        return Ok(true);
    };
    Ok(!(val == "0" || val.eq_ignore_ascii_case("false")))
}

fn read_text(reader: &mut Reader<&[u8]>, end: QName<'_>) -> Result<String, SharedStringsError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => {
                let t: Cow<'_, str> = e.unescape()?;
                text.push_str(&t);
            }
            Event::CData(e) => {
                text.push_str(std::str::from_utf8(e.as_ref())?);
            }
            Event::End(e) if e.name() == end => break,
            Event::Eof => return Err(SharedStringsError::Malformed("unexpected eof in <t>")),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn attr_value(
    e: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> Result<Option<String>, SharedStringsError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn parse_size_100pt(val: &str) -> Option<u16> {
    let val = val.trim();
    if val.is_empty() {
        return None;
    }

    if let Some((int_part, frac_part)) = val.split_once('.') {
        let int: u16 = int_part.parse().ok()?;
        let mut frac = frac_part.chars().take(2).collect::<String>();
        while frac.len() < 2 {
            frac.push('0');
        }
        let frac: u16 = frac.parse().ok()?;
        int.checked_mul(100)?.checked_add(frac)
    } else {
        let int: u16 = val.parse().ok()?;
        int.checked_mul(100)
    }
}
