use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use std::io::Cursor;

use super::{parse_worksheet_conditional_formatting, ConditionalFormattingError, ParsedConditionalFormatting};

const NS_X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const NS_XM: &str = "http://schemas.microsoft.com/office/excel/2006/main";

#[derive(Debug, thiserror::Error)]
pub enum ConditionalFormattingStreamingError {
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error(transparent)]
    Parse(#[from] ConditionalFormattingError),
    #[error("invalid worksheet xml: {0}")]
    Invalid(&'static str),
    #[error("missing worksheet root element")]
    MissingWorksheetRoot,
}

/// Parse conditional formatting rules from a worksheet XML payload without DOM-parsing
/// the entire worksheet.
///
/// This scans the worksheet with [`quick_xml`] and extracts only:
/// - SpreadsheetML `<conditionalFormatting>` blocks (Office 2007 schema)
/// - x14 `<x14:conditionalFormattings>` blocks (extended schema, usually under `<extLst>`)
///
/// The extracted blocks are wrapped in a minimal `<worksheet>` wrapper that copies the
/// namespace declarations from the original root element (supporting prefixed roots like
/// `<x:worksheet ...>`).
///
/// If no conditional formatting is present, this returns [`ParsedConditionalFormatting::default`]
/// without allocating wrapper XML or invoking the DOM parser.
pub fn parse_worksheet_conditional_formatting_streaming(
    worksheet_xml: &str,
) -> Result<ParsedConditionalFormatting, ConditionalFormattingStreamingError> {
    // Common fast-path: most worksheets have no conditional formatting. Avoid
    // any allocations/DOM parsing in that case.
    if !worksheet_xml.contains("conditionalFormatting") {
        return Ok(ParsedConditionalFormatting::default());
    }

    let Some(wrapper) = extract_conditional_formatting_wrapper(worksheet_xml)? else {
        return Ok(ParsedConditionalFormatting::default());
    };

    Ok(parse_worksheet_conditional_formatting(&wrapper)?)
}

fn extract_conditional_formatting_wrapper(
    worksheet_xml: &str,
) -> Result<Option<String>, ConditionalFormattingStreamingError> {
    let mut reader = Reader::from_str(worksheet_xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();

    let mut root_qname: Option<String> = None;
    let mut root_ns_attrs: Vec<(String, String)> = Vec::new();
    let mut has_x14 = false;
    let mut has_xm = false;

    // Depth of open elements inside the worksheet root (root itself excluded).
    let mut depth: usize = 0;

    let mut extracted_blocks: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                if root_qname.is_none() {
                    root_qname = Some(std::str::from_utf8(e.name().as_ref()).unwrap_or("worksheet").to_string());

                    for attr in e.attributes().with_checks(false) {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns" || key.starts_with(b"xmlns:") {
                            let key_str = std::str::from_utf8(key).unwrap_or_default().to_string();
                            let val = attr.unescape_value()?.into_owned();
                            if key_str == "xmlns:x14" {
                                has_x14 = true;
                            } else if key_str == "xmlns:xm" {
                                has_xm = true;
                            }
                            root_ns_attrs.push((key_str, val));
                        }
                    }
                    buf.clear();
                    continue;
                }

                let local = crate::openxml::local_name(e.name().into_inner());
                if local == b"conditionalFormatting" && depth == 0 {
                    let xml =
                        capture_element_xml(&mut reader, Event::Start(e.into_owned()), &mut buf)?;
                    extracted_blocks.push(xml);
                    buf.clear();
                    continue;
                }
                if local == b"conditionalFormattings" {
                    let xml =
                        capture_element_xml(&mut reader, Event::Start(e.into_owned()), &mut buf)?;
                    extracted_blocks.push(xml);
                    buf.clear();
                    continue;
                }

                depth = depth
                    .checked_add(1)
                    .ok_or(ConditionalFormattingStreamingError::Invalid(
                        "worksheet depth overflow",
                    ))?;
            }
            Event::Empty(e) => {
                if root_qname.is_none() {
                    // Degenerate `<worksheet/>` root - treat as no conditional formatting.
                    break;
                }

                let local = crate::openxml::local_name(e.name().into_inner());
                if local == b"conditionalFormatting" && depth == 0 {
                    let xml =
                        capture_element_xml(&mut reader, Event::Empty(e.into_owned()), &mut buf)?;
                    extracted_blocks.push(xml);
                    buf.clear();
                    continue;
                }
                if local == b"conditionalFormattings" {
                    let xml =
                        capture_element_xml(&mut reader, Event::Empty(e.into_owned()), &mut buf)?;
                    extracted_blocks.push(xml);
                    buf.clear();
                    continue;
                }
            }
            Event::End(_) => {
                if root_qname.is_none() {
                    continue;
                }
                if depth == 0 {
                    break;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or(ConditionalFormattingStreamingError::Invalid(
                        "worksheet depth underflow",
                    ))?;
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    if extracted_blocks.is_empty() {
        return Ok(None);
    }

    let root_qname = root_qname.ok_or(ConditionalFormattingStreamingError::MissingWorksheetRoot)?;

    let mut out = String::new();
    // Small-ish. Avoid lots of reallocations when there are many rules.
    let cap: usize = extracted_blocks.iter().map(|s| s.len()).sum::<usize>() + 256;
    out.reserve(cap);

    out.push('<');
    out.push_str(&root_qname);
    for (k, v) in &root_ns_attrs {
        out.push(' ');
        out.push_str(k);
        out.push_str("=\"");
        out.push_str(v);
        out.push('"');
    }
    if !has_x14 {
        out.push_str(" xmlns:x14=\"");
        out.push_str(NS_X14);
        out.push('"');
    }
    if !has_xm {
        out.push_str(" xmlns:xm=\"");
        out.push_str(NS_XM);
        out.push('"');
    }
    out.push('>');

    for block in extracted_blocks {
        out.push_str(&block);
    }

    out.push_str("</");
    out.push_str(&root_qname);
    out.push('>');

    Ok(Some(out))
}

fn capture_element_xml<B: std::io::BufRead>(
    reader: &mut Reader<B>,
    first: Event<'static>,
    buf: &mut Vec<u8>,
) -> Result<String, quick_xml::Error> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    match first {
        Event::Empty(e) => {
            writer.write_event(Event::Empty(e))?;
        }
        Event::Start(e) => {
            writer.write_event(Event::Start(e))?;
            let mut depth: usize = 0;
            loop {
                match reader.read_event_into(buf)? {
                    Event::Start(e) => {
                        depth += 1;
                        writer.write_event(Event::Start(e.into_owned()))?;
                    }
                    Event::Empty(e) => {
                        writer.write_event(Event::Empty(e.into_owned()))?;
                    }
                    Event::End(e) => {
                        writer.write_event(Event::End(e.into_owned()))?;
                        if depth == 0 {
                            break;
                        }
                        depth -= 1;
                    }
                    Event::Eof => break,
                    ev => {
                        writer.write_event(ev.into_owned())?;
                    }
                }
                buf.clear();
            }
        }
        _ => {}
    }

    let bytes = writer.into_inner().into_inner();
    Ok(String::from_utf8_lossy(&bytes).to_string())
}
