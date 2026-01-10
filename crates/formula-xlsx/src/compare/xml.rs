use quick_xml::events::Event;
use quick_xml::Reader;

/// Normalize XML for semantic comparisons by:
/// - parsing the document
/// - sorting attributes
/// - ignoring insignificant whitespace between elements
///
/// This is intentionally *not* a full XML canonicalization algorithm. It is a
/// pragmatic helper for round-trip testing of SpreadsheetML parts.
pub fn normalize_xml(bytes: &[u8]) -> Result<String, quick_xml::Error> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = String::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                out.push('<');
                out.push_str(std::str::from_utf8(e.name().as_ref()).unwrap_or(""));
                write_sorted_attrs(&mut out, &e)?;
                out.push('>');
            }
            Event::Empty(e) => {
                out.push('<');
                out.push_str(std::str::from_utf8(e.name().as_ref()).unwrap_or(""));
                write_sorted_attrs(&mut out, &e)?;
                out.push_str("/>");
            }
            Event::End(e) => {
                out.push_str("</");
                out.push_str(std::str::from_utf8(e.name().as_ref()).unwrap_or(""));
                out.push('>');
            }
            Event::Text(e) => {
                let text = e.unescape()?.into_owned();
                if !text.chars().all(|c| c.is_whitespace()) {
                    out.push_str(&escape_text(&text));
                }
            }
            Event::CData(e) => {
                let text = String::from_utf8_lossy(e.as_ref());
                if !text.chars().all(|c: char| c.is_whitespace()) {
                    out.push_str(&escape_text(&text));
                }
            }
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {}
            Event::Eof => break,
        }
        buf.clear();
    }

    Ok(out)
}

pub fn assert_xml_semantic_eq(expected: &[u8], actual: &[u8]) {
    let expected_norm = normalize_xml(expected).expect("normalize expected xml");
    let actual_norm = normalize_xml(actual).expect("normalize actual xml");
    assert_eq!(expected_norm, actual_norm);
}

fn write_sorted_attrs(
    out: &mut String,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), quick_xml::Error> {
    let mut attrs: Vec<(String, String)> = Vec::new();
    for attr in e.attributes() {
        let attr = attr?;
        let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("").to_string();
        let value = attr.unescape_value()?.into_owned();
        attrs.push((key, value));
    }
    attrs.sort_by(|a, b| a.0.cmp(&b.0));
    for (k, v) in attrs {
        out.push(' ');
        out.push_str(&k);
        out.push_str("=\"");
        out.push_str(&escape_attr(&v));
        out.push('"');
    }
    Ok(())
}

fn escape_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\"', "&quot;")
        .replace('\'', "&apos;")
}

fn escape_text(value: &str) -> String {
    value.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}
