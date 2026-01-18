use quick_xml::events::Event;
use quick_xml::Reader;
use thiserror::Error;

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
    // Track whether whitespace-only text nodes should be preserved in the current
    // element context.
    //
    // Per XML, `xml:space="preserve"` is inherited by descendants. SpreadsheetML
    // commonly uses this on `<t>` elements where a leading/trailing/standalone
    // space is semantically significant.
    let mut preserve_space: Vec<bool> = vec![false];

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                let mut this_preserve = preserve_space.last().copied().unwrap_or(false);
                for attr in e.attributes() {
                    let attr = attr?;
                    if is_xml_space_attr(attr.key.as_ref()) {
                        let value = attr.unescape_value()?.into_owned();
                        match value.as_str() {
                            "preserve" => this_preserve = true,
                            "default" => this_preserve = false,
                            _ => {}
                        }
                    }
                }
                preserve_space.push(this_preserve);
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
                preserve_space.pop();
                out.push_str("</");
                out.push_str(std::str::from_utf8(e.name().as_ref()).unwrap_or(""));
                out.push('>');
            }
            Event::Text(e) => {
                let text = e.unescape()?.into_owned();
                let preserve = preserve_space.last().copied().unwrap_or(false);
                if preserve || !text.chars().all(|c| c.is_whitespace()) {
                    out.push_str(&escape_text(&text));
                }
            }
            Event::CData(e) => {
                let text = String::from_utf8_lossy(e.as_ref());
                let preserve = preserve_space.last().copied().unwrap_or(false);
                if preserve || !text.chars().all(|c: char| c.is_whitespace()) {
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

fn is_xml_space_attr(key: &[u8]) -> bool {
    // The XML namespace (`xml:*`) prefix is reserved and fixed, so we can match
    // the literal attribute name. Keep this helper small and allocation-free.
    key == b"xml:space"
}

#[derive(Debug, Error)]
pub enum XmlSemanticEqError {
    #[error("failed to parse xml: {0}")]
    Parse(#[from] quick_xml::Error),
    #[error("xml differs after normalization\nexpected: {expected}\nactual: {actual}")]
    Mismatch { expected: String, actual: String },
}

pub fn assert_xml_semantic_eq(expected: &[u8], actual: &[u8]) -> Result<(), XmlSemanticEqError> {
    let expected_norm = normalize_xml(expected)?;
    let actual_norm = normalize_xml(actual)?;
    if expected_norm == actual_norm {
        Ok(())
    } else {
        Err(XmlSemanticEqError::Mismatch {
            expected: expected_norm,
            actual: actual_norm,
        })
    }
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

#[cfg(test)]
mod tests {
    use super::normalize_xml;

    #[test]
    fn normalize_xml_respects_xml_space_preserve_for_whitespace_only_text() {
        let xml = br#"<root><t xml:space="preserve"> </t><t> </t></root>"#;
        let normalized = normalize_xml(xml).expect("normalize xml");
        assert_eq!(
            normalized,
            r#"<root><t xml:space="preserve"> </t><t></t></root>"#
        );
    }
}
