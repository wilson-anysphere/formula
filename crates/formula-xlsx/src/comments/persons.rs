use std::collections::BTreeMap;
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

#[derive(Clone, Debug)]
pub struct PersonsParseError;

impl From<quick_xml::Error> for PersonsParseError {
    fn from(_value: quick_xml::Error) -> Self {
        Self
    }
}

pub fn parse_persons_xml(bytes: &[u8]) -> Result<BTreeMap<String, String>, PersonsParseError> {
    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut current_person_id: Option<String> = None;
    let mut out = BTreeMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                match e.local_name().as_ref() {
                    b"person" => {
                        let id = attr_value(&reader, &e, b"id")
                            .or_else(|| attr_value(&reader, &e, b"personId"));
                        if let Some(id) = id {
                            current_person_id = Some(id.clone());
                            if let Some(display) = attr_value(&reader, &e, b"displayName") {
                                out.insert(id, display);
                                current_person_id = None;
                            }
                        }
                    }
                    b"personPr" => {
                        if let Some(id) = current_person_id.clone() {
                            if let Some(display) = attr_value(&reader, &e, b"displayName") {
                                out.insert(id, display);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => match e.local_name().as_ref() {
                b"person" => {
                    let id = attr_value(&reader, &e, b"id")
                        .or_else(|| attr_value(&reader, &e, b"personId"));
                    if let Some(id) = id {
                        if let Some(display) = attr_value(&reader, &e, b"displayName") {
                            out.insert(id, display);
                        } else {
                            current_person_id = Some(id);
                        }
                    }
                }
                b"personPr" => {
                    if let Some(id) = current_person_id.take() {
                        if let Some(display) = attr_value(&reader, &e, b"displayName") {
                            out.insert(id, display);
                        }
                    }
                }
                _ => {}
            },
            Event::End(e) => {
                if e.local_name().as_ref() == b"person" {
                    current_person_id = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(out)
}

fn attr_value<R: std::io::BufRead>(
    reader: &Reader<R>,
    element: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> Option<String> {
    for attr in element.attributes().with_checks(false) {
        let attr = attr.ok()?;
        if attr.key.local_name().as_ref() != key {
            continue;
        }
        let _ = reader;
        return attr.unescape_value().ok().map(|value| value.to_string());
    }
    None
}
