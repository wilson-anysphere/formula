//! SpreadsheetML metadata (`xl/metadata.xml`) parser.
//!
//! Excel stores additional metadata for cells/values in the workbook part
//! `xl/metadata.xml`. Sheet cells can reference these blocks using the `cm`
//! (cell metadata) and `vm` (value metadata) indices on `<c>` elements.
//!
//! Formula currently treats the rich/linked data payloads as opaque, but having
//! a parser makes it possible to:
//! - debug/work with linked data types and rich values
//! - map `cm`/`vm` indices to concrete metadata records
//! - eventually connect metadata records to `xl/richData/*`
//!
//! The data model intentionally captures only the stable core (`<metadataTypes>`
//! and `<rc>` records) while preserving unknown/extension payloads as raw XML.

use std::collections::BTreeMap;
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use thiserror::Error;

/// Parsed representation of `xl/metadata.xml`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataDocument {
    /// `<metadataTypes>` entries (at least `name=` + all attributes).
    pub metadata_types: Vec<MetadataType>,
    /// Raw inner XML payloads from `<futureMetadata><bk>...</bk></futureMetadata>`.
    ///
    /// Excel uses `futureMetadata` for forward-compatible extension blocks (often containing
    /// `extLst` payloads in non-SpreadsheetML namespaces). We keep the inner XML so higher-level
    /// tooling can inspect/debug it without needing to understand every schema.
    pub future_metadata_blocks: Vec<MetadataBlockRaw>,
    /// `<cellMetadata>` blocks referenced via the `cm` index on worksheet `<c>` elements.
    pub cell_metadata: Vec<MetadataBlock>,
    /// `<valueMetadata>` blocks referenced via the `vm` index on worksheet `<c>` elements.
    pub value_metadata: Vec<MetadataBlock>,
}

impl MetadataDocument {
    /// Lookup a cell metadata block by index.
    ///
    /// Excel's indexing has historically been ambiguous across parts. To be resilient we attempt
    /// `idx` as 0-based first, then fall back to `idx-1` when the direct lookup fails.
    pub fn cell_block_by_index(&self, idx: u32) -> Option<&MetadataBlock> {
        self.block_by_index(&self.cell_metadata, idx)
    }

    /// Lookup a value metadata block by index.
    ///
    /// Excel's indexing has historically been ambiguous across parts. To be resilient we attempt
    /// `idx` as 0-based first, then fall back to `idx-1` when the direct lookup fails.
    pub fn value_block_by_index(&self, idx: u32) -> Option<&MetadataBlock> {
        self.block_by_index(&self.value_metadata, idx)
    }

    fn block_by_index<'a>(&'a self, blocks: &'a [MetadataBlock], idx: u32) -> Option<&'a MetadataBlock> {
        blocks
            .get(idx as usize)
            .or_else(|| idx.checked_sub(1).and_then(|i| blocks.get(i as usize)))
    }
}

/// One `<metadataType>` entry from `<metadataTypes>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataType {
    /// The `name=` attribute.
    pub name: String,
    /// All attributes present on the `<metadataType>` element (including `name`).
    pub attributes: BTreeMap<String, String>,
}

/// Raw block payload from `<futureMetadata><bk>...</bk></futureMetadata>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MetadataBlockRaw {
    /// The raw inner XML of the `<bk>` element (everything between `<bk>` and `</bk>`).
    pub inner_xml: String,
}

/// Parsed metadata block from `<cellMetadata>` / `<valueMetadata>`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataBlock {
    /// `<rc t=".." v=".."/>` entries.
    pub records: Vec<MetadataRecord>,
}

/// A metadata record (`<rc t=".." v=".."/>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MetadataRecord {
    /// `t=` attribute.
    pub t: u32,
    /// `v=` attribute.
    pub v: u32,
}

#[derive(Debug, Error)]
pub enum MetadataError {
    #[error("xml parse error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("malformed metadata.xml: {0}")]
    Malformed(&'static str),
}

/// Parse `xl/metadata.xml`.
///
/// The parser is best-effort and intentionally tolerant:
/// - ignores namespaces/prefixes by matching on local-name only
/// - ignores unknown elements
/// - does not require `count=` attributes to be present or correct
pub fn parse_metadata_xml(xml: &str) -> Result<MetadataDocument, MetadataError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut doc = MetadataDocument::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"metadata" => {
                    // Root element; its children contain the data we care about.
                }
                b"metadataTypes" => {
                    doc.metadata_types
                        .extend(parse_metadata_types(&mut reader)?);
                }
                b"futureMetadata" => {
                    doc.future_metadata_blocks
                        .extend(parse_future_metadata(&mut reader)?);
                }
                b"cellMetadata" => {
                    doc.cell_metadata
                        .extend(parse_metadata_blocks(&mut reader, b"cellMetadata")?);
                }
                b"valueMetadata" => {
                    doc.value_metadata
                        .extend(parse_metadata_blocks(&mut reader, b"valueMetadata")?);
                }
                _ => {
                    // Unknown wrapper - skip subtree best-effort.
                    reader.read_to_end_into(e.name(), &mut Vec::new())?;
                }
            },
            Event::Empty(e) => match e.local_name().as_ref() {
                b"metadata" => {}
                // Empty containers mean "no data" - nothing to do.
                b"metadataTypes" | b"futureMetadata" | b"cellMetadata" | b"valueMetadata" => {}
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(doc)
}

fn parse_metadata_types(reader: &mut Reader<&[u8]>) -> Result<Vec<MetadataType>, MetadataError> {
    let mut buf = Vec::new();
    let mut types = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"metadataType" => {
                let attributes = collect_attributes(&e)?;
                let name = attributes.get("name").cloned().unwrap_or_default();
                types.push(MetadataType { name, attributes });
                // `<metadataType>` is expected to be empty, but skip any subtree just in case.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"metadataType" => {
                let attributes = collect_attributes(&e)?;
                let name = attributes.get("name").cloned().unwrap_or_default();
                types.push(MetadataType { name, attributes });
            }
            Event::Start(e) => {
                // Unknown subtree; skip so we don't accidentally treat nested `<metadataType>`s
                // as real entries.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"metadataTypes" => break,
            Event::Eof => {
                return Err(MetadataError::Malformed(
                    "unexpected eof in <metadataTypes>",
                ))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(types)
}

fn parse_future_metadata(reader: &mut Reader<&[u8]>) -> Result<Vec<MetadataBlockRaw>, MetadataError> {
    let mut buf = Vec::new();
    let mut blocks = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"bk" => {
                let inner_xml = capture_inner_xml(reader, b"bk")?;
                blocks.push(MetadataBlockRaw { inner_xml });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"bk" => {
                blocks.push(MetadataBlockRaw {
                    inner_xml: String::new(),
                });
            }
            Event::Start(e) => {
                // Skip unknown subtree.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"futureMetadata" => break,
            Event::Eof => {
                return Err(MetadataError::Malformed(
                    "unexpected eof in <futureMetadata>",
                ))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(blocks)
}

fn parse_metadata_blocks(
    reader: &mut Reader<&[u8]>,
    end_tag: &'static [u8],
) -> Result<Vec<MetadataBlock>, MetadataError> {
    let mut buf = Vec::new();
    let mut blocks = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"bk" => {
                blocks.push(parse_metadata_block(reader)?);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"bk" => {
                blocks.push(MetadataBlock::default());
            }
            Event::Start(e) => {
                // Skip unknown subtree.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == end_tag => break,
            Event::Eof => return Err(MetadataError::Malformed("unexpected eof in metadata block")),
            _ => {}
        }
        buf.clear();
    }

    Ok(blocks)
}

fn parse_metadata_block(reader: &mut Reader<&[u8]>) -> Result<MetadataBlock, MetadataError> {
    let mut buf = Vec::new();
    let mut records = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"rc" => {
                records.push(parse_record(&e)?);
                // Skip any unexpected children.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"rc" => {
                records.push(parse_record(&e)?);
            }
            Event::Start(e) => {
                // Skip unknown subtree.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"bk" => break,
            Event::Eof => return Err(MetadataError::Malformed("unexpected eof in <bk>")),
            _ => {}
        }
        buf.clear();
    }

    Ok(MetadataBlock { records })
}

fn parse_record(e: &quick_xml::events::BytesStart<'_>) -> Result<MetadataRecord, MetadataError> {
    let mut t: u32 = 0;
    let mut v: u32 = 0;
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        match attr.key.as_ref() {
            b"t" => {
                t = attr.unescape_value()?.parse().unwrap_or(0);
            }
            b"v" => {
                v = attr.unescape_value()?.parse().unwrap_or(0);
            }
            _ => {}
        }
    }
    Ok(MetadataRecord { t, v })
}

fn collect_attributes(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<BTreeMap<String, String>, MetadataError> {
    let mut out = BTreeMap::new();
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        let key = std::str::from_utf8(attr.key.as_ref())?.to_string();
        let val = attr.unescape_value()?.into_owned();
        out.insert(key, val);
    }
    Ok(out)
}

fn capture_inner_xml(reader: &mut Reader<&[u8]>, end_local_name: &'static [u8]) -> Result<String, MetadataError> {
    let mut buf = Vec::new();
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut depth: usize = 0;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => {
                depth += 1;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) => {
                writer.write_event(Event::Empty(e.into_owned()))?;
            }
            Event::End(e) => {
                if depth == 0 && e.local_name().as_ref() == end_local_name {
                    break;
                }
                writer.write_event(Event::End(e.into_owned()))?;
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Err(MetadataError::Malformed("unexpected eof capturing inner xml")),
            ev => {
                writer.write_event(ev.into_owned())?;
            }
        }
        buf.clear();
    }

    let bytes = writer.into_inner().into_inner();
    Ok(std::str::from_utf8(&bytes)?.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;

    #[test]
    fn parses_metadata_types_and_attributes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="2">
    <metadataType name="XLDAPR" minSupportedVersion="120000" something="else"/>
    <metadataType name="Foo" custom="bar"/>
  </metadataTypes>
</metadata>"#;

        let doc = parse_metadata_xml(xml).expect("parse metadata.xml");
        assert_eq!(doc.metadata_types.len(), 2);
        assert_eq!(doc.metadata_types[0].name, "XLDAPR");
        assert_eq!(
            doc.metadata_types[0].attributes.get("minSupportedVersion"),
            Some(&"120000".to_string())
        );
        assert_eq!(
            doc.metadata_types[0].attributes.get("something"),
            Some(&"else".to_string())
        );
        assert_eq!(doc.metadata_types[1].name, "Foo");
        assert_eq!(
            doc.metadata_types[1].attributes.get("custom"),
            Some(&"bar".to_string())
        );
    }

    #[test]
    fn parses_cell_and_value_metadata_blocks() {
        let xml = r#"<metadata>
  <cellMetadata>
    <bk>
      <rc t="0" v="1"/>
      <rc t="1" v="2"/>
      <ignored><rc t="99" v="99"/></ignored>
    </bk>
    <bk>
      <rc t="2" v="3"/>
    </bk>
  </cellMetadata>
  <valueMetadata>
    <bk>
      <rc t="5" v="8"/>
    </bk>
  </valueMetadata>
</metadata>"#;

        let doc = parse_metadata_xml(xml).expect("parse metadata.xml");
        assert_eq!(doc.cell_metadata.len(), 2);
        assert_eq!(doc.cell_metadata[0].records.len(), 2);
        assert_eq!(doc.cell_metadata[0].records[0], MetadataRecord { t: 0, v: 1 });
        assert_eq!(doc.cell_metadata[0].records[1], MetadataRecord { t: 1, v: 2 });
        assert_eq!(doc.cell_metadata[1].records, vec![MetadataRecord { t: 2, v: 3 }]);

        assert_eq!(doc.value_metadata.len(), 1);
        assert_eq!(doc.value_metadata[0].records, vec![MetadataRecord { t: 5, v: 8 }]);
    }

    #[test]
    fn preserves_future_metadata_block_inner_xml() {
        let xml = r#"<metadata>
  <futureMetadata name="XLDAPR">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <foo:bar xmlns:foo="urn:example">payload</foo:bar>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
</metadata>"#;

        let doc = parse_metadata_xml(xml).expect("parse metadata.xml");
        assert_eq!(doc.future_metadata_blocks.len(), 1);
        assert!(
            doc.future_metadata_blocks[0]
                .inner_xml
                .contains("foo:bar"),
            "inner_xml was: {}",
            doc.future_metadata_blocks[0].inner_xml
        );
        assert!(doc.future_metadata_blocks[0].inner_xml.contains("payload"));
    }

    #[test]
    fn index_helpers_try_zero_based_then_one_based_fallback() {
        let xml = r#"<metadata>
  <cellMetadata>
    <bk><rc t="0" v="1"/></bk>
    <bk><rc t="1" v="2"/></bk>
  </cellMetadata>
  <valueMetadata>
    <bk><rc t="5" v="8"/></bk>
  </valueMetadata>
</metadata>"#;

        let doc = parse_metadata_xml(xml).expect("parse metadata.xml");

        // 0-based direct lookup
        assert_eq!(
            doc.cell_block_by_index(0)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 0, v: 1 })
        );

        // Still 0-based when it exists.
        assert_eq!(
            doc.cell_block_by_index(1)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 1, v: 2 })
        );

        // Fallback `idx-1` when `idx` is out of range.
        assert_eq!(
            doc.cell_block_by_index(2)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 1, v: 2 })
        );

        assert!(doc.value_block_by_index(0).is_some());
        assert!(doc.value_block_by_index(1).is_some()); // fallback to idx-1
        assert!(doc.value_block_by_index(2).is_none());
    }
}
