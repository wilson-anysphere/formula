//! SpreadsheetML metadata (`xl/metadata.xml`) parser.
//!
//! Excel stores additional metadata for cells/values in the workbook part `xl/metadata.xml`.
//! Worksheet cells can reference these blocks using the `cm` (cell metadata) and `vm` (value
//! metadata) indices on `<c>` elements.
//!
//! This module provides a best-effort, namespace-agnostic parser that is sufficient to resolve a
//! worksheet cell's `vm` (value metadata) index to the rich value record index used by Excel rich
//! data types / images-in-cell (`XLRICHVALUE`).
//!
//! The relevant indirection chain looks like:
//!
//! - `c/@vm` → `<valueMetadata><bk>` index
//! - `<bk><rc t="…" v="…"/>` → `t` indexes `<metadataTypes><metadataType name="…"/>`
//! - When the referenced metadataType is `XLRICHVALUE`, `v` indexes
//!   `<futureMetadata name="XLRICHVALUE"><bk>…</bk></futureMetadata>`
//! - Inside the futureMetadata `<bk>`, an extension element (commonly `xlrd:rvb`) has an `i="N"`
//!   attribute, where `N` is the rich value record index.
//!
//! Real-world files sometimes vary between 0-based and 1-based indices.
//!
//! For `c/@vm` and `c/@cm`, Excel uses 1-based indices. Some producers have been observed to emit
//! 0-based indices, or to be off-by-one. To be resilient when both interpretations are plausible,
//! this module prefers the 1-based interpretation (`idx - 1`) for `vm/cm` and falls back to
//! 0-based (`idx`) when the preferred lookup fails.

use std::collections::BTreeMap;
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use thiserror::Error;

/// A `<bk>` block that may be repeated via run-length encoding.
///
/// Excel sometimes compresses repeated metadata blocks by writing a single `<bk count="N">`
/// entry. For index resolution, this block occupies `count` consecutive indices.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RepeatedBlock<T> {
    /// Number of consecutive indices occupied by this block.
    pub count: u32,
    pub block: T,
}

/// Parsed representation of `xl/metadata.xml`.
///
/// The data model captures only the stable core (`<metadataTypes>` and `<rc>` records) while also
/// preserving unknown/extension payloads from `<futureMetadata>` as raw inner XML.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MetadataPart {
    /// `<metadataTypes>` entries (at least `name=` + all attributes).
    pub metadata_types: Vec<MetadataType>,
    /// Raw inner XML payloads from `<futureMetadata><bk>...</bk></futureMetadata>`.
    ///
    /// Excel uses `futureMetadata` for forward-compatible extension blocks (often containing
    /// `extLst` payloads in non-SpreadsheetML namespaces). We keep the inner XML so higher-level
    /// tooling can inspect/debug it without needing to understand every schema.
    pub future_metadata_blocks: Vec<RepeatedBlock<MetadataBlockRaw>>,
    /// `<cellMetadata>` blocks referenced via the `cm` index on worksheet `<c>` elements.
    pub cell_metadata: Vec<RepeatedBlock<MetadataBlock>>,
    /// `<valueMetadata>` blocks referenced via the `vm` index on worksheet `<c>` elements.
    pub value_metadata: Vec<RepeatedBlock<MetadataBlock>>,
    /// `rvb/@i` values from `<futureMetadata name="XLRICHVALUE">`, indexed by `<bk>` position.
    ///
    /// This is stored separately from `future_metadata_blocks` so callers can resolve rich values
    /// without parsing the extension XML. The entries may be run-length encoded via `bk/@count`.
    xlrichvalue_future_bks: Vec<RepeatedBlock<Option<u32>>>,
}

/// Backwards-compatible name for the parsed `xl/metadata.xml` representation.
pub type MetadataDocument = MetadataPart;

impl MetadataPart {
    /// Lookup a cell metadata block by index.
    ///
    /// Excel's indexing has historically been ambiguous across parts. To be resilient we attempt
    /// `idx` as 1-based first (`idx-1`), then fall back to `idx` as 0-based when the preferred
    /// lookup fails.
    pub fn cell_block_by_index(&self, idx: u32) -> Option<&MetadataBlock> {
        self.block_by_index_one_based_preferred(&self.cell_metadata, idx)
    }

    /// Lookup a value metadata block by index.
    ///
    /// Excel's indexing has historically been ambiguous across parts. To be resilient we attempt
    /// `idx` as 1-based first (`idx-1`), then fall back to `idx` as 0-based when the preferred
    /// lookup fails.
    pub fn value_block_by_index(&self, idx: u32) -> Option<&MetadataBlock> {
        self.block_by_index_one_based_preferred(&self.value_metadata, idx)
    }

    /// Lookup a future-metadata block by index (`<futureMetadata>` `<bk>`).
    pub fn future_block_by_index(&self, idx: u32) -> Option<&MetadataBlockRaw> {
        self.block_by_index(&self.future_metadata_blocks, idx)
    }

    fn block_by_index<'a, T>(&'a self, blocks: &'a [RepeatedBlock<T>], idx: u32) -> Option<&'a T> {
        self.block_by_index_candidate(blocks, idx)
            .or_else(|| idx.checked_sub(1).and_then(|i| self.block_by_index_candidate(blocks, i)))
    }

    fn block_by_index_one_based_preferred<'a, T>(
        &'a self,
        blocks: &'a [RepeatedBlock<T>],
        idx: u32,
    ) -> Option<&'a T> {
        idx.checked_sub(1)
            .and_then(|i| self.block_by_index_candidate(blocks, i))
            .or_else(|| self.block_by_index_candidate(blocks, idx))
    }

    fn block_by_index_candidate<'a, T>(
        &'a self,
        blocks: &'a [RepeatedBlock<T>],
        idx: u32,
    ) -> Option<&'a T> {
        // Walk blocks cumulatively to support run-length encoding via `bk/@count`.
        let mut cursor: u32 = 0;
        for block in blocks {
            let count = block.count.max(1);
            let end = cursor.saturating_add(count);
            if idx < end {
                return Some(&block.block);
            }
            cursor = end;
        }
        None
    }

    /// Resolve a worksheet cell's `vm=` index to the rich value record index (`XLRICHVALUE`).
    ///
    /// Excel's documented behavior is that `c/@vm` is a 1-based index into `<valueMetadata>`'s
    /// `<bk>` list, but real-world files exist that appear to use a 0-based scheme (notably
    /// `vm="0"`).
    ///
    /// This method is best-effort and intentionally tolerant:
    /// - ignores unrelated metadata types and tags
    /// - handles missing/invalid numeric attributes by skipping those entries
    /// - tries both 1-based and 0-based interpretations for `vm`, plus 0/1-based interpretations
    ///   for `rc/@t` and `rc/@v`
    pub fn vm_to_rich_value_index(&self, vm_raw: u32) -> Option<u32> {
        // `vm` is 1-based in the OOXML schema, but tolerate 0-based indexing as a fallback.
        //
        // Try 1-based first so workbooks like `fixtures/xlsx/basic/image-in-cell.xlsx` resolve
        // `vm="1"` to the first `<bk>` record (not the second).
        let mut vm_candidates: [Option<u32>; 2] = [vm_raw.checked_sub(1), Some(vm_raw)];
        // Prefer deterministic iteration order; `checked_sub` might yield `None` for `vm_raw=0`.
        if vm_candidates[0].is_none() {
            vm_candidates.swap(0, 1);
        }

        for vm_idx in vm_candidates.into_iter().flatten() {
            let Some(vm_bk) = self.block_by_index_candidate(&self.value_metadata, vm_idx) else {
                continue;
            };

            // A `<bk>` may contain multiple `<rc>` records for different metadata types.
            for rc in &vm_bk.records {
                // `t` indexes into `<metadataTypes>`.
                for t_idx in index_candidates(rc.t, self.metadata_types.len()) {
                    let Some(metadata_type) = self.metadata_types.get(t_idx) else {
                        continue;
                    };
                    if !metadata_type.name.eq_ignore_ascii_case("XLRICHVALUE") {
                        continue;
                    }

                    // `v` indexes into `<futureMetadata name="XLRICHVALUE"><bk>...</bk></futureMetadata>`.
                    if let Some(Some(rich_idx)) = self
                        .block_by_index(&self.xlrichvalue_future_bks, rc.v)
                        .copied()
                    {
                        return Some(rich_idx);
                    }
                }
            }
        }

        None
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
pub fn parse_metadata_xml(bytes: &[u8]) -> Result<MetadataPart, MetadataError> {
    let xml = std::str::from_utf8(bytes)?;
    // Be tolerant of leading whitespace/newlines before the XML declaration and optional UTF-8 BOM.
    let xml = xml.trim_start_matches('\u{feff}').trim_start();

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    let mut doc = MetadataPart::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) => match e.local_name().as_ref() {
                b"metadata" => {
                    // Root element; its children contain the data we care about.
                }
                b"metadataTypes" => {
                    doc.metadata_types.extend(parse_metadata_types(&mut reader)?);
                }
                b"futureMetadata" => {
                    let name = read_attr_local_string(&e, b"name")?;
                    let is_xlrichvalue = name
                        .as_deref()
                        .is_some_and(|n| n.eq_ignore_ascii_case("XLRICHVALUE"));
                    let (blocks, rvb_indices) = parse_future_metadata(&mut reader, is_xlrichvalue)?;
                    doc.future_metadata_blocks.extend(blocks);
                    if is_xlrichvalue {
                        doc.xlrichvalue_future_bks.extend(rvb_indices);
                    }
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

fn parse_future_metadata(
    reader: &mut Reader<&[u8]>,
    is_xlrichvalue: bool,
) -> Result<
    (
        Vec<RepeatedBlock<MetadataBlockRaw>>,
        Vec<RepeatedBlock<Option<u32>>>,
    ),
    MetadataError,
> {
    let mut buf = Vec::new();
    let mut blocks = Vec::new();
    let mut rvb_indices: Vec<RepeatedBlock<Option<u32>>> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"bk" => {
                let count = parse_bk_count(&e)?;
                let inner_xml = capture_inner_xml(reader, b"bk")?;
                if is_xlrichvalue {
                    rvb_indices.push(RepeatedBlock {
                        count,
                        block: extract_rvb_i(&inner_xml),
                    });
                }
                blocks.push(RepeatedBlock {
                    count,
                    block: MetadataBlockRaw { inner_xml },
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"bk" => {
                let count = parse_bk_count(&e)?;
                if is_xlrichvalue {
                    rvb_indices.push(RepeatedBlock { count, block: None });
                }
                blocks.push(RepeatedBlock {
                    count,
                    block: MetadataBlockRaw {
                        inner_xml: String::new(),
                    },
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

    Ok((blocks, rvb_indices))
}

fn parse_metadata_blocks(
    reader: &mut Reader<&[u8]>,
    end_tag: &'static [u8],
) -> Result<Vec<RepeatedBlock<MetadataBlock>>, MetadataError> {
    let mut buf = Vec::new();
    let mut blocks = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"bk" => {
                let count = parse_bk_count(&e)?;
                blocks.push(RepeatedBlock {
                    count,
                    block: parse_metadata_block(reader)?,
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"bk" => {
                let count = parse_bk_count(&e)?;
                blocks.push(RepeatedBlock {
                    count,
                    block: MetadataBlock::default(),
                });
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
                if let Some(record) = parse_record(&e)? {
                    records.push(record);
                }
                // Skip any unexpected children.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"rc" => {
                if let Some(record) = parse_record(&e)? {
                    records.push(record);
                }
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

fn parse_record(e: &quick_xml::events::BytesStart<'_>) -> Result<Option<MetadataRecord>, MetadataError> {
    let mut t: Option<u32> = None;
    let mut v: Option<u32> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        match crate::openxml::local_name(attr.key.as_ref()) {
            b"t" => {
                let val = attr.unescape_value()?;
                t = val.parse::<u32>().ok();
            }
            b"v" => {
                let val = attr.unescape_value()?;
                v = val.parse::<u32>().ok();
            }
            _ => {}
        }
    }

    Ok(match (t, v) {
        (Some(t), Some(v)) => Some(MetadataRecord { t, v }),
        _ => None,
    })
}

fn collect_attributes(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<BTreeMap<String, String>, MetadataError> {
    let mut out = BTreeMap::new();
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        let key = std::str::from_utf8(crate::openxml::local_name(attr.key.as_ref()))?.to_string();
        let val = attr.unescape_value()?.into_owned();
        out.insert(key, val);
    }
    Ok(out)
}

fn capture_inner_xml(
    reader: &mut Reader<&[u8]>,
    end_local_name: &'static [u8],
) -> Result<String, MetadataError> {
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
                if depth == 0 {
                    return Err(MetadataError::Malformed(
                        "unexpected end tag while capturing inner xml",
                    ));
                }
                writer.write_event(Event::End(e.into_owned()))?;
                depth -= 1;
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

fn parse_bk_count(e: &quick_xml::events::BytesStart<'_>) -> Result<u32, MetadataError> {
    // `bk/@count` is optional run-length encoding. Default to 1 when missing/invalid.
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        if crate::openxml::local_name(attr.key.as_ref()) == b"count" {
            return Ok(attr
                .unescape_value()?
                .trim()
                .parse::<u32>()
                .ok()
                .filter(|v| *v >= 1)
                .unwrap_or(1));
        }
    }
    Ok(1)
}

fn extract_rvb_i(inner_xml: &str) -> Option<u32> {
    // `bk` payloads are extension-heavy (often include non-SpreadsheetML prefixes). quick_xml is
    // intentionally namespace-agnostic so we can scan by local-name only.
    let mut reader = Reader::from_str(inner_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e) | Event::Empty(e)) if e.local_name().as_ref() == b"rvb" => {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.ok()?;
                    if crate::openxml::local_name(attr.key.as_ref()) == b"i" {
                        let val = attr.unescape_value().ok()?;
                        return val.parse::<u32>().ok();
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }

    None
}

fn index_candidates(raw: u32, len: usize) -> impl Iterator<Item = usize> {
    let raw_idx = raw as usize;
    let a = (raw_idx < len).then_some(raw_idx);
    let b = if raw > 0 {
        let fallback = (raw - 1) as usize;
        (fallback < len && fallback != raw_idx).then_some(fallback)
    } else {
        None
    };
    [a, b].into_iter().filter_map(|v| v)
}

fn read_attr_local_string(
    e: &quick_xml::events::BytesStart<'_>,
    key: &'static [u8],
) -> Result<Option<String>, MetadataError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(quick_xml::Error::from)?;
        if crate::openxml::local_name(attr.key.as_ref()) == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
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

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");
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

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");
        assert_eq!(doc.cell_metadata.len(), 2);
        assert_eq!(doc.cell_metadata[0].block.records.len(), 2);
        assert_eq!(
            doc.cell_metadata[0].block.records[0],
            MetadataRecord { t: 0, v: 1 }
        );
        assert_eq!(
            doc.cell_metadata[0].block.records[1],
            MetadataRecord { t: 1, v: 2 }
        );
        assert_eq!(
            doc.cell_metadata[1].block.records,
            vec![MetadataRecord { t: 2, v: 3 }]
        );

        assert_eq!(doc.value_metadata.len(), 1);
        assert_eq!(
            doc.value_metadata[0].block.records,
            vec![MetadataRecord { t: 5, v: 8 }]
        );
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

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");
        assert_eq!(doc.future_metadata_blocks.len(), 1);
        assert!(
            doc.future_metadata_blocks[0]
                .block
                .inner_xml
                .contains("foo:bar"),
            "inner_xml was: {}",
            doc.future_metadata_blocks[0].block.inner_xml
        );
        assert!(doc.future_metadata_blocks[0].block.inner_xml.contains("payload"));
    }

    #[test]
    fn bk_count_is_respected_for_value_metadata_lookup() {
        // One <bk count="3"> should occupy indices 0,1,2 for lookup purposes.
        let xml = r#"<metadata>
  <valueMetadata>
    <bk count="3"><rc t="0" v="1"/></bk>
    <bk><rc t="0" v="2"/></bk>
  </valueMetadata>
 </metadata>"#;

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");

        // A cell with vm="2" should still resolve into the first <bk count="3"> block.
        assert_eq!(
            doc.value_block_by_index(2)
                .and_then(|b| b.records.first().copied()),
            Some(MetadataRecord { t: 0, v: 1 })
        );
    }

    #[test]
    fn bk_count_is_respected_for_future_metadata_lookup() {
        // One <bk count="2"> should occupy indices 0,1 for lookup purposes.
        let xml = r#"<metadata>
  <futureMetadata name="XLDAPR">
    <bk count="2"><foo/></bk>
    <bk><bar/></bk>
  </futureMetadata>
 </metadata>"#;

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");
        assert_eq!(doc.future_metadata_blocks.len(), 2);

        // An rc with v="1" should resolve into the first block (<bk count="2">).
        let block = doc.future_block_by_index(1).expect("future block resolves");
        assert!(block.inner_xml.contains("foo"));
    }

    #[test]
    fn index_helpers_prefer_one_based_for_cm_vm_and_fallback_to_zero_based() {
        let xml = r#"<metadata>
  <cellMetadata>
    <bk><rc t="0" v="1"/></bk>
    <bk><rc t="1" v="2"/></bk>
  </cellMetadata>
  <valueMetadata>
    <bk><rc t="5" v="8"/></bk>
  </valueMetadata>
</metadata>"#;

        let doc = parse_metadata_xml(xml.as_bytes()).expect("parse metadata.xml");

        // Still resolves `0` via 0-based lookup (the preferred 1-based candidate is out of range).
        assert_eq!(
            doc.cell_block_by_index(0)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 0, v: 1 })
        );

        // Prefer 1-based (`idx-1`) for cm/vm: `1` resolves to the first block.
        assert_eq!(
            doc.cell_block_by_index(1)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 0, v: 1 })
        );

        // And `2` resolves to the second block.
        assert_eq!(
            doc.cell_block_by_index(2)
                .unwrap()
                .records
                .first()
                .copied(),
            Some(MetadataRecord { t: 1, v: 2 })
        );

        assert!(doc.cell_block_by_index(3).is_none());

        assert!(doc.value_block_by_index(0).is_some());
        assert!(doc.value_block_by_index(1).is_some()); // 1-based (`idx-1`)
        assert!(doc.value_block_by_index(2).is_none());
    }

    #[test]
    fn vm_to_rich_value_index_zero_based() {
        let xml = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes>
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="7"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata>
    <bk>
      <rc t="0" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

        let metadata = parse_metadata_xml(xml.as_bytes()).unwrap();
        assert_eq!(metadata.vm_to_rich_value_index(0), Some(7));
    }

    #[test]
    fn vm_to_rich_value_index_one_based_fallback() {
        let xml = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes>
    <metadataType name="xlrichvalue"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="7"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata>
    <bk>
      <!-- All of these indices are 1-based in this synthetic example. -->
      <rc t="1" v="1"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

        let metadata = parse_metadata_xml(xml.as_bytes()).unwrap();
        assert_eq!(metadata.vm_to_rich_value_index(1), Some(7));
    }

    #[test]
    fn vm_to_rich_value_index_prefers_one_based_when_indices_overlap() {
        // Real Excel fixtures often use:
        // - `vm` as 1-based indices into `<valueMetadata>`
        // - but with `vm="0"` also observed for the first entry in some workbooks.
        //
        // When `<valueMetadata>` has multiple `<bk>` entries, interpreting `vm="1"` as 0-based would
        // point at the *second* entry, which is not correct for modern Excel outputs.
        let xml = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><xlrd:rvb i="0"/></ext></extLst></bk>
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000001}"><xlrd:rvb i="1"/></ext></extLst></bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
"#;

        let metadata = parse_metadata_xml(xml.as_bytes()).unwrap();
        assert_eq!(metadata.vm_to_rich_value_index(0), Some(0));
        assert_eq!(metadata.vm_to_rich_value_index(1), Some(0));
        assert_eq!(metadata.vm_to_rich_value_index(2), Some(1));
    }
}
