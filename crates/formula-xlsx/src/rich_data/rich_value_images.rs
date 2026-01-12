use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml;
use crate::{XlsxError, XlsxPackage};

/// A parsed `<rv>` entry from a `xl/richData/richValue*.xml` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueEntry {
    /// The rich value part this entry came from (e.g. `xl/richData/richValue1.xml`).
    pub source_part: String,
    /// The first `r:embed="rId*"` discovered within the `<rv>` subtree (best-effort).
    pub embed_rel_id: Option<String>,
}

/// Warnings produced while resolving rich values across parts.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RichValueWarning {
    /// Multiple `<rv>` elements declared the same explicit global index.
    ///
    /// Resolution is deterministic: the first `<rv>` encountered wins.
    DuplicateIndex {
        index: u32,
        first_part: String,
        second_part: String,
    },
    /// `metadata.xml` referenced a rich value index that is not present in any richValue part.
    MissingRichValue { index: u32 },
    /// A rich value referenced a relationship ID, but it could not be resolved to a package part.
    MissingRelationship {
        index: u32,
        source_part: String,
        rel_id: String,
    },
    /// A relationship resolved to a target part that was missing from the package.
    MissingTargetPart { index: u32, part: String },
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RichValueIndex {
    pub entries: HashMap<u32, RichValueEntry>,
    pub warnings: Vec<RichValueWarning>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ExtractedRichValueImages {
    /// Map of rich value global index -> raw image bytes.
    pub images: HashMap<u32, Vec<u8>>,
    pub warnings: Vec<RichValueWarning>,
}

impl XlsxPackage {
    /// Extract images referenced by rich values.
    ///
    /// This supports multi-part rich value stores (`xl/richData/richValue*.xml`) and honors
    /// explicit `<rv i="…">` / `<rv id="…">` / `<rv idx="…">` global indices when present.
    ///
    /// Note: this is currently a best-effort extractor intended for cell-image rich values. It
    /// only looks for `r:embed` on `<blip>` elements within `<rv>` and resolves the target via
    /// the part's `.rels` file.
    pub fn extract_rich_value_images(&self) -> Result<ExtractedRichValueImages, XlsxError> {
        let Some(metadata) = self.part("xl/metadata.xml") else {
            return Ok(ExtractedRichValueImages::default());
        };

        let referenced = parse_metadata_rich_value_indices(metadata)?;
        if referenced.is_empty() {
            return Ok(ExtractedRichValueImages::default());
        }

        let index = build_rich_value_index(self.parts_map())?;

        let mut out = ExtractedRichValueImages {
            images: HashMap::new(),
            warnings: index.warnings,
        };

        for rv_index in referenced {
            let Some(entry) = index.entries.get(&rv_index) else {
                out.warnings
                    .push(RichValueWarning::MissingRichValue { index: rv_index });
                continue;
            };

            let Some(rel_id) = &entry.embed_rel_id else {
                // Not an image rich value (or we failed to parse). Silently skip.
                continue;
            };

            let Some(target_part) =
                openxml::resolve_relationship_target(self, &entry.source_part, rel_id)?
            else {
                out.warnings.push(RichValueWarning::MissingRelationship {
                    index: rv_index,
                    source_part: entry.source_part.clone(),
                    rel_id: rel_id.clone(),
                });
                continue;
            };

            let Some(bytes) = self.part(&target_part) else {
                out.warnings.push(RichValueWarning::MissingTargetPart {
                    index: rv_index,
                    part: target_part,
                });
                continue;
            };

            out.images.insert(rv_index, bytes.to_vec());
        }

        Ok(out)
    }
}

fn build_rich_value_index(parts: &BTreeMap<String, Vec<u8>>) -> Result<RichValueIndex, XlsxError> {
    let mut parsed: Vec<ParsedRv> = Vec::new();

    // Deterministic part ordering (BTreeMap keys are already sorted).
    for (part_name, bytes) in parts.iter().filter(|(name, _)| is_rich_value_part(name)) {
        parsed.extend(parse_rich_value_part(part_name, bytes)?);
    }

    let mut warnings = Vec::new();
    let mut entries: HashMap<u32, RichValueEntry> = HashMap::new();

    let mut max_explicit: Option<u32> = None;
    for rv in &parsed {
        let Some(idx) = rv.explicit_index else {
            continue;
        };

        if let Some(existing) = entries.get(&idx) {
            warnings.push(RichValueWarning::DuplicateIndex {
                index: idx,
                first_part: existing.source_part.clone(),
                second_part: rv.entry.source_part.clone(),
            });
            continue;
        }

        max_explicit = Some(max_explicit.map(|m| m.max(idx)).unwrap_or(idx));
        entries.insert(idx, rv.entry.clone());
    }

    // Assign indices for entries without explicit IDs.
    let mut next = match max_explicit {
        Some(max) => max.checked_add(1).ok_or_else(|| {
            XlsxError::Invalid("rich value index overflow while assigning implicit ids".to_string())
        })?,
        None => 0,
    };
    for rv in &parsed {
        if rv.explicit_index.is_some() {
            continue;
        }
        // `next` should not collide with explicit indices given our starting point, but keep it
        // robust (and deterministic) if the input is pathological.
        while entries.contains_key(&next) {
            next = next.checked_add(1).ok_or_else(|| {
                XlsxError::Invalid(
                    "rich value index overflow while resolving implicit id collisions".to_string(),
                )
            })?;
        }
        entries.insert(next, rv.entry.clone());
        next = next.checked_add(1).ok_or_else(|| {
            XlsxError::Invalid("rich value index overflow while assigning implicit ids".to_string())
        })?;
    }

    Ok(RichValueIndex { entries, warnings })
}

#[derive(Debug, Clone)]
struct ParsedRv {
    explicit_index: Option<u32>,
    entry: RichValueEntry,
}

fn is_rich_value_part(part_name: &str) -> bool {
    const PREFIX: &str = "xl/richData/richValue";
    if !part_name.starts_with(PREFIX) || !part_name.ends_with(".xml") {
        return false;
    }
    let rest = &part_name[PREFIX.len()..part_name.len() - ".xml".len()];
    rest.is_empty() || rest.chars().all(|c| c.is_ascii_digit())
}

fn parse_rich_value_part(part_name: &str, bytes: &[u8]) -> Result<Vec<ParsedRv>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(bytes));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"rv" => {
                let explicit_index = parse_rv_index_attr(&e)?;
                let embed_rel_id = parse_first_embed_rel_id_within_rv(&mut reader)?;
                out.push(ParsedRv {
                    explicit_index,
                    entry: RichValueEntry {
                        source_part: part_name.to_string(),
                        embed_rel_id,
                    },
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"rv" => {
                let explicit_index = parse_rv_index_attr(&e)?;
                out.push(ParsedRv {
                    explicit_index,
                    entry: RichValueEntry {
                        source_part: part_name.to_string(),
                        embed_rel_id: None,
                    },
                });
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rv_index_attr(e: &quick_xml::events::BytesStart<'_>) -> Result<Option<u32>, XlsxError> {
    for attr in e.attributes() {
        let attr = attr?;
        let key = openxml::local_name(attr.key.as_ref());
        if key.eq_ignore_ascii_case(b"i")
            || key.eq_ignore_ascii_case(b"id")
            || key.eq_ignore_ascii_case(b"idx")
        {
            let value = attr.unescape_value()?.into_owned();
            let idx = value.parse::<u32>().map_err(|e| {
                XlsxError::Invalid(format!(
                    "invalid rich value <rv> global index {value:?}: {e}"
                ))
            })?;
            return Ok(Some(idx));
        }
    }
    Ok(None)
}

fn parse_first_embed_rel_id_within_rv<R: std::io::BufRead>(
    reader: &mut Reader<R>,
) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut embed_rel_id: Option<String> = None;

    // Track nesting depth of `<rv>` tags so we don't get confused if the subtree contains nested
    // `<rv>` elements (unlikely but defensive).
    let mut rv_depth: usize = 0;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"rv" => {
                rv_depth += 1;
            }
            Event::End(e) if e.local_name().as_ref() == b"rv" => {
                if rv_depth == 0 {
                    break;
                }
                rv_depth -= 1;
            }
            Event::Start(e) | Event::Empty(e) => {
                if embed_rel_id.is_none()
                    && openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"blip")
                {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"embed") {
                            embed_rel_id = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }
                }
            }
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while parsing <rv> subtree".to_string(),
                ))
            }
            _ => {}
        }

        buf.clear();
    }

    Ok(embed_rel_id)
}

fn parse_metadata_rich_value_indices(metadata: &[u8]) -> Result<Vec<u32>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(metadata));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut seen: HashSet<u32> = HashSet::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"rvb") =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"i") {
                        let value = attr.unescape_value()?.into_owned();
                        if let Ok(idx) = value.parse::<u32>() {
                            if seen.insert(idx) {
                                out.push(idx);
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(out)
}
