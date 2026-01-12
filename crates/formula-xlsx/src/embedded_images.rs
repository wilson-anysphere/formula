use std::collections::HashMap;
use std::io::Cursor;

use formula_model::CellRef;
use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml;
use crate::package::{XlsxError, XlsxPackage};
use crate::path;

const REL_TYPE_SHEET_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const REL_TYPE_RD_RICH_VALUE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
const REL_TYPE_RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// Embedded ("Place in Cell") image mapping extracted from an XLSX package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedImageCell {
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub sheet_part: String,
    pub cell: CellRef,
    /// Normalized OPC part path for the image payload (e.g. `xl/media/image1.png`).
    pub image_target: String,
    pub bytes: Vec<u8>,
    pub alt_text: Option<String>,
    /// Whether the image is marked as decorative.
    ///
    /// Derived from the `CalcOrigin` field in rich value data (`5` for decorative, `6` for
    /// informative images with alt text).
    pub decorative: bool,
}

#[derive(Debug, Clone)]
struct RichValueImage {
    local_image_identifier: u32,
    calc_origin: u32,
    alt_text: Option<String>,
}

/// Extract embedded images-in-cells ("Place in Cell") using the `vm` + `metadata.xml` + `xl/richData/*`
/// schema.
///
/// The extractor is intentionally resilient: missing parts or broken mappings result in skipped cells.
/// Only invalid ZIP/XML/UTF-8 errors are returned as [`XlsxError`].
pub fn extract_embedded_images(pkg: &XlsxPackage) -> Result<Vec<EmbeddedImageCell>, XlsxError> {
    // Discover the relevant rich data parts via `xl/_rels/workbook.xml.rels`.
    let workbook_rels_xml = match pkg.part("xl/_rels/workbook.xml.rels") {
        Some(bytes) => bytes,
        None => return Ok(Vec::new()),
    };

    let relationships = openxml::parse_relationships(workbook_rels_xml)?;
    let mut metadata_part: Option<String> = None;
    let mut rdrichvalue_part: Option<String> = None;
    let mut rich_value_rel_part: Option<String> = None;

    for rel in relationships {
        if metadata_part.is_none() && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_SHEET_METADATA) {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            metadata_part = Some(target);
        } else if rdrichvalue_part.is_none() && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RD_RICH_VALUE)
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rdrichvalue_part = Some(target);
        } else if rich_value_rel_part.is_none()
            && rel.type_uri.eq_ignore_ascii_case(REL_TYPE_RICH_VALUE_REL)
        {
            let target = path::resolve_target("xl/workbook.xml", &rel.target);
            rich_value_rel_part = Some(target);
        }
    }

    let vm_to_rich_value_index = match metadata_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_vm_to_rich_value_index(bytes)?,
        None => HashMap::new(),
    };

    let rich_values = match rdrichvalue_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_rdrichvalue(bytes)?,
        None => HashMap::new(),
    };

    let local_image_identifier_to_rid = match rich_value_rel_part
        .as_deref()
        .and_then(|part| pkg.part(part).map(|bytes| (part, bytes)))
    {
        Some((_part, bytes)) => parse_rich_value_rel_ids(bytes)?,
        None => Vec::new(),
    };

    let rid_to_target: HashMap<String, String> = match rich_value_rel_part.as_deref() {
        Some(rich_value_rel_part) => {
            let rels_part = path::rels_for_part(rich_value_rel_part);
            let rels_bytes = match pkg.part(&rels_part) {
                Some(bytes) => bytes,
                None => &[],
            };

            if rels_bytes.is_empty() {
                HashMap::new()
            } else {
                let rels = openxml::parse_relationships(rels_bytes)?;
                rels.into_iter()
                    .filter(|rel| {
                        rel.type_uri.eq_ignore_ascii_case(REL_TYPE_IMAGE)
                            && !rel
                                .target_mode
                                .as_deref()
                                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    })
                    .map(|rel| {
                        let target = path::resolve_target(rich_value_rel_part, &rel.target);
                        (rel.id, target)
                    })
                    .collect()
            }
        }
        None => HashMap::new(),
    };

    let worksheet_parts: Vec<String> = match pkg.worksheet_parts() {
        Ok(infos) => infos.into_iter().map(|info| info.worksheet_part).collect(),
        Err(XlsxError::MissingPart(_))
        | Err(XlsxError::Invalid(_))
        | Err(XlsxError::InvalidSheetId) => pkg
            .part_names()
            .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
            .map(str::to_string)
            .collect(),
        Err(err) => return Err(err),
    };

    let mut out = Vec::new();
    for worksheet_part in worksheet_parts {
        let sheet_bytes = match pkg.part(&worksheet_part) {
            Some(bytes) => bytes,
            None => continue,
        };

        let mut reader = Reader::from_reader(Cursor::new(sheet_bytes));
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf)? {
                Event::Eof => break,
                Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                    let mut cell_ref: Option<CellRef> = None;
                    let mut vm: Option<u32> = None;

                    for attr in e.attributes() {
                        let attr = attr?;
                        let value = attr.unescape_value()?.into_owned();
                        match attr.key.as_ref() {
                            b"r" => {
                                if let Ok(parsed) = CellRef::from_a1(&value) {
                                    cell_ref = Some(parsed);
                                }
                            }
                            b"vm" => {
                                if let Ok(parsed) = value.parse::<u32>() {
                                    vm = Some(parsed);
                                }
                            }
                            _ => {}
                        }
                    }

                    let Some(cell) = cell_ref else { continue };
                    let Some(vm) = vm else { continue };

                    let Some(rich_value_index) = vm_to_rich_value_index.get(&vm).copied() else {
                        continue;
                    };
                    let Some(rich_value) = rich_values.get(&rich_value_index) else {
                        continue;
                    };

                    let rid = match local_image_identifier_to_rid
                        .get(rich_value.local_image_identifier as usize)
                    {
                        Some(rid) => rid,
                        None => continue,
                    };

                    let target = match rid_to_target.get(rid) {
                        Some(target) => target,
                        None => continue,
                    };

                    let bytes = match pkg.part(target) {
                        Some(bytes) => bytes.to_vec(),
                        None => continue,
                    };

                    out.push(EmbeddedImageCell {
                        sheet_part: worksheet_part.clone(),
                        cell,
                        image_target: target.clone(),
                        bytes,
                        alt_text: rich_value.alt_text.clone(),
                        decorative: rich_value.calc_origin == 5,
                    });
                }
                _ => {}
            }
            buf.clear();
        }
    }

    Ok(out)
}

fn parse_vm_to_rich_value_index(xml: &[u8]) -> Result<HashMap<u32, u32>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut xlr_type_index: Option<u32> = None;
    // `<rc t="...">` uses a **1-based** index into the `<metadataTypes>` list.
    let mut next_metadata_type_index = 1u32;

    // `futureMetadata` bk entries in order: each entry contains `xlrd:rvb i="..."`
    let mut future_bk_rich_value_index: Vec<Option<u32>> = Vec::new();
    let mut in_future_xlrichvalue = false;
    let mut current_future_bk: Option<usize> = None;

    // `valueMetadata` bk entries in order: each contains `rc t="..." v="..."`
    let mut value_bk_rc_records: Vec<Vec<(u32, u32)>> = Vec::new();
    let mut in_value_metadata = false;
    let mut current_value_bk: Option<usize> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let name = e.local_name();
                let name = name.as_ref();

                if name.eq_ignore_ascii_case(b"metadataType") {
                    let mut type_name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"name") {
                            type_name = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }

                    if type_name
                        .as_deref()
                        .is_some_and(|v| v.trim().eq_ignore_ascii_case("XLRICHVALUE"))
                    {
                        xlr_type_index = Some(next_metadata_type_index);
                    }
                    next_metadata_type_index = next_metadata_type_index.saturating_add(1);
                } else if name.eq_ignore_ascii_case(b"futureMetadata") {
                    let mut future_name: Option<String> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"name") {
                            future_name = Some(attr.unescape_value()?.into_owned());
                            break;
                        }
                    }
                    in_future_xlrichvalue = future_name
                        .as_deref()
                        .is_some_and(|v| v.trim().eq_ignore_ascii_case("XLRICHVALUE"));
                    current_future_bk = None;
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"bk") {
                    future_bk_rich_value_index.push(None);
                    current_future_bk = Some(future_bk_rich_value_index.len() - 1);
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"rvb") {
                    let mut i_value: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"i") {
                            i_value = attr
                                .unescape_value()?
                                .parse::<u32>()
                                .ok();
                            break;
                        }
                    }
                    if let (Some(idx), Some(i_value)) = (current_future_bk, i_value) {
                        if future_bk_rich_value_index.get(idx).is_some_and(Option::is_none) {
                            future_bk_rich_value_index[idx] = Some(i_value);
                        }
                    }
                } else if name.eq_ignore_ascii_case(b"valueMetadata") {
                    in_value_metadata = true;
                    current_value_bk = None;
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"bk") {
                    value_bk_rc_records.push(Vec::new());
                    current_value_bk = Some(value_bk_rc_records.len() - 1);
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"rc") {
                    let mut t: Option<u32> = None;
                    let mut v: Option<u32> = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = openxml::local_name(attr.key.as_ref());
                        if key.eq_ignore_ascii_case(b"t") {
                            t = attr.unescape_value()?.parse::<u32>().ok();
                        } else if key.eq_ignore_ascii_case(b"v") {
                            v = attr.unescape_value()?.parse::<u32>().ok();
                        }
                    }

                    if let (Some(current), Some(t), Some(v)) = (current_value_bk, t, v) {
                        if let Some(list) = value_bk_rc_records.get_mut(current) {
                            list.push((t, v));
                        }
                    }
                }
            }
            Event::End(e) => {
                let name = e.local_name();
                let name = name.as_ref();

                if name.eq_ignore_ascii_case(b"futureMetadata") {
                    in_future_xlrichvalue = false;
                    current_future_bk = None;
                } else if in_future_xlrichvalue && name.eq_ignore_ascii_case(b"bk") {
                    current_future_bk = None;
                } else if name.eq_ignore_ascii_case(b"valueMetadata") {
                    in_value_metadata = false;
                    current_value_bk = None;
                } else if in_value_metadata && name.eq_ignore_ascii_case(b"bk") {
                    current_value_bk = None;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    let Some(xlr_type_index) = xlr_type_index else {
        return Ok(HashMap::new());
    };

    let mut vm_to_rich_value_index: HashMap<u32, u32> = HashMap::new();
    for (bk_idx, rc_records) in value_bk_rc_records.iter().enumerate() {
        // Excel emits `rc/@t` (metadata type index) as either 0-based *or* 1-based depending on
        // version/producer. Prefer the 0-based index we derived while parsing `metadataTypes`, but
        // also accept the 1-based equivalent so we can resolve real Excel files.
        let xlr_type_index_0 = xlr_type_index;
        let xlr_type_index_1 = xlr_type_index.saturating_add(1);
        let Some(future_idx) = rc_records.iter().find_map(|(t, v)| {
            (*t == xlr_type_index_0 || *t == xlr_type_index_1).then_some(*v)
        }) else {
            continue;
        };

        let Some(Some(rich_value_index)) = future_bk_rich_value_index.get(future_idx as usize) else {
            continue;
        };

        // Worksheet `vm` attribute is 1-based.
        vm_to_rich_value_index.insert((bk_idx as u32) + 1, *rich_value_index);
    }

    Ok(vm_to_rich_value_index)
}

fn parse_rdrichvalue(xml: &[u8]) -> Result<HashMap<u32, RichValueImage>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_rv = false;
    let mut in_v = false;
    let mut current_v_text = String::new();
    let mut current_values: Vec<String> = Vec::new();
    let mut rich_value_index: u32 = 0;
    let mut out: HashMap<u32, RichValueImage> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                in_rv = true;
                in_v = false;
                current_v_text.clear();
                current_values.clear();
            }
            Event::Empty(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                // Empty rv elements aren't expected for local images.
                in_rv = false;
                in_v = false;
                current_v_text.clear();
                rich_value_index = rich_value_index.saturating_add(1);
            }
            Event::End(e) if e.local_name().as_ref().eq_ignore_ascii_case(b"rv") => {
                in_rv = false;
                in_v = false;
                current_v_text.clear();

                let local_image_identifier = current_values
                    .get(0)
                    .and_then(|v| v.trim().parse::<u32>().ok());
                let calc_origin = current_values
                    .get(1)
                    .and_then(|v| v.trim().parse::<u32>().ok());
                let alt_text = current_values
                    .get(2)
                    .map(|v| v.trim().to_string())
                    .filter(|v| !v.is_empty());

                if let (Some(local_image_identifier), Some(calc_origin)) =
                    (local_image_identifier, calc_origin)
                {
                    out.insert(
                        rich_value_index,
                        RichValueImage {
                            local_image_identifier,
                            calc_origin,
                            alt_text,
                        },
                    );
                }

                rich_value_index = rich_value_index.saturating_add(1);
            }
            Event::Start(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                in_v = true;
                current_v_text.clear();
            }
            Event::Empty(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                current_values.push(String::new());
                in_v = false;
                current_v_text.clear();
            }
            Event::End(e) if in_rv && e.local_name().as_ref().eq_ignore_ascii_case(b"v") => {
                current_values.push(std::mem::take(&mut current_v_text));
                in_v = false;
            }
            Event::Text(e) if in_rv && in_v => {
                current_v_text.push_str(&e.unescape()?.into_owned());
            }
            Event::CData(e) if in_rv && in_v => {
                current_v_text.push_str(
                    &e.decode()
                        .map_err(quick_xml::Error::from)?
                        .into_owned(),
                );
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_rich_value_rel_ids(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut out: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"rel" => {
                let mut rid: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                        rid = Some(attr.unescape_value()?.into_owned());
                        break;
                    }
                }
                if let Some(rid) = rid {
                    out.push(rid);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}
