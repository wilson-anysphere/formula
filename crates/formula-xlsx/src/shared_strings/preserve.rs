use std::borrow::Cow;
use std::collections::HashMap;
use std::io::BufRead;
use std::str::Utf8Error;

use crate::xml::{prefixed_tag, SPREADSHEETML_NS};
use formula_model::rich_text::{RichText, RichTextRunStyle, Underline};
use formula_model::Color;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::QName;
use quick_xml::{Reader, Writer};
use thiserror::Error;

/// A single `<si>...</si>` entry from `sharedStrings.xml`.
#[derive(Debug, Clone)]
pub(crate) struct SharedStringEntry {
    pub(crate) rich: RichText,
    /// The exact XML bytes for the `<si>...</si>` element as it appeared in the source file.
    pub(crate) raw_xml: Vec<u8>,
}

#[derive(Debug, Error)]
pub(crate) enum PreserveSharedStringsError {
    #[error("xml parse error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] Utf8Error),
    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),
    #[error("malformed sharedStrings.xml: {0}")]
    Malformed(&'static str),
}

/// Lossless editor for `xl/sharedStrings.xml` that preserves unknown XML subtrees.
///
/// The editor:
/// - Parses the original XML into `SharedStringEntry` records (decoded visible text + raw `<si>` bytes).
/// - Allows appending new `<si>` entries without rewriting the existing ones.
/// - Re-emits the original XML with updated `uniqueCount` (and optionally `count`).
#[derive(Debug, Clone)]
pub(crate) struct SharedStringsEditor {
    preamble: Vec<u8>,
    sst_start_tag: Vec<u8>,
    inner: Vec<u8>,
    suffix: Vec<u8>,
    insert_pos: usize,
    sst_tag_name: Vec<u8>,
    sst_was_self_closing: bool,
    original_count: Option<u32>,
    /// Prefix bound to the SpreadsheetML namespace. If SpreadsheetML is the default namespace,
    /// this will be `None`.
    spreadsheetml_prefix: Option<String>,
    /// Whether SpreadsheetML is the default namespace (`xmlns="…/main"`).
    spreadsheetml_is_default: bool,

    entries: Vec<SharedStringEntry>,
    appended: Vec<SharedStringEntry>,
    plain_index: HashMap<String, u32>,
    dirty: bool,
}

impl SharedStringsEditor {
    #[allow(dead_code)]
    pub(crate) fn new_empty() -> Self {
        let preamble = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
        let sst_start_tag =
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#.to_vec();
        let suffix = br#"</sst>"#.to_vec();
        Self {
            preamble,
            sst_start_tag,
            inner: Vec::new(),
            suffix,
            insert_pos: 0,
            sst_tag_name: b"sst".to_vec(),
            sst_was_self_closing: false,
            original_count: None,
            spreadsheetml_prefix: None,
            spreadsheetml_is_default: true,
            entries: Vec::new(),
            appended: Vec::new(),
            plain_index: HashMap::new(),
            dirty: false,
        }
    }

    pub(crate) fn parse(xml: &[u8]) -> Result<Self, PreserveSharedStringsError> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);

        let mut buf = Vec::new();

        let mut sst_tag_name: Option<Vec<u8>> = None;
        let mut sst_start_start: Option<usize> = None;
        let mut sst_start_end: Option<usize> = None;
        let mut sst_end_start: Option<usize> = None;
        let mut sst_was_self_closing = false;

        let mut original_count: Option<u32> = None;
        let mut spreadsheetml_prefix: Option<String> = None;
        let mut spreadsheetml_is_default: bool = false;

        let mut entries = Vec::new();
        let mut plain_index: HashMap<String, u32> = HashMap::new();

        let mut inside_sst = false;
        let mut last_si_end: Option<usize> = None;

        loop {
            let pos_before = usize_pos(reader.buffer_position())?;
            let event = reader.read_event_into(&mut buf)?;
            let pos_after = usize_pos(reader.buffer_position())?;

            match event {
                Event::Start(e) if local_name(e.name().as_ref()) == b"sst" => {
                    sst_tag_name = Some(e.name().as_ref().to_vec());
                    sst_start_start = Some(pos_before);
                    sst_start_end = Some(pos_after);
                    inside_sst = true;
                    original_count = parse_u32_attr(&e, b"count")?;
                    (spreadsheetml_is_default, spreadsheetml_prefix) =
                        spreadsheetml_namespace_style_from_sst_start(&e)?;
                }
                Event::Empty(e) if local_name(e.name().as_ref()) == b"sst" => {
                    // `<sst .../>` - treat as an empty document with no entries.
                    sst_tag_name = Some(e.name().as_ref().to_vec());
                    sst_start_start = Some(pos_before);
                    sst_start_end = Some(pos_after);
                    sst_end_start = Some(pos_after);
                    sst_was_self_closing = true;
                    inside_sst = false;
                    original_count = parse_u32_attr(&e, b"count")?;
                    (spreadsheetml_is_default, spreadsheetml_prefix) =
                        spreadsheetml_namespace_style_from_sst_start(&e)?;
                }
                Event::Start(e) if inside_sst && local_name(e.name().as_ref()) == b"si" => {
                    let si_start = pos_before;
                    let rich = parse_si_visible_text(&mut reader)?;
                    let si_end = usize_pos(reader.buffer_position())?;
                    let raw_xml = xml
                        .get(si_start..si_end)
                        .ok_or(PreserveSharedStringsError::Malformed("si range out of bounds"))?
                        .to_vec();

                    let idx = entries.len() as u32;
                    if rich.runs.is_empty() {
                        plain_index.entry(rich.text.clone()).or_insert(idx);
                    }
                    entries.push(SharedStringEntry { rich, raw_xml });
                    last_si_end = Some(si_end);
                }
                Event::Empty(e) if inside_sst && local_name(e.name().as_ref()) == b"si" => {
                    let raw_xml = xml
                        .get(pos_before..pos_after)
                        .ok_or(PreserveSharedStringsError::Malformed("si range out of bounds"))?
                        .to_vec();
                    let rich = RichText::new("");
                    let idx = entries.len() as u32;
                    plain_index.entry(rich.text.clone()).or_insert(idx);
                    entries.push(SharedStringEntry { rich, raw_xml });
                    last_si_end = Some(pos_after);
                }
                Event::End(e) if local_name(e.name().as_ref()) == b"sst" => {
                    sst_end_start = Some(pos_before);
                    inside_sst = false;
                }
                Event::Eof => break,
                _ => {}
            }

            buf.clear();
        }

        let sst_start_start =
            sst_start_start.ok_or(PreserveSharedStringsError::Malformed("missing <sst>"))?;
        let sst_start_end =
            sst_start_end.ok_or(PreserveSharedStringsError::Malformed("missing <sst>"))?;
        let sst_end_start =
            sst_end_start.ok_or(PreserveSharedStringsError::Malformed("missing </sst>"))?;
        let sst_tag_name =
            sst_tag_name.ok_or(PreserveSharedStringsError::Malformed("missing <sst>"))?;

        let preamble = xml[..sst_start_start].to_vec();
        let sst_start_tag = xml[sst_start_start..sst_start_end].to_vec();
        let inner = xml[sst_start_end..sst_end_start].to_vec();
        let suffix = xml[sst_end_start..].to_vec();

        let insert_pos_abs = last_si_end.unwrap_or(sst_start_end);
        let insert_pos = insert_pos_abs
            .checked_sub(sst_start_end)
            .ok_or(PreserveSharedStringsError::Malformed("invalid <sst> offsets"))?;

        Ok(Self {
            preamble,
            sst_start_tag,
            inner,
            suffix,
            insert_pos,
            sst_tag_name,
            sst_was_self_closing,
            original_count,
            spreadsheetml_prefix,
            spreadsheetml_is_default,
            entries,
            appended: Vec::new(),
            plain_index,
            dirty: false,
        })
    }

    pub(crate) fn is_dirty(&self) -> bool {
        self.dirty
    }

    pub(crate) fn len(&self) -> usize {
        self.entries.len() + self.appended.len()
    }

    pub(crate) fn original_count(&self) -> Option<u32> {
        self.original_count
    }

    pub(crate) fn rich_at(&self, idx: u32) -> Option<&RichText> {
        let idx = idx as usize;
        if idx < self.entries.len() {
            return Some(&self.entries[idx].rich);
        }
        self.appended
            .get(idx.saturating_sub(self.entries.len()))
            .map(|entry| &entry.rich)
    }

    fn spreadsheetml_prefix_for_inserts(&self) -> Option<&str> {
        if self.spreadsheetml_is_default {
            None
        } else {
            self.spreadsheetml_prefix.as_deref()
        }
    }

    pub(crate) fn get_or_insert_plain(&mut self, text: &str) -> u32 {
        if let Some(idx) = self.plain_index.get(text).copied() {
            return idx;
        }

        let idx = self.len() as u32;
        let rich = RichText::new(text.to_string());
        let raw_xml = write_si_xml(&rich, self.spreadsheetml_prefix_for_inserts());
        self.appended.push(SharedStringEntry { rich, raw_xml });
        self.plain_index.insert(text.to_string(), idx);
        self.dirty = true;
        idx
    }

    pub(crate) fn get_or_insert_rich(&mut self, rich: &RichText) -> u32 {
        // Note: rich text insertions are expected to be rare, so linear search is acceptable.
        for (idx, entry) in self
            .entries
            .iter()
            .chain(self.appended.iter())
            .enumerate()
        {
            if &entry.rich == rich {
                return idx as u32;
            }
        }

        let idx = self.len() as u32;
        let raw_xml = write_si_xml(rich, self.spreadsheetml_prefix_for_inserts());
        self.appended.push(SharedStringEntry {
            rich: rich.clone(),
            raw_xml,
        });
        self.dirty = true;
        idx
    }

    pub(crate) fn to_xml_bytes(
        &self,
        count_hint: Option<u32>,
    ) -> Result<Vec<u8>, PreserveSharedStringsError> {
        let unique_count = self.len() as u32;

        let appended_len = self
            .appended
            .iter()
            .try_fold(0usize, |acc, entry| acc.checked_add(entry.raw_xml.len()))
            .ok_or(PreserveSharedStringsError::Malformed(
                "appended shared strings XML is too large",
            ))?;
        let estimated_len = self
            .preamble
            .len()
            .checked_add(self.sst_start_tag.len())
            .and_then(|n| n.checked_add(self.inner.len()))
            .and_then(|n| n.checked_add(appended_len))
            .and_then(|n| n.checked_add(self.suffix.len()))
            .and_then(|n| n.checked_add(64))
            .ok_or(PreserveSharedStringsError::Malformed(
                "shared strings XML is too large",
            ))?;

        let mut out = Vec::new();
        out.try_reserve_exact(estimated_len)
            .map_err(|_| PreserveSharedStringsError::AllocationFailure("shared strings xml"))?;
        out.extend_from_slice(&self.preamble);

        let start_tag = patch_sst_start_tag(
            &self.sst_start_tag,
            count_hint,
            unique_count,
            self.sst_was_self_closing,
        )?;
        out.extend_from_slice(&start_tag);

        let (left, right) = self.inner.split_at(self.insert_pos.min(self.inner.len()));
        out.extend_from_slice(left);
        for entry in &self.appended {
            out.extend_from_slice(&entry.raw_xml);
        }
        out.extend_from_slice(right);

        if self.sst_was_self_closing {
            out.extend_from_slice(b"</");
            out.extend_from_slice(&self.sst_tag_name);
            out.extend_from_slice(b">");
        }
        out.extend_from_slice(&self.suffix);
        Ok(out)
    }
}

fn usize_pos(pos: u64) -> Result<usize, PreserveSharedStringsError> {
    usize::try_from(pos).map_err(|_| PreserveSharedStringsError::Malformed("xml too large"))
}

fn patch_sst_start_tag(
    tag: &[u8],
    count: Option<u32>,
    unique_count: u32,
    was_self_closing: bool,
) -> Result<Vec<u8>, PreserveSharedStringsError> {
    let mut tag = std::str::from_utf8(tag)?.to_string();

    // If the file used `<sst .../>`, convert to an open tag so we can append children.
    if was_self_closing {
        if let Some(idx) = tag.rfind("/>") {
            tag.replace_range(idx..idx + 2, ">");
        }
    }

    upsert_attr(&mut tag, "uniqueCount", &unique_count.to_string());
    if let Some(count) = count {
        upsert_attr(&mut tag, "count", &count.to_string());
    }

    Ok(tag.into_bytes())
}

fn upsert_attr(tag: &mut String, name: &str, value: &str) {
    if let Some((start, end)) = find_attr_value_range(tag.as_bytes(), name.as_bytes()) {
        tag.replace_range(start..end, value);
        return;
    }

    // Insert attribute before the closing `>`.
    if let Some(idx) = tag.rfind('>') {
        tag.insert_str(idx, &format!(" {name}=\"{value}\""));
    }
}

fn find_attr_value_range(tag: &[u8], name: &[u8]) -> Option<(usize, usize)> {
    // Scan for ` name="..."`, allowing arbitrary whitespace around the `=`.
    let mut i = 0usize;
    while i + name.len() + 3 < tag.len() {
        if tag[i..].starts_with(name) {
            // Require a whitespace boundary before the attribute name to avoid matching
            // substrings (e.g. `count` in `uniqueCount`).
            if i > 0 && !tag[i - 1].is_ascii_whitespace() {
                i += 1;
                continue;
            }

            let mut j = i + name.len();
            while j < tag.len() && tag[j].is_ascii_whitespace() {
                j += 1;
            }
            if tag.get(j) != Some(&b'=') {
                i += 1;
                continue;
            }
            j += 1;
            while j < tag.len() && tag[j].is_ascii_whitespace() {
                j += 1;
            }

            let quote = *tag.get(j)?;
            if quote != b'"' && quote != b'\'' {
                i += 1;
                continue;
            }

            let value_start = j + 1;
            let mut j = value_start;
            while j < tag.len() && tag[j] != quote {
                j += 1;
            }
            if j >= tag.len() {
                return None;
            }
            return Some((value_start, j));
        }
        i += 1;
    }
    None
}

fn parse_si_visible_text<R: BufRead>(
    reader: &mut Reader<R>,
) -> Result<RichText, PreserveSharedStringsError> {
    let mut buf = Vec::new();
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                let t = read_text(reader, QName(b"t"))?;
                segments.push((t, RichTextRunStyle::default()));
            }
            Event::Start(e) if e.local_name().as_ref() == b"r" => {
                segments.push(parse_r(reader)?);
            }
            Event::Start(e) => {
                // Skip subtrees that are not part of the visible string (phonetic/ruby, extLst,
                // unknown future extensions). This prevents misinterpreting `<t>` nodes inside
                // those structures as display text.
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"si" => break,
            Event::Eof => return Err(PreserveSharedStringsError::Malformed("unexpected EOF in <si>")),
            _ => {}
        }
        buf.clear();
    }

    if segments.iter().all(|(_, style)| style.is_empty()) {
        Ok(RichText::new(
            segments.into_iter().map(|(text, _)| text).collect::<String>(),
        ))
    } else {
        Ok(RichText::from_segments(segments))
    }
}

fn parse_r<R: BufRead>(
    reader: &mut Reader<R>,
) -> Result<(String, RichTextRunStyle), PreserveSharedStringsError> {
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
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"r" => break,
            Event::Eof => return Err(PreserveSharedStringsError::Malformed("unexpected EOF in <r>")),
            _ => {}
        }
        buf.clear();
    }

    Ok((text, style))
}

fn parse_rpr<R: BufRead>(
    reader: &mut Reader<R>,
) -> Result<RichTextRunStyle, PreserveSharedStringsError> {
    let mut buf = Vec::new();
    let mut style = RichTextRunStyle::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) => parse_rpr_tag(&e, &mut style)?,
            Event::Start(e) => {
                parse_rpr_tag(&e, &mut style)?;
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"rPr" => break,
            Event::Eof => return Err(PreserveSharedStringsError::Malformed("unexpected EOF in <rPr>")),
            _ => {}
        }
        buf.clear();
    }

    Ok(style)
}

fn parse_rpr_tag(
    e: &BytesStart<'_>,
    style: &mut RichTextRunStyle,
) -> Result<(), PreserveSharedStringsError> {
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

fn parse_bool_val(e: &BytesStart<'_>) -> Result<bool, PreserveSharedStringsError> {
    let Some(val) = attr_value(e, b"val")? else {
        return Ok(true);
    };
    Ok(!(val == "0" || val.eq_ignore_ascii_case("false")))
}

fn read_text<R: BufRead>(
    reader: &mut Reader<R>,
    end: QName<'_>,
) -> Result<String, PreserveSharedStringsError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => {
                let t: Cow<'_, str> = e.unescape()?;
                text.push_str(&t);
            }
            Event::CData(e) => text.push_str(std::str::from_utf8(e.as_ref())?),
            Event::End(e) if local_name(e.name().as_ref()) == end.as_ref() => break,
            Event::Eof => return Err(PreserveSharedStringsError::Malformed("unexpected EOF in <t>")),
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn attr_value(
    e: &BytesStart<'_>,
    key: &[u8],
) -> Result<Option<String>, PreserveSharedStringsError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn parse_u32_attr(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<u32>, PreserveSharedStringsError> {
    Ok(attr_value(e, key)?.and_then(|s| s.parse::<u32>().ok()))
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

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|&b| b == b':').next().unwrap_or(name)
}

fn spreadsheetml_namespace_style_from_sst_start(
    e: &BytesStart<'_>,
) -> Result<(bool, Option<String>), PreserveSharedStringsError> {
    let mut spreadsheetml_is_default = false;
    let mut spreadsheetml_prefix_decl: Option<String> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let value = attr.value.as_ref();

        if key == b"xmlns" && value == SPREADSHEETML_NS.as_bytes() {
            spreadsheetml_is_default = true;
            continue;
        }

        if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            if value == SPREADSHEETML_NS.as_bytes() {
                spreadsheetml_prefix_decl = Some(String::from_utf8_lossy(prefix).into_owned());
            }
        }
    }

    if spreadsheetml_is_default {
        return Ok((true, None));
    }

    let name = e.name();
    let name = name.as_ref();
    let element_prefix = name
        .iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
        .map(|bytes| String::from_utf8_lossy(bytes).into_owned());

    // Prefer the prefix used by the `<sst>` element itself, if present. This preserves the source
    // file's namespace style when appending new `<si>` entries.
    let spreadsheetml_prefix = element_prefix.or(spreadsheetml_prefix_decl);

    Ok((false, spreadsheetml_prefix))
}

fn write_si_xml(item: &RichText, spreadsheetml_prefix: Option<&str>) -> Vec<u8> {
    let mut writer = Writer::new(Vec::new());
    if let Err(err) = write_si(&mut writer, spreadsheetml_prefix, item) {
        debug_assert!(false, "failed to write <si> xml: {err}");
    }
    writer.into_inner()
}

fn write_si(
    writer: &mut Writer<Vec<u8>>,
    spreadsheetml_prefix: Option<&str>,
    item: &RichText,
) -> Result<(), quick_xml::Error> {
    let si_name = prefixed_tag(spreadsheetml_prefix, "si");
    writer.write_event(Event::Start(BytesStart::new(si_name.as_str())))?;

    if item.runs.is_empty() {
        write_t(writer, spreadsheetml_prefix, &item.text)?;
    } else {
        let r_name = prefixed_tag(spreadsheetml_prefix, "r");
        let rpr_name = prefixed_tag(spreadsheetml_prefix, "rPr");
        for run in &item.runs {
            writer.write_event(Event::Start(BytesStart::new(r_name.as_str())))?;

            if !run.style.is_empty() {
                writer.write_event(Event::Start(BytesStart::new(rpr_name.as_str())))?;
                write_rpr(writer, spreadsheetml_prefix, &run.style)?;
                writer.write_event(Event::End(BytesEnd::new(rpr_name.as_str())))?;
            }

            let segment = item.slice_run_text(run);
            write_t(writer, spreadsheetml_prefix, segment)?;

            writer.write_event(Event::End(BytesEnd::new(r_name.as_str())))?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new(si_name.as_str())))?;
    Ok(())
}

fn write_t(
    writer: &mut Writer<Vec<u8>>,
    spreadsheetml_prefix: Option<&str>,
    text: &str,
) -> Result<(), quick_xml::Error> {
    let t_name = prefixed_tag(spreadsheetml_prefix, "t");
    let mut t = BytesStart::new(t_name.as_str());
    if needs_space_preserve(text) {
        t.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(t))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(t_name.as_str())))?;
    Ok(())
}

fn write_rpr(
    writer: &mut Writer<Vec<u8>>,
    spreadsheetml_prefix: Option<&str>,
    style: &RichTextRunStyle,
) -> Result<(), quick_xml::Error> {
    if let Some(font) = &style.font {
        let rfont_name = prefixed_tag(spreadsheetml_prefix, "rFont");
        let mut rfont = BytesStart::new(rfont_name.as_str());
        rfont.push_attribute(("val", font.as_str()));
        writer.write_event(Event::Empty(rfont))?;
    }

    if let Some(size_100pt) = style.size_100pt {
        let sz_name = prefixed_tag(spreadsheetml_prefix, "sz");
        let mut sz = BytesStart::new(sz_name.as_str());
        let value = format_size_100pt(size_100pt);
        sz.push_attribute(("val", value.as_str()));
        writer.write_event(Event::Empty(sz))?;
    }

    if let Some(color) = style.color {
        let c_name = prefixed_tag(spreadsheetml_prefix, "color");
        let mut c = BytesStart::new(c_name.as_str());
        let value = format!("{:08X}", color.argb().unwrap_or(0));
        c.push_attribute(("rgb", value.as_str()));
        writer.write_event(Event::Empty(c))?;
    }

    if let Some(bold) = style.bold {
        let b_name = prefixed_tag(spreadsheetml_prefix, "b");
        let mut b = BytesStart::new(b_name.as_str());
        if !bold {
            b.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(b))?;
    }

    if let Some(italic) = style.italic {
        let i_name = prefixed_tag(spreadsheetml_prefix, "i");
        let mut i = BytesStart::new(i_name.as_str());
        if !italic {
            i.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(i))?;
    }

    if let Some(ul) = style.underline {
        let u_name = prefixed_tag(spreadsheetml_prefix, "u");
        let mut u = BytesStart::new(u_name.as_str());
        if let Some(val) = ul.to_ooxml() {
            u.push_attribute(("val", val));
        }
        writer.write_event(Event::Empty(u))?;
    }

    Ok(())
}

fn needs_space_preserve(text: &str) -> bool {
    text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace)
}

fn format_size_100pt(size_100pt: u16) -> String {
    let int = size_100pt / 100;
    let frac = size_100pt % 100;
    if frac == 0 {
        return int.to_string();
    }

    let mut s = format!("{int}.{frac:02}");
    while s.ends_with('0') {
        s.pop();
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;
    use roxmltree::Document;

    #[test]
    fn appends_si_entries_preserving_prefix_only_spreadsheetml() {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:sst xmlns:x="{ns}" count="1" uniqueCount="1"><x:si><x:t>foo</x:t></x:si></x:sst>"#,
            ns = SPREADSHEETML_NS
        );

        let mut editor = SharedStringsEditor::parse(xml.as_bytes()).unwrap();
        editor.get_or_insert_plain("bar");
        let updated = editor.to_xml_bytes(None).unwrap();

        let updated_str = std::str::from_utf8(&updated).unwrap();
        let doc = Document::parse(updated_str).unwrap();

        assert!(updated_str.contains("<x:si"));
        assert!(updated_str.contains("<x:t>bar</x:t>"));
        assert!(!updated_str.contains("<si"));
        assert!(!updated_str.contains("<t"));

        let bar_t = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "t" && n.text() == Some("bar"))
            .unwrap();
        let bar_si = bar_t
            .ancestors()
            .find(|n| n.is_element() && n.tag_name().name() == "si")
            .unwrap();
        assert_eq!(bar_si.tag_name().namespace(), Some(SPREADSHEETML_NS));
    }

    #[test]
    fn appends_si_entries_default_namespace_keeps_unprefixed_tags() {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{ns}" count="1" uniqueCount="1"><si><t>foo</t></si></sst>"#,
            ns = SPREADSHEETML_NS
        );

        let mut editor = SharedStringsEditor::parse(xml.as_bytes()).unwrap();
        editor.get_or_insert_plain("bar");
        let updated = editor.to_xml_bytes(None).unwrap();

        let updated_str = std::str::from_utf8(&updated).unwrap();
        let doc = Document::parse(updated_str).unwrap();

        assert!(updated_str.contains("<si"));
        assert!(updated_str.contains("<t>bar</t>"));
        assert!(!updated_str.contains("<x:si"));
        assert!(!updated_str.contains("<x:t"));

        let bar_t = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "t" && n.text() == Some("bar"))
            .unwrap();
        let bar_si = bar_t
            .ancestors()
            .find(|n| n.is_element() && n.tag_name().name() == "si")
            .unwrap();
        assert_eq!(bar_si.tag_name().namespace(), Some(SPREADSHEETML_NS));
    }
}
