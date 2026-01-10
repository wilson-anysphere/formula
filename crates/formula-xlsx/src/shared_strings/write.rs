use formula_model::rich_text::{RichText, RichTextRunStyle};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use thiserror::Error;

use super::SharedStrings;

const OOXML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

#[derive(Debug, Error)]
pub enum WriteSharedStringsError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
}

pub fn write_shared_strings_xml(
    shared_strings: &SharedStrings,
) -> Result<String, WriteSharedStringsError> {
    let mut writer = Writer::new(Vec::new());

    writer.write_event(Event::Decl(BytesDecl::new(
        "1.0",
        Some("UTF-8"),
        Some("yes"),
    )))?;

    let mut sst = BytesStart::new("sst");
    sst.push_attribute(("xmlns", OOXML_NS));
    let count = shared_strings.items.len().to_string();
    sst.push_attribute(("count", count.as_str()));
    sst.push_attribute(("uniqueCount", count.as_str()));
    writer.write_event(Event::Start(sst))?;

    for item in &shared_strings.items {
        write_si(&mut writer, item)?;
    }

    writer.write_event(Event::End(BytesEnd::new("sst")))?;
    Ok(String::from_utf8(writer.into_inner())?)
}

fn write_si(writer: &mut Writer<Vec<u8>>, item: &RichText) -> std::io::Result<()> {
    writer.write_event(Event::Start(BytesStart::new("si")))?;

    if item.runs.is_empty() {
        write_t(writer, &item.text)?;
    } else {
        for run in &item.runs {
            writer.write_event(Event::Start(BytesStart::new("r")))?;

            if !run.style.is_empty() {
                writer.write_event(Event::Start(BytesStart::new("rPr")))?;
                write_rpr(writer, &run.style)?;
                writer.write_event(Event::End(BytesEnd::new("rPr")))?;
            }

            let segment = item.slice_run_text(run);
            write_t(writer, segment)?;

            writer.write_event(Event::End(BytesEnd::new("r")))?;
        }
    }

    writer.write_event(Event::End(BytesEnd::new("si")))?;
    Ok(())
}

fn write_t(writer: &mut Writer<Vec<u8>>, text: &str) -> std::io::Result<()> {
    let mut t = BytesStart::new("t");
    if needs_space_preserve(text) {
        t.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(t))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new("t")))?;
    Ok(())
}

fn write_rpr(
    writer: &mut Writer<Vec<u8>>,
    style: &RichTextRunStyle,
) -> std::io::Result<()> {
    if let Some(font) = &style.font {
        let mut rfont = BytesStart::new("rFont");
        rfont.push_attribute(("val", font.as_str()));
        writer.write_event(Event::Empty(rfont))?;
    }

    if let Some(size_100pt) = style.size_100pt {
        let mut sz = BytesStart::new("sz");
        let value = format_size_100pt(size_100pt);
        sz.push_attribute(("val", value.as_str()));
        writer.write_event(Event::Empty(sz))?;
    }

    if let Some(color) = style.color {
        let mut c = BytesStart::new("color");
        let value = format!("{:08X}", color.argb);
        c.push_attribute(("rgb", value.as_str()));
        writer.write_event(Event::Empty(c))?;
    }

    if let Some(bold) = style.bold {
        let mut b = BytesStart::new("b");
        if !bold {
            b.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(b))?;
    }

    if let Some(italic) = style.italic {
        let mut i = BytesStart::new("i");
        if !italic {
            i.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(i))?;
    }

    if let Some(ul) = style.underline {
        let mut u = BytesStart::new("u");
        if let Some(val) = ul.to_ooxml() {
            u.push_attribute(("val", val));
        }
        writer.write_event(Event::Empty(u))?;
    }

    Ok(())
}

fn needs_space_preserve(text: &str) -> bool {
    text.starts_with(' ') || text.ends_with(' ')
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
