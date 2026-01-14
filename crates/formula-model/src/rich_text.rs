use serde::{Deserialize, Serialize};

use crate::Color;

/// Rich (multi-style) text representation.
///
/// The `text` field contains the full string content. `runs` contains style
/// overrides applied to ranges in `text`.
///
/// ## Indexing
/// Run `start`/`end` offsets are **Unicode scalar value** (`char`) indices into
/// `text` (not UTF-8 byte offsets). This makes indices stable across UTF-8
/// encodings but still does not correspond to user-perceived grapheme clusters.
#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct RichText {
    pub text: String,
    pub runs: Vec<RichTextRun>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phonetic: Option<String>,
}

impl RichText {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            runs: Vec::new(),
            phonetic: None,
        }
    }

    pub fn plain_text(&self) -> &str {
        &self.text
    }

    pub fn is_plain(&self) -> bool {
        self.runs.is_empty()
    }

    pub fn char_len(&self) -> usize {
        self.text.chars().count()
    }

    pub fn from_segments(segments: impl IntoIterator<Item = (String, RichTextRunStyle)>) -> Self {
        let mut text = String::new();
        let mut runs = Vec::new();
        let mut cursor = 0usize;

        for (segment_text, style) in segments {
            let start = cursor;
            cursor += segment_text.chars().count();
            let end = cursor;
            text.push_str(&segment_text);
            runs.push(RichTextRun { start, end, style });
        }

        Self {
            text,
            runs,
            phonetic: None,
        }
    }

    pub fn slice_run_text(&self, run: &RichTextRun) -> &str {
        slice_by_char_range(&self.text, run.start, run.end)
    }
}

impl PartialEq for RichText {
    fn eq(&self, other: &Self) -> bool {
        // Preserve historical RichText semantics: equality is based on the visible
        // text + style runs only.
        self.text == other.text && self.runs == other.runs
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct RichTextRun {
    pub start: usize,
    pub end: usize,
    pub style: RichTextRunStyle,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct RichTextRunStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<Underline>,
    pub color: Option<Color>,
    pub font: Option<String>,
    /// Font size in 1/100 points (e.g. 1100 = 11pt), matching [`crate::Font`].
    pub size_100pt: Option<u16>,
}

impl RichTextRunStyle {
    pub fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.color.is_none()
            && self.font.is_none()
            && self.size_100pt.is_none()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Underline {
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
    None,
}

impl Underline {
    pub fn from_ooxml(val: Option<&str>) -> Option<Self> {
        match val {
            None => Some(Underline::Single),
            Some("single") => Some(Underline::Single),
            Some("double") => Some(Underline::Double),
            Some("singleAccounting") => Some(Underline::SingleAccounting),
            Some("doubleAccounting") => Some(Underline::DoubleAccounting),
            Some("none") => Some(Underline::None),
            _ => None,
        }
    }

    pub fn to_ooxml(self) -> Option<&'static str> {
        match self {
            Underline::Single => None,
            Underline::Double => Some("double"),
            Underline::SingleAccounting => Some("singleAccounting"),
            Underline::DoubleAccounting => Some("doubleAccounting"),
            Underline::None => Some("none"),
        }
    }
}

fn slice_by_char_range(text: &str, start: usize, end: usize) -> &str {
    if start == end {
        return "";
    }

    let mut start_byte = None;
    let mut end_byte = None;

    for (i, (byte_idx, _ch)) in text.char_indices().enumerate() {
        if i == start {
            start_byte = Some(byte_idx);
        }
        if i == end {
            end_byte = Some(byte_idx);
            break;
        }
    }

    let start_byte = start_byte.unwrap_or_else(|| text.len());
    let end_byte = end_byte.unwrap_or_else(|| text.len());

    &text[start_byte..end_byte]
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn from_segments_builds_runs_with_char_indices() {
        let rt = RichText::from_segments(vec![
            ("Hi ".to_string(), RichTextRunStyle::default()),
            (
                "世界".to_string(),
                RichTextRunStyle {
                    bold: Some(true),
                    ..Default::default()
                },
            ),
        ]);

        assert_eq!(rt.text, "Hi 世界");
        assert_eq!(rt.runs.len(), 2);
        assert_eq!(rt.runs[0].start, 0);
        assert_eq!(rt.runs[0].end, 3);
        assert_eq!(rt.runs[1].start, 3);
        assert_eq!(rt.runs[1].end, 5);
        assert_eq!(rt.slice_run_text(&rt.runs[1]), "世界");
    }
}
