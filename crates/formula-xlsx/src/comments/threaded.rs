use formula_model::{CellRef, Comment, CommentAuthor, CommentKind, Reply};

use std::collections::BTreeMap;

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ThreadedParseError {
    Utf8(std::str::Utf8Error),
}

impl From<std::str::Utf8Error> for ThreadedParseError {
    fn from(value: std::str::Utf8Error) -> Self {
        Self::Utf8(value)
    }
}

#[derive(Clone, Debug)]
struct RawThreadedComment {
    id: String,
    parent_id: Option<String>,
    cell_ref: CellRef,
    author_id: String,
    author_name: String,
    created_at: i64,
    content: String,
    resolved: bool,
}

pub fn parse_threaded_comments_xml(
    bytes: &[u8],
    persons: &BTreeMap<String, String>,
) -> Result<Vec<Comment>, ThreadedParseError> {
    let xml = std::str::from_utf8(bytes)?;

    let mut raw_comments = Vec::new();
    let mut cursor = 0usize;
    while let Some(open_rel) = xml[cursor..].find("<threadedComment") {
        let open_pos = cursor + open_rel;
        let close_pos = match xml[open_pos..].find("</threadedComment>") {
            Some(rel) => open_pos + rel + "</threadedComment>".len(),
            None => break,
        };

        let comment_xml = &xml[open_pos..close_pos];
        let id = extract_attr(comment_xml, "id")
            .or_else(|| extract_attr(comment_xml, "guid"))
            .unwrap_or_else(|| format!("threaded:{}", raw_comments.len()));
        let parent_id = extract_attr(comment_xml, "parentId");
        let cell_ref = extract_attr(comment_xml, "ref")
            .and_then(|a1| CellRef::from_a1(&a1).ok());
        let Some(cell_ref) = cell_ref else {
            cursor = close_pos;
            continue;
        };
        let author_id = extract_attr(comment_xml, "personId")
            .or_else(|| extract_attr(comment_xml, "authorId"))
            .unwrap_or_default();
        let author_name = extract_attr(comment_xml, "author")
            .or_else(|| extract_attr(comment_xml, "displayName"))
            .or_else(|| persons.get(&author_id).cloned())
            .unwrap_or_else(|| author_id.clone());
        let created_at = extract_attr(comment_xml, "dT")
            .and_then(|dt| parse_iso8601_ms(&dt))
            .unwrap_or(0);
        let resolved = extract_attr(comment_xml, "done")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .or_else(|| {
                extract_attr(comment_xml, "resolved").map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            })
            .unwrap_or(false);

        let content = extract_all_tag_text(comment_xml, "t")
            .into_iter()
            .map(xml_unescape)
            .collect::<Vec<_>>()
            .join("");

        raw_comments.push(RawThreadedComment {
            id,
            parent_id,
            cell_ref,
            author_id,
            author_name,
            created_at,
            content,
            resolved,
        });

        cursor = close_pos;
    }

    let mut by_id = std::collections::BTreeMap::<String, RawThreadedComment>::new();
    for raw in raw_comments {
        by_id.insert(raw.id.clone(), raw);
    }

    let mut roots = Vec::new();
    for raw in by_id.values() {
        if raw.parent_id.is_none() {
            roots.push(raw.clone());
        }
    }

    let mut out = Vec::new();
    for root in roots {
        let mut replies = Vec::new();
        for raw in by_id.values() {
            if raw.parent_id.as_deref() == Some(&root.id) {
                replies.push(Reply {
                    id: raw.id.clone(),
                    author: CommentAuthor {
                        id: raw.author_id.clone(),
                        name: raw.author_name.clone(),
                    },
                    created_at: raw.created_at,
                    updated_at: raw.created_at,
                    content: raw.content.clone(),
                    mentions: Vec::new(),
                });
            }
        }

        out.push(Comment {
            id: root.id.clone(),
            cell_ref: root.cell_ref,
            author: CommentAuthor {
                id: root.author_id.clone(),
                name: root.author_name.clone(),
            },
            created_at: root.created_at,
            updated_at: root.created_at,
            resolved: root.resolved,
            kind: CommentKind::Threaded,
            content: root.content.clone(),
            mentions: Vec::new(),
            replies,
        });
    }

    Ok(out)
}

pub fn write_threaded_comments_xml(comments: &[Comment]) -> Vec<u8> {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(
        r#"<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">"#,
    );
    xml.push('\n');

    for comment in comments {
        xml.push_str("  <threadedComment id=\"");
        xml.push_str(&xml_escape(&comment.id));
        xml.push_str("\" ref=\"");
        formula_model::push_a1_cell_ref(comment.cell_ref.row, comment.cell_ref.col, false, false, &mut xml);
        xml.push_str("\" personId=\"");
        xml.push_str(&xml_escape(&comment.author.id));
        xml.push_str("\" author=\"");
        xml.push_str(&xml_escape(&comment.author.name));
        xml.push_str("\" done=\"");
        xml.push_str(if comment.resolved { "1" } else { "0" });
        xml.push_str("\">\n");
        xml.push_str("    <text><r><t xml:space=\"preserve\">");
        xml.push_str(&xml_escape(&comment.content));
        xml.push_str("</t></r></text>\n");
        xml.push_str("  </threadedComment>\n");

        for reply in &comment.replies {
            xml.push_str("  <threadedComment id=\"");
            xml.push_str(&xml_escape(&reply.id));
            xml.push_str("\" parentId=\"");
            xml.push_str(&xml_escape(&comment.id));
            xml.push_str("\" ref=\"");
            formula_model::push_a1_cell_ref(comment.cell_ref.row, comment.cell_ref.col, false, false, &mut xml);
            xml.push_str("\" personId=\"");
            xml.push_str(&xml_escape(&reply.author.id));
            xml.push_str("\" author=\"");
            xml.push_str(&xml_escape(&reply.author.name));
            xml.push_str("\">\n");
            xml.push_str("    <text><r><t xml:space=\"preserve\">");
            xml.push_str(&xml_escape(&reply.content));
            xml.push_str("</t></r></text>\n");
            xml.push_str("  </threadedComment>\n");
        }
    }

    xml.push_str("</threadedComments>\n");
    xml.into_bytes()
}

fn extract_all_tag_text(xml: &str, tag: &str) -> Vec<String> {
    let mut out = Vec::new();
    let open_pat = format!("<{}", tag);
    let close_pat = format!("</{}>", tag);

    let mut cursor = 0usize;
    while let Some(open_pos_rel) = xml[cursor..].find(&open_pat) {
        let open_pos = cursor + open_pos_rel;
        // Ensure we match the full tag name, not a prefix (e.g. `<t>` vs `<text>`).
        let Some(next) = xml.as_bytes().get(open_pos + open_pat.len()) else {
            break;
        };
        if !matches!(*next, b'>' | b'/' | b' ' | b'\t' | b'\r' | b'\n') {
            cursor = open_pos + open_pat.len();
            continue;
        }
        let open_end = match xml[open_pos..].find('>') {
            Some(end) => open_pos + end + 1,
            None => break,
        };
        let close_pos = match xml[open_end..].find(&close_pat) {
            Some(rel) => open_end + rel,
            None => break,
        };

        out.push(xml[open_end..close_pos].to_string());
        cursor = close_pos + close_pat.len();
    }

    out
}

fn extract_attr(xml: &str, attr: &str) -> Option<String> {
    let pattern = format!(r#"{attr}=""#);
    let start = xml.find(&pattern)? + pattern.len();
    let end = xml[start..].find('"')? + start;
    Some(xml[start..end].to_string())
}

fn xml_unescape(value: String) -> String {
    value
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn parse_iso8601_ms(value: &str) -> Option<i64> {
    // Minimal ISO8601 parsing: YYYY-MM-DDTHH:MM:SS(.mmm)Z
    // We only need stable ordering / basic timestamps; return None on unexpected formats.
    let value = value.strip_suffix('Z')?;
    let (date, time) = value.split_once('T')?;
    let (year, rest) = date.split_once('-')?;
    let (month, day) = rest.split_once('-')?;

    let (time_part, millis_part) = match time.split_once('.') {
        Some((t, ms)) => (t, Some(ms)),
        None => (time, None),
    };
    let (hour, rest) = time_part.split_once(':')?;
    let (minute, second) = rest.split_once(':')?;

    let year: i32 = year.parse().ok()?;
    let month: u32 = month.parse().ok()?;
    let day: u32 = day.parse().ok()?;
    let hour: u32 = hour.parse().ok()?;
    let minute: u32 = minute.parse().ok()?;
    let second: u32 = second.parse().ok()?;
    let millis: u32 = millis_part
        .and_then(|ms| ms.get(0..3))
        .and_then(|ms| ms.parse().ok())
        .unwrap_or(0);

    datetime_to_ms(year, month, day, hour, minute, second, millis)
}

fn datetime_to_ms(
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    millis: u32,
) -> Option<i64> {
    let days = days_since_epoch(year, month, day)?;
    let seconds = days * 86_400 + hour as i64 * 3_600 + minute as i64 * 60 + second as i64;
    Some(seconds * 1_000 + millis as i64)
}

fn days_since_epoch(year: i32, month: u32, day: u32) -> Option<i64> {
    // Unix epoch: 1970-01-01
    if month == 0 || month > 12 || day == 0 || day > 31 {
        return None;
    }
    let mut days = 0i64;
    if year >= 1970 {
        for y in 1970..year {
            days += if is_leap_year(y) { 366 } else { 365 };
        }
    } else {
        for y in (year..1970).rev() {
            days -= if is_leap_year(y) { 366 } else { 365 };
        }
    }

    let month_lengths = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    for m in 1..month {
        let mut len = month_lengths[(m - 1) as usize];
        if m == 2 && is_leap_year(year) {
            len = 29;
        }
        days += len as i64;
    }

    days += (day as i64) - 1;
    Some(days)
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}
