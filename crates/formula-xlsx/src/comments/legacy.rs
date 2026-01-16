use core::fmt::Write as _;

use formula_model::{CellRef, Comment, CommentAuthor, CommentKind};

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum LegacyParseError {
    Utf8(std::str::Utf8Error),
}

impl From<std::str::Utf8Error> for LegacyParseError {
    fn from(value: std::str::Utf8Error) -> Self {
        Self::Utf8(value)
    }
}

pub fn parse_comments_xml(bytes: &[u8]) -> Result<Vec<Comment>, LegacyParseError> {
    let xml = std::str::from_utf8(bytes)?;

    let authors = extract_all_tag_text(xml, "author")
        .into_iter()
        .map(xml_unescape)
        .collect::<Vec<_>>();

    let mut comments = Vec::new();
    let mut cursor = 0usize;
    while let Some(start) = xml[cursor..].find("<comment") {
        let start = cursor + start;
        let end = match xml[start..].find("</comment>") {
            Some(end_rel) => start + end_rel + "</comment>".len(),
            None => break,
        };

        let comment_xml = &xml[start..end];
        let cell_ref = extract_attr(comment_xml, "ref")
            .and_then(|a1| CellRef::from_a1(&a1).ok());
        let Some(cell_ref) = cell_ref else {
            cursor = end;
            continue;
        };
        let author_id = extract_attr(comment_xml, "authorId").unwrap_or_else(|| "0".to_string());
        let author_name = author_id
            .parse::<usize>()
            .ok()
            .and_then(|idx| authors.get(idx).cloned())
            .unwrap_or_else(|| author_id.clone());

        let text = extract_all_tag_text(comment_xml, "t")
            .into_iter()
            .map(xml_unescape)
            .collect::<Vec<_>>()
            .join("");

        let mut id = String::new();
        id.push_str("note:");
        formula_model::push_a1_cell_ref(cell_ref.row, cell_ref.col, false, false, &mut id);
        id.push(':');
        let _ = write!(id, "{}", comments.len());
        comments.push(Comment {
            id,
            cell_ref,
            author: CommentAuthor {
                id: author_id,
                name: author_name,
            },
            created_at: 0,
            updated_at: 0,
            resolved: false,
            kind: CommentKind::Note,
            content: text,
            mentions: Vec::new(),
            replies: Vec::new(),
        });

        cursor = end;
    }

    Ok(comments)
}

pub fn write_comments_xml(comments: &[Comment]) -> Vec<u8> {
    let mut authors = Vec::new();
    let mut author_index = std::collections::BTreeMap::<String, usize>::new();

    for comment in comments {
        if comment.author.name.is_empty() {
            continue;
        }
        if !author_index.contains_key(&comment.author.name) {
            let idx = authors.len();
            authors.push(comment.author.name.clone());
            author_index.insert(comment.author.name.clone(), idx);
        }
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push('\n');
    xml.push_str(r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#);
    xml.push('\n');
    xml.push_str("  <authors>\n");
    for author in &authors {
        xml.push_str("    <author>");
        xml.push_str(&xml_escape(author));
        xml.push_str("</author>\n");
    }
    xml.push_str("  </authors>\n");
    xml.push_str("  <commentList>\n");

    for comment in comments {
        let author_id = author_index
            .get(&comment.author.name)
            .copied()
            .unwrap_or(0);

        xml.push_str("    <comment ref=\"");
        formula_model::push_a1_cell_ref(comment.cell_ref.row, comment.cell_ref.col, false, false, &mut xml);
        xml.push_str("\" authorId=\"");
        xml.push_str(&author_id.to_string());
        xml.push_str("\">\n");
        xml.push_str("      <text><r><t xml:space=\"preserve\">");
        xml.push_str(&xml_escape(&comment.content));
        xml.push_str("</t></r></text>\n");
        xml.push_str("    </comment>\n");
    }

    xml.push_str("  </commentList>\n");
    xml.push_str("</comments>\n");
    xml.into_bytes()
}

pub fn parse_vml_drawing_cells(bytes: &[u8]) -> Result<Vec<CellRef>, LegacyParseError> {
    let xml = std::str::from_utf8(bytes)?;
    let mut out = Vec::new();

    let mut cursor = 0usize;
    while let Some(client_rel) = xml[cursor..].find("<x:ClientData") {
        let client_start = cursor + client_rel;
        let client_end = match xml[client_start..].find("</x:ClientData>") {
            Some(rel) => client_start + rel + "</x:ClientData>".len(),
            None => break,
        };
        let client_xml = &xml[client_start..client_end];

        if !client_xml.contains("ObjectType=\"Note\"") {
            cursor = client_end;
            continue;
        }

        let row_text = extract_first_tag_text(client_xml, "x:Row");
        let col_text = extract_first_tag_text(client_xml, "x:Column");
        if let (Some(row_text), Some(col_text)) = (row_text, col_text) {
            if let (Ok(row), Ok(col)) = (row_text.parse::<usize>(), col_text.parse::<usize>()) {
                out.push(CellRef::new(row as u32, col as u32));
            }
        }

        cursor = client_end;
    }

    Ok(out)
}

fn extract_all_tag_text(xml: &str, tag: &str) -> Vec<String> {
    let mut out = Vec::new();
    let open_pat = format!("<{}", tag);
    let close_pat = format!("</{}>", tag);

    let mut cursor = 0usize;
    while let Some(open_pos_rel) = xml[cursor..].find(&open_pat) {
        let open_pos = cursor + open_pos_rel;
        // Ensure we match the full tag name, not a prefix (e.g. `<author>` vs `<authors>`).
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

fn extract_first_tag_text(xml: &str, tag: &str) -> Option<String> {
    let open_pat = format!("<{}>", tag);
    let close_pat = format!("</{}>", tag);
    let start = xml.find(&open_pat)? + open_pat.len();
    let end = xml[start..].find(&close_pat)? + start;
    Some(xml[start..end].to_string())
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
