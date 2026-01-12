/// Helpers for the Windows CF_HTML clipboard format.
///
/// Windows applications commonly expose HTML on the clipboard via the registered clipboard
/// format name `HTML Format`. The payload is an ASCII header containing byte offsets, followed by
/// the HTML document itself.
///
/// The frontend can already *consume* CF_HTML payloads via `apps/desktop/src/clipboard/html.js`,
/// but we still need to *produce* CF_HTML when writing rich clipboard content on Windows.
///
/// See: https://learn.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
use std::borrow::Cow;

const START_FRAGMENT_MARKER: &str = "<!--StartFragment-->";
const END_FRAGMENT_MARKER: &str = "<!--EndFragment-->";

fn wrap_html_fragment(fragment: &str) -> String {
    format!(
        "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body>{START_FRAGMENT_MARKER}{fragment}{END_FRAGMENT_MARKER}</body></html>"
    )
}

fn ensure_fragment_markers(html: &str) -> Cow<'_, str> {
    if html.contains(START_FRAGMENT_MARKER) && html.contains(END_FRAGMENT_MARKER) {
        Cow::Borrowed(html)
    } else {
        Cow::Owned(wrap_html_fragment(html))
    }
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn build_header(
    start_html: usize,
    end_html: usize,
    start_fragment: usize,
    end_fragment: usize,
) -> Result<String, String> {
    const MAX_OFFSET: usize = 9_999_999_999;
    for (name, value) in [
        ("StartHTML", start_html),
        ("EndHTML", end_html),
        ("StartFragment", start_fragment),
        ("EndFragment", end_fragment),
    ] {
        if value > MAX_OFFSET {
            return Err(format!(
                "CF_HTML offset {name}={value} exceeds the 10-digit limit ({MAX_OFFSET})"
            ));
        }
    }

    // 10-digit offsets are the conventional CF_HTML encoding.
    Ok(format!(
        "Version:0.9\r\n\
StartHTML:{start_html:010}\r\n\
EndHTML:{end_html:010}\r\n\
StartFragment:{start_fragment:010}\r\n\
EndFragment:{end_fragment:010}\r\n\
StartSelection:{start_fragment:010}\r\n\
EndSelection:{end_fragment:010}\r\n"
    ))
}

/// Build a CF_HTML payload (bytes) from a provided HTML string.
///
/// The output does **not** include a trailing null terminator. Callers that write the payload into
/// Win32 global memory should add a `\0` byte as many clipboard producers do.
pub(crate) fn build_cf_html_payload(html: &str) -> Result<Vec<u8>, String> {
    let html_doc = ensure_fragment_markers(html);
    let html_bytes = html_doc.as_bytes();

    let start_fragment_rel = find_subslice(html_bytes, START_FRAGMENT_MARKER.as_bytes())
        .map(|idx| idx + START_FRAGMENT_MARKER.len())
        .unwrap_or(0);
    let end_fragment_rel = find_subslice(html_bytes, END_FRAGMENT_MARKER.as_bytes())
        .unwrap_or(html_bytes.len());

    // Placeholder header to compute the byte length of the header itself. The header uses fixed
    // width (10-digit) offsets, so the header length is stable as long as we don't exceed that
    // width (enforced below).
    let placeholder = build_header(0, 0, 0, 0)?;
    let start_html = placeholder.len();
    let end_html = start_html + html_bytes.len();
    let start_fragment = start_html + start_fragment_rel;
    let end_fragment = start_html + end_fragment_rel;

    let header = build_header(start_html, end_html, start_fragment, end_fragment)?;
    debug_assert_eq!(
        header.len(),
        start_html,
        "CF_HTML header length must remain stable"
    );

    let mut out = header.into_bytes();
    out.extend_from_slice(html_bytes);
    Ok(out)
}

/// Decode a CF_HTML payload (as a string) into the contained HTML fragment or document.
///
/// Prefer the `StartFragment`/`EndFragment` byte ranges when present. If fragment offsets are
/// missing, attempt to extract the `<!--StartFragment-->` region from the HTML document. Returns
/// `None` when the payload is empty or does not contain usable HTML.
pub(crate) fn decode_cf_html(payload: &str) -> Option<String> {
    let payload = payload.trim_end_matches('\0');
    if payload.is_empty() {
        return None;
    }

    fn parse_offset(payload: &str, key: &str) -> Option<usize> {
        let idx = payload.find(key)?;
        let after = &payload[idx + key.len()..];
        let digits: String = after.chars().take_while(|c| c.is_ascii_digit()).collect();
        if digits.is_empty() {
            return None;
        }
        digits.parse::<usize>().ok()
    }

    let bytes = payload.as_bytes();

    // 1) Prefer explicit fragment offsets.
    if let (Some(start), Some(end)) = (
        parse_offset(payload, "StartFragment:"),
        parse_offset(payload, "EndFragment:"),
    ) {
        if start < end && end <= bytes.len() {
            return Some(String::from_utf8_lossy(&bytes[start..end]).into_owned());
        }
    }

    // 2) Fall back to StartHTML/EndHTML and extract markers if present.
    if let (Some(start), Some(end)) = (
        parse_offset(payload, "StartHTML:"),
        parse_offset(payload, "EndHTML:"),
    ) {
        if start < end && end <= bytes.len() {
            let html_doc = String::from_utf8_lossy(&bytes[start..end]).into_owned();
            if let Some(fragment) = extract_fragment_markers(&html_doc) {
                return Some(fragment);
            }
            if !html_doc.is_empty() {
                return Some(html_doc);
            }
        }
    }

    // 3) Try extracting standard markers anywhere in the payload.
    if let Some(fragment) = extract_fragment_markers(payload) {
        return Some(fragment);
    }

    // 4) Last resort: strip the header by finding the first '<' character.
    payload.find('<').map(|pos| payload[pos..].to_string())
}

fn extract_fragment_markers(html_doc: &str) -> Option<String> {
    let start = html_doc.find(START_FRAGMENT_MARKER)? + START_FRAGMENT_MARKER.len();
    let end_rel = html_doc[start..].find(END_FRAGMENT_MARKER)?;
    let end = start + end_rel;
    Some(html_doc[start..end].to_string())
}

#[cfg(test)]
mod tests {
    use super::{build_cf_html_payload, decode_cf_html};

    fn parse_offset(s: &str, name: &str) -> usize {
        for line in s.split("\r\n") {
            let Some(rest) = line.strip_prefix(&format!("{name}:")) else {
                continue;
            };
            return rest.trim().parse().expect("offset parse");
        }
        panic!("missing {name} offset");
    }

    #[test]
    fn build_cf_html_payload_offsets_point_to_expected_ranges() {
        let fragment = "<table><tr><td>A</td></tr></table>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let s = String::from_utf8(payload).expect("utf8");

        let start_html = parse_offset(&s, "StartHTML");
        let end_html = parse_offset(&s, "EndHTML");
        let start_fragment = parse_offset(&s, "StartFragment");
        let end_fragment = parse_offset(&s, "EndFragment");

        assert!(start_html < end_html);
        assert!(end_html <= s.len());
        assert!(start_fragment < end_fragment);
        assert!(end_fragment <= s.len());

        let html = &s[start_html..end_html];
        assert!(
            html.contains(fragment),
            "HTML range should contain original fragment"
        );

        let frag = &s[start_fragment..end_fragment];
        assert_eq!(frag, fragment);
    }

    #[test]
    fn decode_cf_html_extracts_fragment_from_offsets() {
        let fragment = "<b>Hello</b>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let s = String::from_utf8(payload).expect("utf8");
        let decoded = decode_cf_html(&s).expect("decoded");
        assert_eq!(decoded, fragment);
    }

    #[test]
    fn decode_cf_html_extracts_fragment_from_markers_when_no_header() {
        let payload =
            "<html><body><!--StartFragment--><p>Hi</p><!--EndFragment--></body></html>\0";
        let decoded = decode_cf_html(payload).expect("decoded");
        assert_eq!(decoded, "<p>Hi</p>");
    }

    #[test]
    fn encode_decode_round_trip_preserves_fragment() {
        let fragment = "<span data-x=\"1\">Hello</span>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let s = String::from_utf8(payload).expect("utf8");
        let decoded = decode_cf_html(&s).expect("decoded");
        assert_eq!(decoded, fragment);
    }
}
