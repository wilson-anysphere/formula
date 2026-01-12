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
    let end_fragment_rel =
        find_subslice(html_bytes, END_FRAGMENT_MARKER.as_bytes()).unwrap_or(html_bytes.len());

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
    decode_cf_html_bytes(payload.as_bytes())
}

/// Decode a CF_HTML payload from raw bytes.
///
/// This is the preferred decode entrypoint for clipboard reads because CF_HTML offsets are byte
/// offsets. Passing through `String::from_utf8_lossy` before slicing can distort offsets in the
/// presence of invalid UTF-8.
pub(crate) fn decode_cf_html_bytes(payload: &[u8]) -> Option<String> {
    // Some producers include NUL termination.
    let mut end = payload.len();
    while end > 0 && payload[end - 1] == 0 {
        end -= 1;
    }
    let payload = &payload[..end];

    if payload.is_empty() {
        return None;
    }

    fn parse_offset(payload: &[u8], key: &[u8]) -> Option<usize> {
        let idx = find_subslice(payload, key)?;
        let mut i = idx + key.len();
        // Be permissive about optional whitespace after the colon.
        while i < payload.len() && payload[i].is_ascii_whitespace() {
            i += 1;
        }
        let mut value: usize = 0;
        let mut any = false;
        while i < payload.len() {
            let b = payload[i];
            if b.is_ascii_digit() {
                any = true;
                value = value.saturating_mul(10).saturating_add((b - b'0') as usize);
                i += 1;
            } else {
                break;
            }
        }
        if any { Some(value) } else { None }
    }

    fn extract_fragment_markers_bytes(html_doc: &[u8]) -> Option<String> {
        let start_marker = START_FRAGMENT_MARKER.as_bytes();
        let end_marker = END_FRAGMENT_MARKER.as_bytes();
        let start = find_subslice(html_doc, start_marker)? + start_marker.len();
        let end_rel = find_subslice(&html_doc[start..], end_marker)?;
        let end = start + end_rel;
        Some(String::from_utf8_lossy(&html_doc[start..end]).into_owned())
    }

    let bytes = payload;
    let header_end = bytes.iter().position(|&b| b == b'<').unwrap_or(bytes.len());
    let header = &bytes[..header_end];

    // 1) Prefer explicit fragment offsets.
    if let (Some(start), Some(end)) = (
        parse_offset(header, b"StartFragment:"),
        parse_offset(header, b"EndFragment:"),
    ) {
        if start < bytes.len() {
            let end = end.min(bytes.len());
            if end > start {
                let fragment_bytes = &bytes[start..end];

                // Some producers incorrectly include the fragment markers in the Start/EndFragment
                // span; strip them if present.
                if let Some(fragment) = extract_fragment_markers_bytes(fragment_bytes) {
                    return Some(fragment);
                }

                return Some(String::from_utf8_lossy(fragment_bytes).into_owned());
            }
        }
    }

    // 2) Fall back to StartHTML/EndHTML and extract markers if present.
    if let (Some(start), Some(end)) = (
        parse_offset(header, b"StartHTML:"),
        parse_offset(header, b"EndHTML:"),
    ) {
        if start < bytes.len() {
            let end = end.min(bytes.len());
            if end > start {
                let html_doc = &bytes[start..end];
                if let Some(fragment) = extract_fragment_markers_bytes(html_doc) {
                    return Some(fragment);
                }
                if !html_doc.is_empty() {
                    return Some(String::from_utf8_lossy(html_doc).into_owned());
                }
            }
        }
    }

    // 3) Try extracting standard markers anywhere in the payload.
    if let Some(fragment) = extract_fragment_markers_bytes(bytes) {
        return Some(fragment);
    }

    // 4) Last resort: strip the header by finding the first '<' byte.
    let start = bytes.iter().position(|&b| b == b'<')?;
    Some(String::from_utf8_lossy(&bytes[start..]).into_owned())
}

/// Best-effort extraction of the HTML fragment from a CF_HTML payload.
///
/// If parsing fails, returns an empty string (to avoid surfacing the CF_HTML header blob as HTML).
pub(crate) fn extract_cf_html_fragment_best_effort(payload: &[u8]) -> String {
    if let Some(decoded) = decode_cf_html_bytes(payload) {
        return decoded;
    }
    String::new()
}

#[cfg(test)]
mod tests {
    use super::{build_cf_html_payload, decode_cf_html, decode_cf_html_bytes, extract_cf_html_fragment_best_effort};

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
    fn build_cf_html_payload_offsets_work_with_utf8_fragments() {
        // Include multi-byte UTF-8 codepoints to ensure offsets are byte offsets (CF_HTML spec),
        // and that our computed offsets land on valid UTF-8 boundaries.
        let fragment = "<p>héllö π 😀</p>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let s = String::from_utf8(payload.clone()).expect("utf8");

        let start_fragment = parse_offset(&s, "StartFragment");
        let end_fragment = parse_offset(&s, "EndFragment");
        assert_eq!(&s[start_fragment..end_fragment], fragment);

        // Also ensure our decoder path still extracts the fragment.
        let mut nul_terminated = payload;
        nul_terminated.push(0);
        let decoded = decode_cf_html_bytes(&nul_terminated).expect("decoded");
        assert_eq!(decoded, fragment);
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
    fn decode_cf_html_bytes_extracts_fragment_from_offsets() {
        let fragment = "<b>Hello</b>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let decoded = decode_cf_html_bytes(&payload).expect("decoded");
        assert_eq!(decoded, fragment);
    }

    #[test]
    fn decode_cf_html_extracts_fragment_from_markers_when_no_header() {
        let payload = "<html><body><!--StartFragment--><p>Hi</p><!--EndFragment--></body></html>\0";
        let decoded = decode_cf_html(payload).expect("decoded");
        assert_eq!(decoded, "<p>Hi</p>");
    }

    #[test]
    fn decode_cf_html_bytes_strips_markers_when_offsets_include_them() {
        let fragment = "<b>Hello</b>";
        let html_doc = super::wrap_html_fragment(fragment);
        let html_bytes = html_doc.as_bytes();

        let start_marker_rel = super::find_subslice(html_bytes, super::START_FRAGMENT_MARKER.as_bytes())
            .expect("start fragment marker");
        let end_marker_end_rel = super::find_subslice(html_bytes, super::END_FRAGMENT_MARKER.as_bytes())
            .expect("end fragment marker")
            + super::END_FRAGMENT_MARKER.len();

        let placeholder = super::build_header(0, 0, 0, 0).expect("placeholder header");
        let start_html = placeholder.len();
        let end_html = start_html + html_bytes.len();
        // Intentionally include the markers inside the fragment offsets.
        let start_fragment = start_html + start_marker_rel;
        let end_fragment = start_html + end_marker_end_rel;
        let header =
            super::build_header(start_html, end_html, start_fragment, end_fragment).expect("header");
        debug_assert_eq!(header.len(), start_html);

        let mut payload = header.into_bytes();
        payload.extend_from_slice(html_bytes);

        let decoded = decode_cf_html_bytes(&payload).expect("decoded");
        assert_eq!(decoded, fragment);
    }

    #[test]
    fn encode_decode_round_trip_preserves_fragment() {
        let fragment = "<span data-x=\"1\">Hello</span>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let s = String::from_utf8(payload).expect("utf8");
        let decoded = decode_cf_html(&s).expect("decoded");
        assert_eq!(decoded, fragment);
    }

    #[test]
    fn extract_cf_html_bytes_roundtrip_preserves_fragment() {
        let fragment = "<table><tr><td>Hello</td></tr></table>";
        let payload = build_cf_html_payload(fragment).expect("payload");
        let extracted = extract_cf_html_fragment_best_effort(&payload);
        assert_eq!(extracted, fragment);
    }

    #[test]
    fn extract_cf_html_fragment_best_effort_does_not_return_header_blob() {
        // Malformed CF_HTML payload: contains only header fields and no HTML content.
        let payload = b"Version:0.9\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\n";
        let extracted = extract_cf_html_fragment_best_effort(payload);
        assert!(extracted.is_empty());
    }

    #[test]
    fn decode_cf_html_bytes_ignores_offset_keys_inside_html() {
        // This looks like a CF_HTML header field, but it's actually inside the HTML body. We should
        // not treat it as an offset directive.
        let html = "<html><body>StartFragment:0000000010</body></html>";
        let decoded = decode_cf_html_bytes(html.as_bytes()).expect("decoded");
        assert_eq!(decoded, html);
    }

    #[test]
    fn decode_cf_html_bytes_clamps_out_of_range_end_html() {
        // Some producers include trailing NUL padding in their EndHTML calculations.
        // Ensure we clamp offsets so we can still extract the HTML.
        let html = "<html><body><p>Hi</p></body></html>";

        // Build a minimal header with fixed-width offsets.
        let header_placeholder =
            format!("Version:0.9\r\nStartHTML:{:010}\r\nEndHTML:{:010}\r\n", 0, 0);
        let start_html = header_placeholder.len();
        let extra_nuls = 2usize;
        let end_html = start_html + html.as_bytes().len() + extra_nuls;
        let header = format!(
            "Version:0.9\r\nStartHTML:{start_html:010}\r\nEndHTML:{end_html:010}\r\n"
        );

        let mut payload = header.into_bytes();
        payload.extend_from_slice(html.as_bytes());
        payload.extend_from_slice(&[0u8; 2]);

        let decoded = decode_cf_html_bytes(&payload).expect("decoded");
        assert_eq!(decoded, html);
    }
}
