use url::Url;

use crate::oauth_redirect_ipc::{
    MAX_PENDING_BYTES as MAX_OAUTH_REDIRECT_PENDING_BYTES,
    MAX_PENDING_URLS as MAX_OAUTH_REDIRECT_PENDING_URLS,
};
use crate::open_file_ipc::{
    MAX_PENDING_BYTES as MAX_OPEN_FILE_PENDING_BYTES,
    MAX_PENDING_PATHS as MAX_OPEN_FILE_PENDING_URLS,
};

#[derive(Debug, Default, PartialEq, Eq)]
pub struct OpenedUrlClassification {
    /// Custom-scheme OAuth PKCE redirects (e.g. `formula://oauth/callback?...`).
    pub oauth_redirects: Vec<String>,
    /// Everything else from the OS "opened URLs" event, forwarded through the open-file pipeline.
    ///
    /// Note that the open-file extractor further filters these down to supported spreadsheet file
    /// extensions / `file://...` URLs.
    pub file_open_candidates: Vec<String>,
}

/// Split OS-delivered URLs into OAuth redirect deep links vs open-file candidates.
///
/// macOS can deliver custom-scheme URL opens to an already-running Tauri instance via
/// `tauri::RunEvent::Opened { urls, .. }`. We need to ensure `formula://...` deep links are
/// forwarded to the frontend (for OAuth), while still preserving Finder-style file open behavior
/// for `file://...` URLs.
pub fn classify_opened_urls(urls: &[Url]) -> OpenedUrlClassification {
    let schemes = crate::deep_link_schemes::configured_schemes();

    // Treat OS-delivered URL-open events as untrusted input. Bound allocations so a malicious
    // sender cannot OOM the host by delivering a huge list of opened URLs while the app is
    // already running.
    //
    // When the cap is exceeded, we drop the **oldest** entries and keep the most recent ones so
    // the latest user action wins. Caps are aligned with the pending IPC queues for these
    // pipelines.
    let mut oauth_rev = Vec::with_capacity(MAX_OAUTH_REDIRECT_PENDING_URLS.min(urls.len()));
    let mut oauth_bytes = 0usize;

    let mut file_rev = Vec::with_capacity(MAX_OPEN_FILE_PENDING_URLS.min(urls.len()));
    let mut file_bytes = 0usize;

    // Walk backwards so we keep the most recent URLs, then reverse at the end to preserve the
    // original order among kept entries.
    for url in urls.iter().rev() {
        let raw = url.as_str();
        let len = raw.len();
        let is_deep_link = schemes.iter().any(|scheme| scheme.as_str() == url.scheme());
        if is_deep_link {
            if oauth_rev.len() >= MAX_OAUTH_REDIRECT_PENDING_URLS {
                continue;
            }
            if len > MAX_OAUTH_REDIRECT_PENDING_BYTES {
                continue;
            }
            if oauth_bytes.saturating_add(len) > MAX_OAUTH_REDIRECT_PENDING_BYTES {
                continue;
            }
            oauth_bytes += len;
            oauth_rev.push(raw.to_string());
        } else {
            if file_rev.len() >= MAX_OPEN_FILE_PENDING_URLS {
                continue;
            }
            if len > MAX_OPEN_FILE_PENDING_BYTES {
                continue;
            }
            if file_bytes.saturating_add(len) > MAX_OPEN_FILE_PENDING_BYTES {
                continue;
            }
            file_bytes += len;
            file_rev.push(raw.to_string());
        }
    }

    oauth_rev.reverse();
    file_rev.reverse();

    OpenedUrlClassification {
        oauth_redirects: oauth_rev,
        file_open_candidates: file_rev,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn classifies_formula_scheme_urls_as_oauth_redirects() {
        let oauth = Url::parse("formula://oauth/callback?code=123&state=xyz").unwrap();
        let other = Url::parse("https://example.com").unwrap();

        let classified = classify_opened_urls(&[oauth.clone(), other.clone()]);

        assert_eq!(classified.oauth_redirects, vec![oauth.to_string()]);
        assert_eq!(classified.file_open_candidates, vec![other.to_string()]);
    }

    #[test]
    fn keeps_file_urls_in_file_open_candidates() {
        let dir = tempdir().unwrap();
        let file_path = dir.path().join("book.xlsx");
        let file_url = Url::from_file_path(&file_path).unwrap();

        let classified = classify_opened_urls(&[file_url.clone()]);
        assert!(classified.oauth_redirects.is_empty());
        assert_eq!(classified.file_open_candidates, vec![file_url.to_string()]);
    }

    #[test]
    fn caps_opened_urls_by_count_dropping_oldest() {
        let oauth_urls: Vec<Url> = (0..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3))
            .map(|idx| Url::parse(&format!("formula://oauth/u{idx}")).unwrap())
            .collect();

        let classified = classify_opened_urls(&oauth_urls);
        assert_eq!(classified.oauth_redirects.len(), MAX_OAUTH_REDIRECT_PENDING_URLS);

        let expected: Vec<String> = (3..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3))
            .map(|idx| format!("formula://oauth/u{idx}"))
            .collect();
        assert_eq!(classified.oauth_redirects, expected);
        assert!(classified.file_open_candidates.is_empty());
    }

    #[test]
    fn caps_opened_urls_by_total_bytes_dropping_oldest_deterministically() {
        let oauth_entry_len = 4096;
        let oauth_prefix_len = "formula://oauth/000-".len();
        let oauth_payload = "x".repeat(oauth_entry_len - oauth_prefix_len);
        let oauth_urls: Vec<Url> = (0..MAX_OAUTH_REDIRECT_PENDING_URLS)
            .map(|i| Url::parse(&format!("formula://oauth/{i:03}-{oauth_payload}")).unwrap())
            .collect();
        assert_eq!(oauth_urls[0].as_str().len(), oauth_entry_len);

        let file_entry_len = 8192;
        let file_prefix_len = "https://example.com/000-".len();
        let file_payload = "x".repeat(file_entry_len - file_prefix_len);
        let file_urls: Vec<Url> = (0..MAX_OPEN_FILE_PENDING_URLS)
            .map(|i| Url::parse(&format!("https://example.com/{i:03}-{file_payload}")).unwrap())
            .collect();
        assert_eq!(file_urls[0].as_str().len(), file_entry_len);

        let mut mixed = Vec::new();
        mixed.extend_from_slice(&oauth_urls);
        mixed.extend_from_slice(&file_urls);

        let classified = classify_opened_urls(&mixed);

        let expected_oauth_len = MAX_OAUTH_REDIRECT_PENDING_BYTES / oauth_entry_len;
        assert_eq!(classified.oauth_redirects.len(), expected_oauth_len);
        assert_eq!(
            classified.oauth_redirects,
            oauth_urls[MAX_OAUTH_REDIRECT_PENDING_URLS - expected_oauth_len..]
                .iter()
                .map(|url| url.to_string())
                .collect::<Vec<_>>()
        );

        let expected_file_len = MAX_OPEN_FILE_PENDING_BYTES / file_entry_len;
        assert_eq!(classified.file_open_candidates.len(), expected_file_len);
        assert_eq!(
            classified.file_open_candidates,
            file_urls[MAX_OPEN_FILE_PENDING_URLS - expected_file_len..]
                .iter()
                .map(|url| url.to_string())
                .collect::<Vec<_>>()
        );
    }
}
