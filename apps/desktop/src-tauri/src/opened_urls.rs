use url::Url;

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
    let mut out = OpenedUrlClassification::default();
    for url in urls {
        if url.scheme() == "formula" {
            out.oauth_redirects.push(url.to_string());
        } else {
            out.file_open_candidates.push(url.to_string());
        }
    }
    out
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
}

