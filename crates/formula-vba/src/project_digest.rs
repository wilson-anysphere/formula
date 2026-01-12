use md5::Md5;
use sha1::Sha1;
use sha2::Digest as _;
use sha2::Sha256;

use crate::{
    content_normalized_data,
    contents_hash::project_normalized_data_v3,
    forms_normalized_data,
    signature::SignatureError,
    OleFile,
    ParseError,
};

/// Digest algorithm used by [`compute_vba_project_digest`].
///
/// Note: actual Office VBA signature *binding* uses an MD5 digest (16 bytes) per MS-OSHARED §4.3.
/// The hash algorithm OID stored in Authenticode `DigestInfo` is often SHA-256 in the wild, but is
/// not used to select the VBA project digest algorithm.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DigestAlg {
    /// MD5 (16 bytes). Per MS-OSHARED §4.3 this is the VBA project "source hash" algorithm even
    /// when the PKCS#7/CMS signature uses SHA-1/SHA-256.
    Md5,
    Sha1,
    Sha256,
}

enum Hasher {
    Md5(Md5),
    Sha1(Sha1),
    Sha256(Sha256),
}

impl Hasher {
    fn new(alg: DigestAlg) -> Self {
        match alg {
            DigestAlg::Md5 => Self::Md5(Md5::new()),
            DigestAlg::Sha1 => Self::Sha1(Sha1::new()),
            DigestAlg::Sha256 => Self::Sha256(Sha256::new()),
        }
    }

    fn update(&mut self, bytes: &[u8]) {
        match self {
            Hasher::Md5(h) => h.update(bytes),
            Hasher::Sha1(h) => h.update(bytes),
            Hasher::Sha256(h) => h.update(bytes),
        }
    }

    fn finalize(self) -> Vec<u8> {
        match self {
            Hasher::Md5(h) => h.finalize().to_vec(),
            Hasher::Sha1(h) => h.finalize().to_vec(),
            Hasher::Sha256(h) => h.finalize().to_vec(),
        }
    }
}

/// Compute a digest over a VBA project's MS-OVBA §2.4.2 digest transcript.
///
/// Transcript (deterministic):
///
/// `transcript = ContentNormalizedData || FormsNormalizedData`
///
/// Where:
/// - `ContentNormalizedData` is computed by [`crate::content_normalized_data`] (MS-OVBA §2.4.2.1)
/// - `FormsNormalizedData` is computed by [`crate::forms_normalized_data`] (MS-OVBA §2.4.2.2)
///
/// The transcript is then hashed using the requested `alg` (MD5/SHA-1/SHA-256). Even though Office
/// signature *binding* uses MD5, callers may request other algorithms for debugging or comparison.
pub fn compute_vba_project_digest(
    vba_project_bin: &[u8],
    alg: DigestAlg,
) -> Result<Vec<u8>, SignatureError> {
    // Prefer the MS-OVBA normalization transcript when possible, since this is what VBA signatures
    // are intended to bind to.
    //
    // This can fail for malformed/minimal OLE payloads (e.g. tests that intentionally omit a valid
    // `VBA/dir` compressed container). In that case, we fall back to hashing the raw OLE streams
    // directly as a best-effort digest.
    if let Ok(content_normalized) = content_normalized_data(vba_project_bin) {
        let mut hasher = Hasher::new(alg);
        hasher.update(&content_normalized);
        // Agile Content Hash extends the legacy transcript with FormsNormalizedData. For projects
        // without designer storages, this is typically empty, so hashing it is equivalent to the
        // legacy digest.
        if let Ok(forms_normalized) = forms_normalized_data(vba_project_bin) {
            hasher.update(&forms_normalized);
        }
        return Ok(hasher.finalize());
    }

    let mut ole = OleFile::open(vba_project_bin)?;
    let streams = ole.list_streams()?;

    let mut paths = streams
        .into_iter()
        .filter(|path| !path.split('/').any(is_signature_component))
        .collect::<Vec<_>>();

    // Canonical, deterministic ordering:
    // - case-insensitive compare on the full path
    // - tie-break with the original path bytes
    paths.sort_by(|a, b| a.to_lowercase().cmp(&b.to_lowercase()).then(a.cmp(b)));

    let mut hasher = Hasher::new(alg);

    for path in paths {
        let bytes = ole.read_stream_opt(&path)?.unwrap_or_default();

        // Stream name (UTF-16LE) + NUL terminator.
        for unit in path.encode_utf16() {
            hasher.update(&unit.to_le_bytes());
        }
        hasher.update(&0u16.to_le_bytes());

        // Stream length (little-endian) followed by bytes.
        let len_u32 = u32::try_from(bytes.len()).unwrap_or(u32::MAX);
        hasher.update(&len_u32.to_le_bytes());
        hasher.update(&bytes);
    }

    Ok(hasher.finalize())
}

/// Compute the MS-OVBA "Contents Hash" v3 digest of a `vbaProject.bin`.
///
/// This is the project digest used by the MS-OVBA §2.4.2 v3 transcript (`ProjectNormalizedData`
/// constructed from `V3ContentNormalizedData` + `FormsNormalizedData`), commonly associated with
/// the `\x05DigitalSignatureExt` signature stream.
pub fn compute_vba_project_digest_v3(
    vba_project_bin: &[u8],
    alg: DigestAlg,
) -> Result<Vec<u8>, ParseError> {
    let normalized = project_normalized_data_v3(vba_project_bin)?;
    let mut hasher = Hasher::new(alg);
    hasher.update(&normalized);
    Ok(hasher.finalize())
}

fn is_signature_component(component: &str) -> bool {
    let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
    matches!(
        trimmed,
        "DigitalSignature" | "DigitalSignatureEx" | "DigitalSignatureExt"
    )
}
