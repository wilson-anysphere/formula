use sha1::Sha1;
use sha2::Sha256;
use sha2::Digest as _;

use crate::{OleFile, signature::SignatureError};

/// Digest algorithm used by MS-OVBA VBA project signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DigestAlg {
    Sha1,
    Sha256,
}

enum Hasher {
    Sha1(Sha1),
    Sha256(Sha256),
}

impl Hasher {
    fn new(alg: DigestAlg) -> Self {
        match alg {
            DigestAlg::Sha1 => Self::Sha1(Sha1::new()),
            DigestAlg::Sha256 => Self::Sha256(Sha256::new()),
        }
    }

    fn update(&mut self, bytes: &[u8]) {
        match self {
            Hasher::Sha1(h) => h.update(bytes),
            Hasher::Sha256(h) => h.update(bytes),
        }
    }

    fn finalize(self) -> Vec<u8> {
        match self {
            Hasher::Sha1(h) => h.finalize().to_vec(),
            Hasher::Sha256(h) => h.finalize().to_vec(),
        }
    }
}

/// Best-effort MS-OVBA project digest over the VBA project's OLE streams.
///
/// This is intended to match how Office binds a `\x05DigitalSignature*` stream to the rest of the
/// project. The exact MS-OVBA transcript is underspecified in our implementation, but we aim for a
/// deterministic and collision-resistant digest:
///
/// 1. Enumerate all streams in the OLE file.
/// 2. Exclude any `DigitalSignature*` stream/storage.
/// 3. Sort remaining stream paths in a deterministic, case-insensitive order.
/// 4. Hash each stream as:
///    `UTF-16LE(path) || 0x0000 || u32_le(len(bytes)) || bytes`.
pub fn compute_vba_project_digest(
    vba_project_bin: &[u8],
    alg: DigestAlg,
) -> Result<Vec<u8>, SignatureError> {
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

fn is_signature_component(component: &str) -> bool {
    let trimmed = component.trim_start_matches(|c: char| c <= '\u{001F}');
    matches!(
        trimmed,
        "DigitalSignature" | "DigitalSignatureEx" | "DigitalSignatureExt"
    )
}

