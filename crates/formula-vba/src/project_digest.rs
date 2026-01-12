use md5::Md5;
use sha1::Sha1;
use sha2::Digest as _;
use sha2::Sha256;

use crate::{
    content_normalized_data,
    contents_hash::project_normalized_data_v3,
    forms_normalized_data,
    signature::SignatureError,
    OleError,
    ParseError,
};

/// Digest algorithm used by helper digest computations in this module.
///
/// Notes:
/// - For legacy VBA signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`), Office
///   stores 16-byte **MD5** binding digest bytes even when `DigestInfo.digestAlgorithm` advertises
///   SHA-256 (MS-OSHARED §4.3). In practice, digest *length* is often a more reliable indicator than
///   the OID for these variants.
/// - For v3 (`\x05DigitalSignatureExt`), MS-OVBA defines `ContentsHashV3` as a **SHA-256** over the
///   v3 `ProjectNormalizedData` transcript; this is typically the binding digest bytes stored in the
///   signature.
/// - Other algorithms are supported here for debugging/compatibility checks; they are not expected
///   for Office-produced VBA signature binding.
///
/// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/40c8dab3-e8db-4c66-a6be-8cec06351b1e
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DigestAlg {
    /// MD5 (16 bytes).
    Md5,
    /// SHA-1 (supported for legacy/debugging; not spec-correct for VBA signature binding).
    Sha1,
    /// SHA-256 (supported for legacy/debugging; not spec-correct for VBA signature binding).
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

/// Compute a digest over a VBA project's MS-OVBA §2.4.2 (v1/v2) normalized-data transcript.
///
/// Transcript (deterministic):
///
/// `transcript = ContentNormalizedData || FormsNormalizedData`
///
/// Where:
/// - `ContentNormalizedData` is computed by [`crate::content_normalized_data`] (MS-OVBA §2.4.2.1)
/// - `FormsNormalizedData` is computed by [`crate::forms_normalized_data`] (MS-OVBA §2.4.2.2)
///
/// Note: For projects without designer/UserForm storages, `FormsNormalizedData` is typically empty,
/// so this digest is equivalent to the v1 "Content Hash" (`hash(ContentNormalizedData)`).
///
/// The transcript is then hashed using the requested `alg` (MD5/SHA-1/SHA-256). Even though Office
/// signature *binding* uses MD5 for legacy VBA signature streams per MS-OSHARED §4.3, callers may
/// request other algorithms for debugging or comparison.
///
/// This function is intentionally strict: if the transcript cannot be produced (missing or
/// unparseable required streams), it returns an error rather than falling back to hashing raw OLE
/// streams.
pub fn compute_vba_project_digest(
    vba_project_bin: &[u8],
    alg: DigestAlg,
) -> Result<Vec<u8>, SignatureError> {
    let content_normalized = content_normalized_data(vba_project_bin)
        .map_err(parse_error_to_signature_error)?;
    let forms_normalized =
        forms_normalized_data(vba_project_bin).map_err(parse_error_to_signature_error)?;
    let mut hasher = Hasher::new(alg);
    hasher.update(&content_normalized);
    hasher.update(&forms_normalized);
    Ok(hasher.finalize())
}

/// Compute a digest for v3 (`\x05DigitalSignatureExt`) signature binding.
///
/// Spec reference:
///
/// MS-OVBA §2.4.2.7 defines `ContentsHashV3` as:
///
/// `ContentsHashV3 = SHA-256(ProjectNormalizedDataV3)`
///
/// Where `ProjectNormalizedDataV3` is the deterministic transcript produced by
/// [`crate::project_normalized_data_v3`] (filtered `PROJECT` stream properties ||
/// `V3ContentNormalizedData` || `FormsNormalizedData`).
///
/// This helper accepts `alg` for debugging/compatibility checks; callers verifying Office-produced
/// VBA signatures should use `DigestAlg::Sha256`.
pub fn compute_vba_project_digest_v3(
    vba_project_bin: &[u8],
    alg: DigestAlg,
) -> Result<Vec<u8>, ParseError> {
    let normalized = project_normalized_data_v3(vba_project_bin)?;
    let mut hasher = Hasher::new(alg);
    hasher.update(&normalized);
    Ok(hasher.finalize())
}

fn parse_error_to_signature_error(err: ParseError) -> SignatureError {
    match err {
        ParseError::Ole(err) => SignatureError::Ole(err),
        err => SignatureError::Ole(OleError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            err,
        ))),
    }
}
