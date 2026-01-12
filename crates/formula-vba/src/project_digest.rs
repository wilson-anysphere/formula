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
/// - For **VBA signatures**, Office stores a 16-byte **MD5** digest for binding even when
///   `DigestInfo.digestAlgorithm` indicates SHA-256 (MS-OSHARED §4.3). This applies to all
///   `\x05DigitalSignature*` variants, including `\x05DigitalSignatureExt`.
/// - Other algorithms are supported here for legacy/debugging purposes only; they are not
///   spec-correct for VBA signature *binding*.
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
/// signature *binding* uses MD5 per MS-OSHARED §4.3, callers may request other algorithms for
/// debugging or comparison.
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
/// Spec reference (important):
///
/// MS-OVBA §2.4.2.7 defines the **V3 Content Hash** used for VBA signature binding as:
///
/// `V3ContentHash = MD5(V3ContentNormalizedData || ProjectNormalizedData)`
///
/// And MS-OSHARED §4.3 specifies that the digest bytes embedded in VBA signatures are always a
/// 16-byte **MD5** for binding, even when `DigestInfo.digestAlgorithm.algorithm` advertises SHA-256.
///
/// This helper accepts `alg` for legacy/debugging purposes; callers verifying Office-produced VBA
/// signatures should use `DigestAlg::Md5`.
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
