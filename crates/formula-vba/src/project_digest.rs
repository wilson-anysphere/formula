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

/// Digest algorithm used by [`compute_vba_project_digest`].
///
/// Note: actual Office VBA signature *binding* uses an MD5 digest (16 bytes) per MS-OSHARED §4.3.
/// The hash algorithm OID stored in Authenticode `DigestInfo` is often SHA-256 in the wild, but is
/// not used to select the VBA project digest algorithm.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DigestAlg {
    /// MD5 (16 bytes).
    ///
    /// Per MS-OSHARED §4.3 this is the VBA project "source hash" algorithm even when the PKCS#7/CMS
    /// signature uses SHA-1/SHA-256 (and even when `DigestInfo.digestAlgorithm` indicates SHA-256).
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
/// Note: For projects without designer/UserForm storages, `FormsNormalizedData` is typically empty,
/// so this digest is equivalent to the v1 "Content Hash" (`hash(ContentNormalizedData)`).
///
/// The transcript is then hashed using the requested `alg` (MD5/SHA-1/SHA-256). Even though Office
/// signature *binding* uses MD5, callers may request other algorithms for debugging or comparison.
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

fn parse_error_to_signature_error(err: ParseError) -> SignatureError {
    match err {
        ParseError::Ole(err) => SignatureError::Ole(err),
        err => SignatureError::Ole(OleError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            err,
        ))),
    }
}
