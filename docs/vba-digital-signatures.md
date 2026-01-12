# VBA Digital Signatures (vbaProject.bin)

This document captures how Excel/VBA macro signatures are stored in `xl/vbaProject.bin`, and how we
extract and **bind** those signatures against the MS-OVBA “VBA project digest”.

In this repo, `crates/formula-vba` implements best-effort:

- signature stream parsing + PKCS#7/CMS internal verification (`verify_vba_digital_signature`)
- extraction of the signed digest (`SpcIndirectDataContent` → `DigestInfo`)
- MS-OVBA-style project digest binding verification (signature is bound to the VBA project OLE
  streams), exposed via `VbaDigitalSignature::binding`

## Where signatures live

`xl/vbaProject.bin` is an OLE compound document. VBA signatures are stored as special OLE streams
whose names begin with the control character `0x05` (U+0005):

- `\x05DigitalSignature`
- `\x05DigitalSignatureEx`
- `\x05DigitalSignatureExt`

Notes:

- This is **not** the same as an OPC/package-level Digital Signature (XML-DSig) stored in
  `_xmlsignatures/*` parts. The VBA signature lives inside the embedded OLE `vbaProject.bin`.
- These can appear at the root, e.g. `\x05DigitalSignature`.
- Some producers store the signature as a **storage** named `\x05DigitalSignature*` containing one
  or more streams, e.g. `\x05DigitalSignature/sig`. Signature discovery should therefore match on
  any *path component*, not only a root stream.
- If more than one signature stream exists, Excel/MS-OVBA prefers the newest stream:
  `DigitalSignatureExt` → `DigitalSignatureEx` → `DigitalSignature`.

## Signature stream payload variants (what the bytes look like)

In the wild, the signature stream bytes are not always “just a PKCS#7 blob”. Common patterns:

1. **Raw PKCS#7/CMS DER**
   - The stream is a DER-encoded CMS `ContentInfo` (ASN.1 `SEQUENCE`, often starting with `0x30`).
   - `ContentInfo.contentType` is typically `signedData` (`1.2.840.113549.1.7.2`).
2. **MS-OFFCRYPTO `DigSigInfoSerialized` wrapper/prefix**
   - Some Office producers wrap or prefix the CMS bytes with a `DigSigInfoSerialized` structure
     (see MS-OFFCRYPTO).
   - `DigSigInfoSerialized` is a little-endian length-prefixed structure; parsing it lets us locate
     the embedded CMS payload deterministically instead of scanning for a DER `SEQUENCE`.
   - Common (but not universal) layout:
     - `u32le cbSignature`
     - `u32le cbSigningCertStore`
     - `u32le cchProjectName` (often a UTF-16 code unit count, but some producers treat this as a
       byte length)
     - followed by variable-size blobs (often: `projectNameUtf16le`, `certStoreBytes`, `signatureBytes`)
     - some producers include a `version` and/or `reserved` DWORD before the length fields.
   - The ordering of the variable blobs can vary across producers/versions. A robust parser should
     compute candidate offsets and **validate** by checking that the bytes at the computed offset
     parse as a CMS `signedData` `ContentInfo`.
   - The wrapper's `cbSignature` region can include padding; prefer using the DER length inside the
     region to find the actual CMS payload size.
3. **Detached `content || pkcs7`**
   - The stream contains `signed_content_bytes` followed by a detached CMS signature (`pkcs7_der`).
   - Verification must pass the prefix bytes as the detached content when verifying the CMS blob.

## Extracting the “signed digest” (Authenticode-style)

Excel’s VBA signature is Authenticode-like: the CMS `SignedData` encapsulates a structure that
includes the digest we need for MS-OVBA binding.

High-level extraction steps:

1. **Obtain the CMS `ContentInfo` bytes**
   - If the stream is raw CMS DER, use it directly.
   - If the stream is a `DigSigInfoSerialized` wrapper, unwrap it (per MS-OFFCRYPTO).
   - If the stream is `content || pkcs7`, split it (find the CMS DER start); the `content` prefix is
     the detached signed content.
2. **Parse CMS and locate `SignedData.encapContentInfo`**
   - `ContentInfo.contentType` should be `signedData` (`1.2.840.113549.1.7.2`).
   - `SignedData.encapContentInfo.eContentType` is typically `SpcIndirectDataContent`
     (`1.3.6.1.4.1.311.2.1.4`).
3. **Extract the `eContent` bytes**
   - If the CMS is *not* detached, take `encapContentInfo.eContent` (an OCTET STRING) and parse the
     contained bytes as DER.
   - If the CMS *is* detached and the stream is `content || pkcs7`, the `content` prefix plays the
     same role as `eContent`.
4. **Parse `SpcIndirectDataContent` and read its `messageDigest: DigestInfo`**
   - `DigestInfo` includes:
     - a hash algorithm identifier (OID)
     - the digest bytes

That `(hash_oid, digest_bytes)` pair is the “signed digest” we use to bind a VBA signature to a
specific VBA project.

### Relevant ASN.1 shapes (high level)

This is a simplified sketch of the structures we care about (names from RFC 5652 / Authenticode):

```text
ContentInfo ::= SEQUENCE {
  contentType OBJECT IDENTIFIER,                -- e.g. signedData (1.2.840.113549.1.7.2)
  content     [0] EXPLICIT ANY OPTIONAL
}

SignedData ::= SEQUENCE {
  ...,
  encapContentInfo EncapsulatedContentInfo,
  ...
}

EncapsulatedContentInfo ::= SEQUENCE {
  eContentType OBJECT IDENTIFIER,               -- e.g. SpcIndirectDataContent (1.3.6.1.4.1.311.2.1.4)
  eContent     [0] EXPLICIT OCTET STRING OPTIONAL
}

SpcIndirectDataContent ::= SEQUENCE {
  data          SpcAttributeTypeAndOptionalValue,  -- ignored for VBA binding
  messageDigest DigestInfo
}

DigestInfo ::= SEQUENCE {
  digestAlgorithm AlgorithmIdentifier,
  digest          OCTET STRING
}
```

For VBA binding, we ignore `SpcIndirectDataContent.data` and use only `messageDigest`.

### OIDs (quick reference)

| Meaning | OID |
|---|---|
| CMS `signedData` | `1.2.840.113549.1.7.2` |
| Authenticode `SpcIndirectDataContent` | `1.3.6.1.4.1.311.2.1.4` |
| Digest algorithm: SHA-1 | `1.3.14.3.2.26` |
| Digest algorithm: SHA-256 | `2.16.840.1.101.3.4.2.1` |

## Binding (MS-OVBA project digest verification)

CMS signature verification alone answers “is this a valid CMS signature over *some bytes*?”, but it
does not, by itself, prove that the signature is bound to the rest of the VBA project.

To bind the signature to the VBA project contents, `formula-vba`:

1. Extracts `DigestInfo` from `SpcIndirectDataContent` (hash algorithm OID + digest bytes).
2. Computes a deterministic **project digest** over the VBA project's OLE streams, excluding any
   `\x05DigitalSignature*` streams/storages.
3. Uses the hash algorithm indicated by the extracted `DigestInfo` when computing the digest.
4. Compares the computed digest bytes to the `DigestInfo` digest bytes.

Result interpretation (current behavior):

- If CMS verification fails ⇒ signature invalid.
- If CMS verification succeeds but digest comparison fails ⇒ signature present but **not bound** to
  the current VBA project bytes (reported as `VbaSignatureBinding::NotBound`).
- If both succeed ⇒ signature is verified and bound (reported as `VbaSignatureBinding::Bound`).

### Implementation notes / caveats

- The binding implementation currently supports SHA-1 and SHA-256 digests (based on the `DigestInfo`
  algorithm OID). If an unknown digest OID is encountered, binding is reported as
  `VbaSignatureBinding::Unknown`.
- Current project digest transcript (implementation detail; `compute_vba_project_digest`):
  1. Enumerate all OLE streams.
  2. Exclude any `DigitalSignature*` stream/storage.
  3. Sort remaining stream paths case-insensitively.
  4. Hash each stream as: `UTF-16LE(path) || 0x0000 || u32_le(len(bytes)) || bytes`.
- The project digest computation is currently **best-effort** and deterministic (to support stable
  tests and predictable behavior), but may not match Excel's exact MS-OVBA transcript for all
  real-world files (e.g. if Excel hashes decompressed module source, only parts of module streams,
  etc.).
- Callers should treat `VbaSignatureBinding::Unknown` as "could not verify binding", not as "bound".
- Treat `binding == Bound` as a strong signal, but validate against Excel fixtures before relying on
  it as a hard security boundary.

## Repo implementation pointers

If you need to update or extend signature handling, start with:

- `crates/formula-vba/src/signature.rs`
  - signature stream discovery (`\x05DigitalSignature*`)
  - CMS verification and the `VbaDigitalSignature::binding` decision
- `crates/formula-vba/src/offcrypto.rs`
  - `[MS-OFFCRYPTO] DigSigInfoSerialized` parsing (deterministic CMS offset/length)
- `crates/formula-vba/src/authenticode.rs`
  - `SignedData.encapContentInfo.eContent` parsing and `SpcIndirectDataContent` → `DigestInfo`
- `crates/formula-vba/src/project_digest.rs`
  - best-effort MS-OVBA-style project digest transcript over OLE streams

## Tests / examples in this repo

The `formula-vba` crate includes unit tests that act as runnable examples of the signature formats
and binding behavior:

- `crates/formula-vba/tests/signature_parse.rs`
  - verifies signature stream discovery including the nested-storage edge case (`\x05DigitalSignature/sig`).
- `crates/formula-vba/tests/signature.rs`
  - verifies PKCS#7 verification behavior, including prefix scanning and detached `content || pkcs7`.
- `crates/formula-vba/tests/signed_digest.rs`
  - exercises `SpcIndirectDataContent` / `DigestInfo` extraction.
- `crates/formula-vba/tests/signature_binding.rs`
  - constructs an Authenticode-like `SpcIndirectDataContent` payload with an embedded digest, signs it,
    and checks that tampering with non-signature streams flips `VbaSignatureBinding` to `NotBound`.

## Specs / references

- **MS-OVBA**: VBA project storage, signature streams, and VBA project digest computation.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- **MS-OFFCRYPTO**: Office cryptography structures, including `DigSigInfoSerialized`.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **RFC 5652**: Cryptographic Message Syntax (CMS) (PKCS#7 `SignedData`).
  - https://www.rfc-editor.org/rfc/rfc5652
