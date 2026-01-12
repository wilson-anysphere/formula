# VBA Digital Signatures (vbaProject.bin / vbaProjectSignature.bin)

This document captures how Excel/VBA macro signatures are stored in XLSM files, and how we extract
and **bind** those signatures against the MS-OVBA “VBA project digest”.

In this repo, `crates/formula-vba` implements best-effort:

- signature stream parsing + PKCS#7/CMS internal verification (`verify_vba_digital_signature`)
- extraction of the signed digest (`SpcIndirectDataContent` → `DigestInfo`)
- MS-OVBA §2.4.2 “Contents Hash” binding verification (signature is bound to the normalized VBA
  project content: `ContentNormalizedData` + optional `FormsNormalizedData`), exposed via
  `VbaDigitalSignature::binding`

API notes:

- `formula_vba::verify_vba_digital_signature` returns a [`VbaDigitalSignature`] with a coarse
  binding enum (`VbaSignatureBinding::{Bound, NotBound, Unknown}`).
- `formula_vba::verify_vba_digital_signature_bound` returns a [`VbaDigitalSignatureBound`] with a
  richer binding enum (`VbaProjectBindingVerification`) and best-effort debug info (algorithm OID,
  signed digest bytes, computed digest bytes).
- PKCS#7/CMS verification is *internal* signature verification only: by default we do **not**
  validate the signer certificate chain (OpenSSL `NOVERIFY`). If you need opt-in “trusted publisher”
  evaluation, use `formula_vba::verify_vba_digital_signature_with_trust` with an explicit root
  certificate set.

## Where signatures live

`xl/vbaProject.bin` is an OLE compound document. VBA signatures are stored as special OLE streams
whose names begin with the control character `0x05` (U+0005):

- `\x05DigitalSignature`
- `\x05DigitalSignatureEx`
- `\x05DigitalSignatureExt`

Notes:

- This is **not** the same as an OPC/package-level Digital Signature (XML-DSig) stored in
  `_xmlsignatures/*` parts. VBA signatures use Authenticode-like PKCS#7/CMS and can appear either as
  OLE streams inside `vbaProject.bin` (described here) or in a dedicated signature part (see below).
- These can appear at the root, e.g. `\x05DigitalSignature`.
- Some producers store the signature as a **storage** named `\x05DigitalSignature*` containing one
  or more streams, e.g. `\x05DigitalSignature/sig`. Signature discovery should therefore match on
  any *path component*, not only a root stream.
- If more than one signature stream exists, Excel prefers the newest stream:
  `DigitalSignatureExt` → `DigitalSignatureEx` → `DigitalSignature`.
  (This stream-name precedence is not normatively specified in MS-OVBA.)

## Signatures stored in external OPC parts (vbaProjectSignature.bin)

Some XLSM producers store the VBA signature **outside** `xl/vbaProject.bin`, in a separate OPC part.

- Common part name: `xl/vbaProjectSignature.bin`
- The part is typically referenced from `xl/_rels/vbaProject.bin.rels` via a relationship with type:
  `http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature`
- The relationship target may point at a different part name (resolve it relative to
  `xl/vbaProject.bin` rather than hard-coding `xl/vbaProjectSignature.bin`).

Payload variants seen in the wild:

- An OLE/CFB container containing a `\x05DigitalSignature*` stream (similar to the embedded case).
- **Raw PKCS#7/CMS DER bytes** (not an OLE compound document) stored directly in the part.

### `formula-xlsx` behavior

When inspecting/verifying signatures, `formula-xlsx`:

1. Prefers the dedicated signature part when present, resolving it via `xl/_rels/vbaProject.bin.rels`
   (with a fallback to `xl/vbaProjectSignature.bin`).
2. Attempts to verify the signature-part bytes:
   - first as an OLE container (delegating to `formula-vba`), and
   - if that fails, as a raw PKCS#7/CMS signature blob.
3. Falls back to scanning `xl/vbaProject.bin` for embedded `\x05DigitalSignature*` streams.

### Binding considerations for external signature parts

The dedicated signature part contains the **signature payload**, but MS-OVBA binding is computed
over the VBA project streams in `xl/vbaProject.bin`. As a result:

- When `xl/vbaProject.bin` is available (normal for XLSM), `formula-xlsx` uses those bytes when
  verifying/binding a signature stored in the signature part.
- If you only have the signature-part bytes (no project bytes), you can still verify the PKCS#7/CMS
  signature, but binding cannot be evaluated and should be treated as `Unknown`.

## Signature stream payload variants (what the bytes look like)

In the wild, the signature stream bytes are not always “just a PKCS#7 blob”. Common patterns:

1. **Raw PKCS#7/CMS DER**
   - The stream is a DER-encoded CMS `ContentInfo` (ASN.1 `SEQUENCE`, often starting with `0x30`).
   - `ContentInfo.contentType` is typically `signedData` (`1.2.840.113549.1.7.2`).
2. **MS-OSHARED `DigSigInfoSerialized` wrapper/prefix**
   - Some Office producers wrap or prefix the CMS bytes with a `DigSigInfoSerialized` structure
     (see MS-OSHARED).
   - `DigSigInfoSerialized` is a little-endian length-prefixed structure; parsing it lets us locate
     the embedded CMS payload deterministically instead of scanning for a DER `SEQUENCE`.
   - Common (but not universal) layout:
     - `u32le cbSignature`
     - `u32le cbSigningCertStore`
     - `u32le cchProjectName` (often a UTF-16 code unit count, but some producers treat this as a
       byte length)
     - followed by variable-size blobs (often: `projectNameUtf16le`, `certStoreBytes`,
       `signatureBytes`)
     - some producers include a `version` and/or `reserved` DWORD before the length fields.
   - The ordering of the variable blobs can vary across producers/versions. A robust parser should
     compute candidate offsets and **validate** by checking that the bytes at the computed offset
     parse as a CMS `signedData` `ContentInfo`.
   - The wrapper's `cbSignature` region can include padding; prefer using the DER length inside the
     region to find the actual CMS payload size.
   - ⚠️ The `cbSigningCertStore` blob may itself contain a PKCS#7/CMS structure (often beginning
     with `0x30`). This means naive “scan for the first PKCS#7 SignedData” logic can pick the
     **certificate store** rather than the actual signature. Prefer the `DigSigInfoSerialized`-derived
     `(offset, len)` when available, and when falling back to scanning heuristics, prefer the **last**
     plausible `SignedData` candidate in the stream.
3. **Detached `content || pkcs7`**
   - The stream contains `signed_content_bytes` followed by a detached CMS signature (`pkcs7_der`).
   - Verification must pass the prefix bytes as the detached content when verifying the CMS blob.

## Extracting the “signed digest” (Authenticode-style)

Excel’s VBA signature is Authenticode-like: the CMS `SignedData` encapsulates a structure that
includes the digest we need for MS-OVBA binding.

High-level extraction steps:

1. **Obtain the CMS `ContentInfo` bytes**
   - If the stream is raw CMS DER, use it directly.
   - If the stream is a `DigSigInfoSerialized` wrapper, unwrap it (per MS-OSHARED).
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
| Digest algorithm: MD5 | `1.2.840.113549.2.5` |
| Digest algorithm: SHA-1 | `1.3.14.3.2.26` |
| Digest algorithm: SHA-256 | `2.16.840.1.101.3.4.2.1` |

## Binding (MS-OVBA project digest verification)

CMS signature verification alone answers “is this a valid CMS signature over *some bytes*?”, but it
does not, by itself, prove that the signature is bound to the rest of the VBA project.

### MS-OVBA §2.4.2 “Contents Hash” (digest transcript versions)

MS-OVBA defines a **versioned** “Contents Hash” computation (§2.4.2) used to bind a
`\x05DigitalSignature*` stream to the rest of `vbaProject.bin`.

Terminology gotcha: there are *two* hash concepts in play:

- **CMS/PKCS#7 signature algorithm**: the algorithm used by the signer to sign the CMS `SignedData`
  (often SHA-256 + RSA today). This is what OpenSSL validates when we say the signature is
  cryptographically verified.
- **MS-OVBA “Contents Hash” digest bytes**: the bytes stored in Authenticode’s
  `SpcIndirectDataContent.messageDigest: DigestInfo.digest`, intended to be compared against a hash
  computed from the current VBA project.

The second one is what MS-OVBA §2.4.2 defines: it specifies a deterministic **byte transcript**
(`ProjectNormalizedData`) and then hashes that transcript to produce the digest bytes embedded in
`DigestInfo.digest`.

#### Signature stream variant → contents-hash version → digest algorithm

MS-OVBA has three on-disk signature stream names. Their **contents-hash version** affects both:
1) how to build `ProjectNormalizedData`, and 2) which hash algorithm produces the digest bytes.

| OLE signature stream | Contents hash version | Digest that `DigestInfo.digest` should match | Digest bytes algorithm | Digest bytes length |
|---|---:|---|---|---:|
| `\x05DigitalSignature` | v1 | **Content Hash** (§2.4.2.3): `MD5(ContentNormalizedData)` | MD5 | 16 |
| `\x05DigitalSignatureEx` | v2 | **Agile Content Hash** (§2.4.2.4): `MD5(ContentNormalizedData \|\| FormsNormalizedData)` | MD5 | 16 |
| `\x05DigitalSignatureExt` | v3 | v3 digest (SHA-256 over `ProjectNormalizedData`; see §2.4.2.5/§2.4.2.6) | SHA-256 | 32 |

Important: for v1/v2, Office can store **MD5 digest bytes** even when
`DigestInfo.digestAlgorithm.algorithm` is *not* the MD5 OID (see MS-OSHARED §4.3, and the note in
`crates/formula-vba/src/signature.rs:405-408`). This is why binding code should primarily trust the
*digest length* and the stream variant, not the `DigestInfo` algorithm OID, when choosing how to
compute the expected digest bytes.

#### §2.4.2.1 ContentNormalizedData (v1)

`ContentNormalizedData` is the v1 transcript for module source + some reference data.

Transcript construction (as defined by MS-OVBA pseudocode; see also
`formula_vba::content_normalized_data` in `crates/formula-vba/src/contents_hash.rs:25`):

1. Read the `VBA/dir` stream.
2. Decompress it using the MS-OVBA `CompressedContainer` algorithm (§2.4.1).
3. Parse the decompressed bytes as a sequence of “dir records”:
   - `id: u16le`
   - `len: u32le`
   - `data: [u8; len]`
4. Initialize `ContentNormalizedData = []`.
5. As you scan records in order:
   - For each **`REFERENCEREGISTERED`** record (`id == 0x000D`), append the record `data` bytes
     verbatim to `ContentNormalizedData`.
   - For each **`REFERENCEPROJECT`** record (`id == 0x000E`), append a *normalized* byte sequence:
     1. Parse `LibidAbsolute` and `LibidRelative` as `u32le length || bytes`.
     2. Parse `MajorVersion` as `u32le` and `MinorVersion` as `u16le`.
     3. Build `TempBuffer = LibidAbsolute || LibidRelative || MajorVersion || MinorVersion`.
     4. Append bytes from `TempBuffer` up to (but not including) the first `0x00` byte.
   - For each module (as described by the module record groups in `VBA/dir`), locate the module
     stream name (`MODULESTREAMNAME` / `MODULENAME`) and `MODULETEXTOFFSET`. The **module ordering**
     is the order of the module groups in `VBA/dir`, not alphabetical and not OLE directory order.
6. For each module stream in that stored order:
   1. Read `VBA/<ModuleStreamName>` bytes.
   2. Find the compressed source container starting at `MODULETEXTOFFSET` (or, if missing, locate the
      start of a compressed container via signature scan; see `contents_hash.rs` / `lib.rs`).
   3. Decompress the module’s `CompressedContainer` bytes.
   4. Normalize the module source bytes:
      - Split into lines where a line break is either `CR` (`0x0D`) or a lone `LF` (`0x0A`); treat
        `CRLF` as a single break (ignore the `LF`).
      - Drop any line whose first token is `Attribute` (ASCII case-insensitive match at start of
        line, followed by end-of-line or whitespace).
      - For every remaining line, append the line bytes followed by `CRLF` (`0x0D 0x0A`).

The resulting concatenated byte sequence is `ContentNormalizedData`.

#### §2.4.2.2 FormsNormalizedData

`FormsNormalizedData` is a transcript over “designer” streams (UserForms, etc).

Transcript construction (see also `formula_vba::forms_normalized_data` in
`crates/formula-vba/src/normalized_data.rs:17`):

1. Enumerate all OLE streams that live under a **root-level storage** that is:
   - not the `VBA` storage, and
   - not a `\x05DigitalSignature*` storage.
2. Recursively include streams in nested storages under those designer roots.
3. Sort streams lexicographically by full OLE path (e.g. `UserForm1/Child/X` before `UserForm1/Y`).
4. Initialize `FormsNormalizedData = []`.
5. For each stream in that sorted order:
   1. Append the stream’s raw bytes.
   2. Pad with `0x00` bytes up to a multiple of **1023 bytes** (MS-OVBA hashes designer data in
      1023-byte blocks and zero-pads the final partial block).

#### §2.4.2.3 Content Hash (v1)

Content Hash (MS-OVBA §2.4.2.3) is the **16-byte MD5 digest** of `ContentNormalizedData`:

```text
ContentHash = MD5(ContentNormalizedData)
```

This is the legacy digest most commonly associated with the `\x05DigitalSignature` stream variant.

#### §2.4.2.4 Agile Content Hash (v2 / forms-aware)

Agile Content Hash (MS-OVBA §2.4.2.4) is the **16-byte MD5 digest** of:

```text
AgileContentHash = MD5(ContentNormalizedData || FormsNormalizedData)
```

This is the digest most commonly associated with the `\x05DigitalSignatureEx` stream variant.

#### §2.4.2.5 V3ContentNormalizedData (Contents hash v3)

`V3ContentNormalizedData` is the v3 replacement for `ContentNormalizedData`, used by the
`DigitalSignatureExt` stream variant.

Transcript construction (MS-OVBA §2.4.2.5):

1. Read and decompress `VBA/dir` as in §2.4.2.1.
2. Parse it as a sequence of `(id, len, data)` records (same physical format as v1).
3. Build a module list from the module record groups in `VBA/dir` (same ordering rule as v1: the
   stored order in the dir stream).
4. Construct `V3ContentNormalizedData` by concatenating the spec-specified record bytes and module
   bytes in the order described by §2.4.2.5.

The main point for implementers: this is **not** just “`ContentNormalizedData` hashed with a stronger
algorithm”. v3 defines a distinct transcript (`V3ContentNormalizedData`) and `DigitalSignatureExt`
uses SHA-256 over `ProjectNormalizedData` (§2.4.2.6).

Repo status:

- `formula-vba` does **not** currently implement `V3ContentNormalizedData`.
- The natural implementation hook is a new function alongside `content_normalized_data` in
  `crates/formula-vba/src/contents_hash.rs` (with golden tests based on real `DigitalSignatureExt`
  fixtures).
- `crates/formula-vba/src/project_digest.rs` already has a `DigestAlg::Sha256` variant, but the
  binding logic in `crates/formula-vba/src/signature.rs` always uses `DigestAlg::Md5` today.

#### §2.4.2.6 ProjectNormalizedData (the hashed transcript)

`ProjectNormalizedData` is the spec-defined transcript that is actually hashed to produce the digest
bytes placed in `DigestInfo.digest`.

Per MS-OVBA §2.4.2.6, the transcript depends on the **contents-hash version**:

- v1 (`DigitalSignature`): `ProjectNormalizedData = ContentNormalizedData`
- v2 (`DigitalSignatureEx`): `ProjectNormalizedData = ContentNormalizedData || FormsNormalizedData`
- v3 (`DigitalSignatureExt`): `ProjectNormalizedData = V3ContentNormalizedData || FormsNormalizedData`

### `formula-vba` binding implementation (current)

To bind the signature to the VBA project contents, `formula-vba` currently:

1. Extracts the signed digest bytes from `SpcIndirectDataContent.messageDigest: DigestInfo`.
   - The `DigestInfo.digestAlgorithm.algorithm` OID is kept for debug display, but is not used to
     select the VBA project digest algorithm.
2. Recomputes the MS-OVBA §2.4.2 “Contents Hash” transcript over the VBA project by building:
   - `ContentNormalizedData` (MS-OVBA §2.4.2.1), and
   - `FormsNormalizedData` (MS-OVBA §2.4.2.2), if the project contains designer storages/streams.
3. Computes **Content Hash** (MS-OVBA §2.4.2.3) as **MD5(ContentNormalizedData)** (16 bytes), per
   MS-OSHARED §4.3.
4. Computes **Agile Content Hash** (MS-OVBA §2.4.2.4) as
    **MD5(ContentNormalizedData || FormsNormalizedData)** (16 bytes), when `FormsNormalizedData` is
   available.
5. Compares the signed digest bytes (`DigestInfo.digest`) against the computed hash bytes:
   - if it matches either Content Hash or Agile Content Hash ⇒ `Bound`
   - if it mismatches and we could compute Agile Content Hash ⇒ `NotBound`
   - if we could not compute Agile Content Hash (e.g. forms data missing/unparseable) ⇒ `Unknown`

Result interpretation (current behavior):

- If CMS verification fails ⇒ signature invalid (`VbaSignatureVerification::SignedInvalid` /
  `SignedParseError`).
- If CMS verification succeeds and binding matches ⇒ signature is verified and bound
  (`VbaSignatureBinding::Bound`).
- If CMS verification succeeds and binding mismatches ⇒ signature is present but **not bound** to the
  current VBA project bytes (`VbaSignatureBinding::NotBound`).
- If CMS verification succeeds but binding cannot be evaluated (missing/unparseable data, or we can't
  distinguish mismatch vs missing forms data) ⇒ binding is conservative/unknown
  (`VbaSignatureBinding::Unknown`).

### Implementation notes / caveats

- Per MS-OSHARED §4.3, the VBA project digest used for binding is always **MD5 (16 bytes)**, even if
  the PKCS#7/CMS signature uses SHA-1/SHA-256 and even if `DigestInfo.digestAlgorithm.algorithm`
  indicates SHA-256. The digest algorithm OID is treated as informational only.
- The primary binding path uses MS-OVBA §2.4.2 normalized data:
  - `ContentNormalizedData` is derived from the `VBA/dir` stream and the module streams referenced
    by it (including module ordering and module source normalization).
  - `FormsNormalizedData` is derived from streams inside root-level “designer” storages (for
    example UserForms), with per-stream padding to 1023-byte blocks.
- `compute_vba_project_digest` is a deterministic **fallback** digest over OLE stream names/bytes.
  It exists to keep binding checks useful for synthetic/partial fixtures when MS-OVBA normalization
  cannot be computed, but it is not expected to match Excel for all real-world files.
- The project digest computation is **best-effort** and deterministic (to support stable tests and
  predictable behavior), but may not match Excel's exact transcript for all real-world files. This
  can produce false negatives (a valid signature treated as not bound).
- `DigitalSignatureExt` / v3 contents hashes (SHA-256 over `ProjectNormalizedData`) are not supported
  today; adding v3 support requires implementing MS-OVBA §2.4.2.5 + §2.4.2.6 and hashing with
  SHA-256 instead of MD5.
- Callers should treat `VbaSignatureBinding::Unknown` as "could not verify binding", not as "bound".
- Treat `binding == Bound` as a strong signal, but validate against Excel fixtures before relying on
  it as a hard security boundary.

## Repo implementation pointers

If you need to update or extend signature handling, start with:

- `crates/formula-vba/src/signature.rs`
  - signature stream discovery (`\x05DigitalSignature*`)
  - CMS verification and the `VbaDigitalSignature::binding` decision
- `crates/formula-vba/src/offcrypto.rs`
  - `[MS-OSHARED] DigSigInfoSerialized` parsing (deterministic CMS offset/length)
- `crates/formula-vba/src/authenticode.rs`
  - `SignedData.encapContentInfo.eContent` parsing and `SpcIndirectDataContent` → `DigestInfo`
- `crates/formula-vba/src/contents_hash.rs`
  - MS-OVBA §2.4.2.1 `ContentNormalizedData` (`content_normalized_data`)
- `crates/formula-vba/src/normalized_data.rs`
  - MS-OVBA §2.4.2.2 `FormsNormalizedData` (`forms_normalized_data`)
- `crates/formula-vba/src/project_digest.rs`
  - deterministic fallback digest over OLE streams (used when MS-OVBA normalization fails)

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

- **MS-OVBA**: VBA project storage and VBA project digest computation / binding verification.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- **MS-OSHARED**: VBA digital signature storage (`DigSigInfoSerialized`) and the MD5 “VBA project hash” rule (MS-OSHARED §4.3).
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/
- **RFC 5652**: Cryptographic Message Syntax (CMS) (PKCS#7 `SignedData`).
  - https://www.rfc-editor.org/rfc/rfc5652
