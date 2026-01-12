# VBA Digital Signatures (vbaProject.bin / vbaProjectSignature.bin)

This document captures how Excel/VBA macro signatures are stored in XLSM files, and how we extract
and **bind** those signatures against the MS-OVBA “VBA project digest”.

In this repo, `crates/formula-vba` implements best-effort:

- signature stream parsing + PKCS#7/CMS internal verification (`verify_vba_digital_signature`)
- extraction of the signed digest:
  - classic Authenticode `SpcIndirectDataContent` → `DigestInfo.digest`
  - newer MS-OSHARED `SpcIndirectDataContentV2` → `SigDataV1Serialized.sourceHash`
- MS-OVBA §2.4.2 “Contents Hash” binding verification (signature is bound to the VBA project via a
  versioned normalized-data transcript, `ProjectNormalizedData`), exposed via
  `VbaDigitalSignature::binding`

API notes:

- `formula_vba::verify_vba_digital_signature` returns a [`VbaDigitalSignature`] with a coarse
  binding enum (`VbaSignatureBinding::{Bound, NotBound, Unknown}`).
- `formula_vba::verify_vba_digital_signature_bound` returns a [`VbaDigitalSignatureBound`] with a
  richer binding enum (`VbaProjectBindingVerification`) and best-effort debug info (algorithm OID,
  signed digest bytes, computed digest bytes).
- Signature structs also expose `stream_kind: VbaSignatureStreamKind` to identify which
  `DigitalSignature*` variant was selected (needed for variant-aware binding/digest logic,
  especially the `DigitalSignatureExt` / V3 variant).
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
2. **MS-OSHARED `DigSigBlob` wrapper (offset-based)**
   - Some producers wrap the PKCS#7 bytes in a `DigSigBlob` (MS-OSHARED §2.3.2.2).
   - The blob contains a `DigSigInfoSerialized` structure (MS-OSHARED §2.3.2.1) that points at the
     embedded signature buffer via `signatureOffset`/`cbSignature`.
3. **Length-prefixed `DigSigInfoSerialized`-like wrapper/prefix**
   - Many real-world Excel `\x05DigitalSignature*` streams start with a shorter, *length-prefixed*
     header that does **not** match the MS-OSHARED `DigSigInfoSerialized` layout.
   - This on-disk shape is still often referred to as `DigSigInfoSerialized` in the wild, and is
     sometimes attributed to MS-OFFCRYPTO in older references.
   - The structure is little-endian and length-prefixed; parsing it lets us locate the embedded CMS
     payload deterministically instead of scanning for a DER `SEQUENCE`.
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
   - Some real-world files have inconsistent size fields or additional unknown fields. A permissive
     parser can often still recover by assuming the signature is the **final** blob and locating it
     by counting back from the end of the stream using `cbSignature`.
   - ⚠️ The `cbSigningCertStore` blob may itself contain a PKCS#7/CMS structure (often beginning
     with `0x30`). This means naive “scan for the first PKCS#7 SignedData” logic can pick the
     **certificate store** rather than the actual signature. Prefer the header-derived `(offset, len)`
     when available, and when falling back to scanning heuristics, prefer the **last** plausible
     `SignedData` candidate in the stream. Also validate the inner `SignedData` structure
     (version/digestAlgorithms/encapContentInfo) to reduce false positives.
4. **Detached `content || pkcs7`**
   - The stream contains `signed_content_bytes` followed by a detached CMS signature (`pkcs7_der`).
   - Verification must pass the prefix bytes as the detached content when verifying the CMS blob.

## Extracting the “signed digest” (Authenticode-style)

Excel’s VBA signature is Authenticode-like: the CMS `SignedData` encapsulates a structure that
includes the digest we need for MS-OVBA binding.

High-level extraction steps:

1. **Obtain the CMS `ContentInfo` bytes**
    - If the stream is raw CMS DER, use it directly.
    - If the stream uses an Office DigSig wrapper, unwrap it:
      - MS-OSHARED `DigSigBlob` / `DigSigInfoSerialized` (offset-based), or
      - the length-prefixed DigSigInfoSerialized-like prefix (common in the wild).
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
4. **Parse the encapsulated signed-content structure**
   - Classic Authenticode uses `SpcIndirectDataContent` and stores the digest in
      `messageDigest: DigestInfo`:
      - `DigestInfo.digestAlgorithm.algorithm` (OID; informational for v1/v2 binding, but meaningful for v3 / `DigitalSignatureExt`)
      - `DigestInfo.digest` (digest bytes)
   - Newer Office VBA signatures can use `SpcIndirectDataContentV2`, which stores the digest in
     `SigDataV1Serialized.sourceHash` (MS-OSHARED).

That `(hash_oid, digest_bytes)` pair is the “signed digest” we extract from the signature.

- For v1/v2 signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`), the digest bytes
  we bind against are an **MD5 digest (16 bytes)** per MS-OSHARED §4.3, even if `hash_oid` indicates
  SHA-256.
- For v3 signature streams (`\x05DigitalSignatureExt`), the digest bytes are a **SHA-256 digest (32
  bytes)** over the v3 `ProjectNormalizedData` transcript (MS-OVBA §2.4.2.5/§2.4.2.6).

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
`DigestInfo.digestAlgorithm.algorithm` is *not* the MD5 OID (see MS-OSHARED §4.3, and the comments
in `crates/formula-vba/src/signature.rs` describing legacy v1/v2 binding). This is why binding code
should primarily trust the *digest length* and the stream variant, not the `DigestInfo` algorithm
OID, when choosing how to compute the expected digest bytes.

Evidence in this repo: the regression test
[`crates/formula-vba/tests/signature_binding_md5_sha256.rs`](../crates/formula-vba/tests/signature_binding_md5_sha256.rs)
constructs a signature where `DigestInfo.digestAlgorithm.algorithm` advertises SHA-256 but the digest
bytes are MD5, and asserts that binding verification still succeeds.

#### §2.4.2.1 ContentNormalizedData (v1)

`ContentNormalizedData` is the v1 transcript for module source + selected project metadata / reference
record data.

Transcript construction (as defined by MS-OVBA pseudocode; see also
`formula_vba::content_normalized_data` in `crates/formula-vba/src/contents_hash.rs`):

1. Read the `VBA/dir` stream.
2. Decompress it using the MS-OVBA `CompressedContainer` algorithm (§2.4.1).
3. Parse the decompressed bytes as a sequence of “dir records”:
   - `id: u16le`
   - `len: u32le`
   - `data: [u8; len]`
4. Initialize `ContentNormalizedData = []`.
5. As you scan records in order:
    - For each **`PROJECTNAME.ProjectName`** record (`id == 0x0004`), append the record `data` bytes
      verbatim to `ContentNormalizedData` (**in `VBA/dir` record order**).
    - For each **`PROJECTCONSTANTS.Constants`** record (`id == 0x000C`), append the record `data`
      bytes verbatim to `ContentNormalizedData` (**in `VBA/dir` record order**).
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

Tests in this repo:
[`crates/formula-vba/tests/contents_hash.rs`](../crates/formula-vba/tests/contents_hash.rs)
exercises record inclusion/order (including `PROJECTNAME`/`PROJECTCONSTANTS`), reference normalization,
module ordering, and module source normalization.

#### §2.4.2.2 FormsNormalizedData

`FormsNormalizedData` is a transcript over “designer” streams (UserForms, etc).

Transcript construction (see also `formula_vba::forms_normalized_data` in
`crates/formula-vba/src/normalized_data.rs`):

1. Read the `PROJECT` stream (text) and discover designer modules by extracting every `BaseClass=...`
   line (MS-OVBA §2.3.1.7). Each `BaseClass` value is a *module identifier* for a designer/UserForm
   module.
2. Read and decompress `VBA/dir`, then use its module record groups to map each `BaseClass` module
   identifier to the module’s `MODULESTREAMNAME`. (`formula-vba` matches by name, falling back to an
   ASCII case-insensitive comparison.)
3. Treat that `MODULESTREAMNAME` as the name of the **root-level designer storage** (MS-OVBA §2.2.10
   requires that such a storage exists at the OLE root).
4. For each discovered designer storage (deduplicated if `PROJECT` contains duplicates):
   1. Traverse the storage recursively. To approximate MS-OVBA’s “stored order” in a deterministic
      way (because our CFB/OLE library does not expose raw sibling ordering), `formula-vba` sorts each
      storage’s immediate children by **case-insensitive entry name** (tie-breaking by the original
      name), then processes them depth-first.
   2. For each stream encountered, append its bytes in **1023-byte blocks**:
      - Split the stream bytes into chunks of up to 1023 bytes.
      - For each chunk, append exactly 1023 bytes: the chunk bytes followed by `0x00` padding for the
        remainder of the block.
      (So a stream of length `n` contributes `1023 * ceil(n/1023)` bytes; the final partial block is
      zero-padded.)

Tests in this repo:
[`crates/formula-vba/tests/forms_normalized_data.rs`](../crates/formula-vba/tests/forms_normalized_data.rs)
covers the 1023-byte padding behavior and deterministic traversal ordering.

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
3. Append the v3-included reference record payload bytes from `VBA/dir`:
   - v1 reference types:
     - **`REFERENCEREGISTERED`** (`0x000D`): append raw `data` bytes
     - **`REFERENCEPROJECT`** (`0x000E`): append normalized bytes (same TempBuffer + “copy until NUL”
       rule as §2.4.2.1)
   - additional v3 reference types (raw `data` bytes):
     - **`REFERENCECONTROL`** (`0x002F`)
     - **`REFERENCEEXTENDED`** (`0x0030`)
     - **`REFERENCEORIGINAL`** (`0x0033`)
4. For each module record group (in `VBA/dir` stored order), append **module metadata bytes before
   the normalized source**:
   1. `MODULENAME` (`0x0019`) record `data` bytes
   2. `MODULESTREAMNAME` (`0x001A`) record `data` bytes, with the trailing reserved `u16` trimmed when
      present (many files store `MODULESTREAMNAME || 0x0000`)
   3. `MODULETYPE` (`0x0021`) record `data` bytes
   4. Then append the module’s normalized source bytes (same newline + `Attribute` stripping rules as
      §2.4.2.1).

The main point for implementers: this is **not** just “`ContentNormalizedData` hashed with a stronger
algorithm”. v3 defines a distinct transcript (`V3ContentNormalizedData`) and `DigitalSignatureExt`
uses SHA-256 over `ProjectNormalizedData` (§2.4.2.6).

Repo status:

- `formula-vba` implements the v3 transcript and digest:
  - `formula_vba::v3_content_normalized_data` and `formula_vba::project_normalized_data_v3` in
    `crates/formula-vba/src/contents_hash.rs`
  - `formula_vba::contents_hash_v3` (SHA-256 digest bytes for `DigitalSignatureExt`) in
    `crates/formula-vba/src/contents_hash.rs`
  - `formula_vba::compute_vba_project_digest_v3` in `crates/formula-vba/src/project_digest.rs`
- Binding logic in `crates/formula-vba/src/signature.rs` treats `\x05DigitalSignatureExt` as a v3
  signature stream and compares its signed digest bytes against the computed v3 project digest
  (hashing v3 `ProjectNormalizedData` with the `DigestInfo` algorithm OID; typically SHA-256, 32 bytes).

Tests in this repo:
[`crates/formula-vba/tests/contents_hash_v3.rs`](../crates/formula-vba/tests/contents_hash_v3.rs)
exercises v3 reference record inclusion, module metadata inclusion, module ordering, and end-to-end
SHA-256 digest computation.

#### §2.4.2.6 ProjectNormalizedData (the hashed transcript)

`ProjectNormalizedData` is the spec-defined transcript that is actually hashed to produce the digest
bytes placed in `DigestInfo.digest`.

Per MS-OVBA §2.4.2.6, the transcript depends on the **contents-hash version**:

- v1 (`DigitalSignature`): `ProjectNormalizedData = ContentNormalizedData`
- v2 (`DigitalSignatureEx`): `ProjectNormalizedData = ContentNormalizedData || FormsNormalizedData`
- v3 (`DigitalSignatureExt`): `ProjectNormalizedData = V3ContentNormalizedData || FormsNormalizedData`

### `formula-vba` binding implementation (current)

To bind the signature to the VBA project contents, `formula-vba` currently:

1. Extracts the signed digest structure from the signature payload, yielding:
   - `hash_oid` (typically `DigestInfo.digestAlgorithm.algorithm`), and
   - `digest_bytes` (the signed digest bytes).
2. Chooses the MS-OVBA §2.4.2 transcript version based on the signature stream variant (`stream_kind`):
   - `DigitalSignature` / `DigitalSignatureEx` (v1/v2):
     - compute **Content Hash** (v1) and **Agile Content Hash** (v2) as MD5 digests over
       `ContentNormalizedData` and `ContentNormalizedData || FormsNormalizedData`.
     - Per MS-OSHARED §4.3, the digest bytes are expected to be **MD5 (16 bytes)** even when
       `hash_oid` indicates SHA-256 (the OID is informational for v1/v2 binding).
   - `DigitalSignatureExt` (v3):
     - compute the v3 digest by hashing v3 `ProjectNormalizedData` (`V3ContentNormalizedData || FormsNormalizedData`)
       with the digest algorithm indicated by `hash_oid` (expected SHA-256, 32 bytes).
3. Compares the computed digest bytes to `digest_bytes` to determine the binding result
   (`VbaSignatureBinding::{Bound, NotBound, Unknown}`).

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

- For v1/v2 signature streams (`\x05DigitalSignature` / `\x05DigitalSignatureEx`), per MS-OSHARED
  §4.3, Office stores the VBA project digest used for binding as **MD5 (16 bytes)**, even if the
  PKCS#7/CMS signature uses SHA-1/SHA-256 and even if `DigestInfo.digestAlgorithm.algorithm`
  indicates SHA-256. For these streams, the digest algorithm OID is treated as informational when
  selecting the expected digest bytes.
- For v3 signature streams (`\x05DigitalSignatureExt`), Office uses **SHA-256 (32 bytes)** over the
  v3 transcript (`V3ContentNormalizedData || FormsNormalizedData`), and `DigestInfo` is expected to
  indicate SHA-256.
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
- `DigitalSignatureExt` / v3 contents hashes are supported, but correctness depends on matching
  MS-OVBA's v3 transcript details. Future work: validate against additional real-world Excel fixtures
  over time.
- Callers should treat `VbaSignatureBinding::Unknown` as "could not verify binding", not as "bound".
- Treat `binding == Bound` as a strong signal, but validate against Excel fixtures before relying on
  it as a hard security boundary.

## Repo implementation pointers

If you need to update or extend signature handling, start with:

- `crates/formula-vba/src/signature.rs`
  - signature stream discovery (`\x05DigitalSignature*`)
  - CMS verification and the `VbaDigitalSignature::binding` decision
- `crates/formula-vba/src/offcrypto.rs`
  - MS-OSHARED `DigSigBlob` parsing + the length-prefixed DigSigInfoSerialized-like wrapper
    (deterministic CMS offset/length)
- `crates/formula-vba/src/authenticode.rs`
  - `SignedData.encapContentInfo.eContent` parsing and `SpcIndirectDataContent` → `DigestInfo`
- `crates/formula-vba/src/contents_hash.rs`
  - MS-OVBA §2.4.2.1 `ContentNormalizedData` (`content_normalized_data`)
  - MS-OVBA §2.4.2.5 `V3ContentNormalizedData` (`v3_content_normalized_data`)
  - MS-OVBA §2.4.2.6 v3 `ProjectNormalizedData` (`project_normalized_data_v3`)
- `crates/formula-vba/src/normalized_data.rs`
  - MS-OVBA §2.4.2.2 `FormsNormalizedData` (`forms_normalized_data`)
- `crates/formula-vba/src/project_digest.rs`
  - `compute_vba_project_digest_v3` (MS-OVBA §2.4.2 v3 digest for `DigitalSignatureExt`)
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
  - constructs an Authenticode-like `SpcIndirectDataContent` payload with an embedded digest, signs
    it, and checks that tampering with non-signature streams flips `VbaSignatureBinding` to
    `NotBound`.
- `crates/formula-vba/tests/signature_binding_md5_sha256.rs`
  - regression: verifies MS-OSHARED §4.3 behavior where `DigestInfo.digest` is MD5 bytes even when
    `DigestInfo.digestAlgorithm.algorithm` advertises SHA-256.
- `crates/formula-vba/tests/contents_hash_v3.rs` and `crates/formula-vba/tests/signature_binding_v3.rs`
  - cover v3 transcript construction and `\x05DigitalSignatureExt` binding behavior.
- `crates/formula-vba/tests/digsig_blob.rs`
  - verifies that MS-OSHARED `DigSigBlob`-wrapped signatures are parsed deterministically (without
    relying on DER scanning heuristics).

## Specs / references

- **MS-OVBA**: VBA project storage and VBA project digest computation / binding verification.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- **MS-OSHARED**: VBA digital signature storage (`DigSigInfoSerialized`) and the MD5 “VBA project hash” rule (MS-OSHARED §4.3).
  - `DigSigInfoSerialized`: MS-OSHARED §2.3.2.1
  - `DigSigBlob`: MS-OSHARED §2.3.2.2
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/30a00273-dbee-422f-b488-f4b8430ae046
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/bc21c922-b7ae-4736-90aa-86afb6403462
- **RFC 5652**: Cryptographic Message Syntax (CMS) (PKCS#7 `SignedData`).
  - https://www.rfc-editor.org/rfc/rfc5652
