# VBA Digital Signatures (vbaProject.bin / vbaProjectSignature.bin)

This document captures how Excel/VBA macro signatures are stored in XLSM files, and how we extract
and **bind** those signatures against the MS-OVBA **Contents Hash** (specifically **ContentsHashV3**
used by `DigitalSignatureExt`).

In this repo, `crates/formula-vba` implements best-effort:

- signature stream discovery + PKCS#7/CMS internal verification (`verify_vba_digital_signature`)
- extraction of the signed binding digest from the Authenticode-style signed content
  - classic `SpcIndirectDataContent` → `DigestInfo.digest`
  - (some producers) MS-OSHARED `SpcIndirectDataContentV2` → `SigDataV1Serialized.sourceHash`
- MS-OVBA “Contents Hash” binding verification (signature is bound to the VBA project)

API notes:

- `formula_vba::verify_vba_digital_signature` returns a [`VbaDigitalSignature`] with a coarse
  binding enum (`VbaSignatureBinding::{Bound, NotBound, Unknown}`) and the detected signature stream
  variant (`stream_kind: VbaSignatureStreamKind`).
- `formula_vba::verify_vba_digital_signature_bound` returns a [`VbaDigitalSignatureBound`] with a
  richer binding enum (`VbaProjectBindingVerification`) and best-effort debug info (the `DigestInfo`
  algorithm OID, signed digest bytes, computed digest bytes).
  - Note: for **legacy** VBA signatures (`DigitalSignature` / `DigitalSignatureEx`), the *binding*
    digest bytes are always a 16-byte MD5 (MS-OSHARED §4.3), even when
    `DigestInfo.digestAlgorithm.algorithm` indicates SHA-256.
    - Practical implication: do not select the v1/v2 binding algorithm from the OID; always compute
      **MD5** over the correct MS-OVBA transcript and compare it to the 16 digest bytes.
    - For `DigitalSignatureExt` (MS-OVBA v3), the signed digest bytes are commonly 32-byte SHA-256.
      `formula-vba` currently verifies v3 binding by computing a **SHA-256** digest over its
      best-effort v3 transcript (`contents_hash_v3`) and comparing it to the signed digest bytes.
      Producers that use a different digest algorithm may therefore be reported as `NotBound`.
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
3. **MS-OSHARED `WordSigBlob` wrapper (offset-based, Unicode length prefix)**
   - Some producers wrap the PKCS#7 bytes in a `WordSigBlob` (MS-OSHARED §2.3.2.3).
   - `WordSigBlob` is similar to `DigSigBlob` but starts with `cch: u16` (a UTF-16 character count /
     half the byte count of the remainder of the structure).
   - The embedded `DigSigInfoSerialized` still provides offsets to the signature buffer, but the
     offsets are **relative to the start of the `cbSigInfo` field** (byte offset 2), not the start
     of the structure.
4. **Length-prefixed `DigSigInfoSerialized`-like wrapper/prefix**
   - Many real-world Excel `\x05DigitalSignature*` streams start with a shorter, *length-prefixed*
     header that does **not** match the MS-OSHARED `DigSigInfoSerialized` layout.
   - The structure is little-endian and length-prefixed; parsing it lets us locate the embedded CMS
     payload deterministically instead of scanning for a DER `SEQUENCE`.
   - ⚠️ The wrapper's `cbSigningCertStore` region may itself contain a PKCS#7/CMS structure (often
     beginning with `0x30`). When scanning heuristically, prefer the **last** plausible `SignedData`
     candidate in the stream and validate the inner `SignedData` structure.
5. **Detached `content || pkcs7`**
   - The stream contains `signed_content_bytes` followed by a detached CMS signature (`pkcs7_der`).
   - Verification must pass the prefix bytes as the detached content when verifying the CMS blob.

## Extracting the “signed digest” (Authenticode-style)

Excel’s VBA signature is Authenticode-like: the CMS `SignedData` encapsulates a structure that
includes the digest we need for MS-OVBA binding.

High-level extraction steps:

1. **Obtain the CMS `ContentInfo` bytes**
   - If the stream is raw CMS DER, use it directly.
   - If the stream uses an Office DigSig wrapper, unwrap it (MS-OSHARED `DigSigBlob` / `WordSigBlob`,
     or the length-prefixed DigSigInfoSerialized-like prefix).
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
     - `DigestInfo.digestAlgorithm.algorithm` (OID; not authoritative for v1/v2 binding)
     - `DigestInfo.digest` (digest bytes)
   - Newer Office producers can use `SpcIndirectDataContentV2`, which stores the digest in
     `SigDataV1Serialized.sourceHash` (MS-OSHARED).

That `(hash_oid, digest_bytes)` pair is the “signed digest” we use to bind a VBA signature to a
specific VBA project.

- For legacy v1/v2 streams (`DigitalSignature` / `DigitalSignatureEx`), `digest_bytes` are always a
  16-byte MD5 per MS-OSHARED §4.3, so the OID is informational only for binding.
- For v3 streams (`DigitalSignatureExt`), `digest_bytes` are commonly 32-byte SHA-256; `formula-vba`
  currently computes a SHA-256 digest for v3 binding verification.

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

## Binding (MS-OVBA Contents Hash / ContentsHashV3 verification)

CMS signature verification alone answers “is this a valid CMS signature over *some bytes*?”, but it
does not, by itself, prove that the signature is bound to the rest of the VBA project.

### Legacy MD5-always rule (v1/v2)

For **legacy** VBA signatures, the digest embedded in:

- classic `SpcIndirectDataContent.messageDigest.digest`, or
- MS-OSHARED `SigDataV1Serialized.sourceHash`

is the value we compare against the computed MS-OVBA binding digest.

In practice, the expected digest algorithm/length depends on the signature stream variant:

Practical implications:

- For v1/v2 (`DigitalSignature` / `DigitalSignatureEx`), the `DigestInfo` *algorithm OID* should be
  treated as informational only for VBA binding; the verifier must compute **MD5** and compare it
  to the 16 digest bytes (MS-OSHARED §4.3).
- For v3 (`DigitalSignatureExt`), Office uses the MS-OVBA v3 transcript and the digest algorithm OID
  is meaningful (SHA-256 is common in the wild). `formula-vba` currently verifies v3 binding by
  computing **SHA-256** via `contents_hash_v3` and comparing it to the signed digest bytes.

The v1/v2 MD5 behavior is specified in **MS-OSHARED §4.3**:
https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/40c8dab3-e8db-4c66-a6be-8cec06351b1e

### Spec-correct transcripts

MS-OVBA defines a **versioned** “Contents Hash” computation (§2.4.2) used to bind a
`\x05DigitalSignature*` stream to the rest of `vbaProject.bin`.

At a high level, MS-OVBA defines both:

- how to build a deterministic transcript of project bytes (“normalized data”), and
- how to hash that transcript to obtain the binding digest embedded in the CMS/Authenticode payload.

For v1/v2, Office’s *binding digest bytes* are MD5 per MS-OSHARED §4.3 (see above). For v3, the
important spec detail is the **transcript definition and concatenation order**.

- **v1 / `DigitalSignature`**: `ContentNormalizedData` (MS-OVBA §2.4.2.1)
- **v2 / `DigitalSignatureEx`**: `ContentNormalizedData || FormsNormalizedData` (MS-OVBA §2.4.2.1 + §2.4.2.2)
- **v3 / `DigitalSignatureExt`**: `ContentBuffer = V3ContentNormalizedData || ProjectNormalizedData` (MS-OVBA §2.4.2.7)

Digest algorithm note (v3):

- MS-OVBA §2.4.2.7 defines the **V3 Content Hash** as the digest of `ContentBuffer`, but the spec
  pseudocode is written in terms of a generic hash function. In real-world `DigitalSignatureExt`
  streams, **SHA-256 (32 bytes)** is commonly observed, but it is not safe to assume this is a
  universal constraint across all producers.

Reference record handling note:

- `ContentNormalizedData` (v1/v2) and `V3ContentNormalizedData` (v3) each incorporate only a subset
  of reference-related `VBA/dir` records, but the allowlist and normalization rules differ between
  versions.
- v1/v2 (`ContentNormalizedData`) excludes `REFERENCENAME` (`0x0016`) and normalizes some reference
  records via MS-OVBA’s TempBuffer + “copy until first NUL byte” rule.
- v3 (`V3ContentNormalizedData`) includes `REFERENCENAME` (`0x0016`) (and its optional UTF-16LE
  `NameUnicode` record, `0x003E`) and incorporates reference records via explicit field-selection
  rules (§2.4.2.5), rather than TempBuffer/copy-until-NUL.

V3 spec references:

- §2.4.2.5 V3ContentNormalizedData:
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/becd5647-d4e9-4d7d-ab86-484421a086eb
- §2.4.2.6 ProjectNormalizedData:
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/eda9d57a-a862-4927-9554-6750dada9b37
- §2.4.2.7 ContentsHashV3:
  https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/601a4412-00cc-46a0-b8e0-3001c011308e

#### V3ContentNormalizedData (MS-OVBA §2.4.2.5)

`V3ContentNormalizedData` is derived from the decompressed `VBA/dir` stream plus each referenced
module stream (`VBA/<ModuleStreamName>`). Key points from the MS-OVBA §2.4.2.5 pseudocode:

- **Reference records:** v3 explicitly appends record IDs / sizes / fields. It does **not** use the
  older v1/v2-style TempBuffer + “copy until first NUL byte” rule for references.
  - Importantly, `REFERENCENAME` (`0x0016`) **does contribute** for v3 (and may be followed by a
    Unicode name record, `0x003E`, which also contributes when present).
- **Module ordering:** modules are processed in the order specified by `VBA/dir` (not alphabetical,
  and not OLE directory enumeration order).
- **Module source normalization:** §2.4.2.5 uses `LF` line endings and has v3-specific `Attribute ...`
  handling (different from v1/v2).

#### ProjectNormalizedData (MS-OVBA §2.4.2.6, `NormalizeProjectStream`)

`ProjectNormalizedData` (MS-OVBA §2.4.2.6) is derived from the textual `PROJECT` stream (not from
`VBA/dir` records). The spec pseudocode normalizes:

- `ProjectProperties` (key/value lines before the first `[Section]` header), excluding `ID=...`
- the optional `[Host Extender Info]` section
- and explicitly ignores the optional `[Workspace]` / ProjectWorkspace section, so machine-local state
  does not affect hashing/signature binding.

#### V3 Content Hash / v3 ContentBuffer (MS-OVBA §2.4.2.7)

MS-OVBA v3 defines the content-hash input as:

```text
ContentBuffer = V3ContentNormalizedData || ProjectNormalizedData
V3ContentHash = Hash(ContentBuffer)
```

Where `Hash` is the digest function used for the signature binding digest (SHA-256 is common in the
wild for `DigitalSignatureExt`, but not something this repo treats as a spec-guaranteed constant).

### `formula-vba` implementation notes (v3)

`formula-vba`’s v3 handling is currently **best-effort** and some helpers/fixtures in this repo do
not yet match the MS-OVBA §2.4.2.6/§2.4.2.7 pseudocode exactly.

- `formula_vba::v3_content_normalized_data` aims to implement MS-OVBA §2.4.2.5.
- `formula_vba::project_normalized_data_v3_transcript` currently returns a library-specific transcript:
  `(filtered PROJECT stream lines) || V3ContentNormalizedData || FormsNormalizedData`.
  This ordering and `PROJECT`-stream filtering do **not** match the spec definitions of
  `ProjectNormalizedData` (§2.4.2.6) or `ContentBuffer` (§2.4.2.7).
- `formula_vba::contents_hash_v3()` always computes **SHA-256** over that library-specific transcript.
  Treat this as a convenience helper for common Office behavior, not as a spec guarantee.
- `formula_vba::forms_normalized_data` builds `FormsNormalizedData`.
- `formula_vba::project_normalized_data_v3` (alias: `project_normalized_data_v3_dir_records`) is a
  metadata-only dir-record transcript derived from selected `VBA/dir` records; it is **not**
  MS-OVBA §2.4.2.6 `ProjectNormalizedData`.

#### `formula_vba::project_normalized_data_v3_dir_records` helper (dir record allowlist)

After MS-OVBA decompression, `VBA/dir` is a sequence of:

```text
id:  u16le
len: u32le
data: [u8; len]
```

This helper concatenates **normalized `data` bytes only**. The `(id, len)` header bytes are never
included.

#### `VBA/dir` record IDs included (project info + module metadata)

All records are processed in the **stored order** from `VBA/dir`.

Project info records:

- Fixed-size records included verbatim: `0x0001`, `0x0002`, `0x0003`, `0x0007`, `0x0008`, `0x0009`,
  `0x0014`
  - (`PROJECTSYSKIND`, `PROJECTLCID`, `PROJECTCODEPAGE`, `PROJECTHELPCONTEXT`, `PROJECTLIBFLAGS`,
    `PROJECTVERSION`, `PROJECTLCIDINVOKE`)
- String records (Unicode preferred when present for records that have both forms):
  - `0x0004` `PROJECTNAME` (ANSI; this helper does not currently include a Unicode variant)
  - `0x0005` `PROJECTDOCSTRING` (ANSI)
  - `0x0040` `PROJECTDOCSTRINGUNICODE` (Unicode/alternate; UTF-16LE payload extraction)
  - `0x0006` `PROJECTHELPFILEPATH` (ANSI)
  - `0x003D` `PROJECTHELPFILEPATH2` (Unicode/alternate; UTF-16LE payload extraction)
  - `0x000C` `PROJECTCONSTANTS` (ANSI)
  - `0x003C` `PROJECTCONSTANTSUNICODE` (Unicode/alternate; UTF-16LE payload extraction)

Module metadata records (for each module record group, in stored order; a new module group begins at
`MODULENAME` (`0x0019`)):

- String records with ANSI/Unicode variants (Unicode preferred when present):
  - `0x0019` `MODULENAME` (ANSI)
  - `0x0047` `MODULENAMEUNICODE` (Unicode)
  - `0x001A` `MODULESTREAMNAME` (ANSI; see reserved trimming below)
  - `0x0032` `MODULESTREAMNAMEUNICODE` (Unicode/alternate)
  - `0x001C` `MODULEDOCSTRING` (ANSI)
  - `0x0048` `MODULEDOCSTRINGUNICODE` (Unicode/alternate)
- Non-string module metadata records included verbatim: `0x001E`, `0x0021`, `0x0025`, `0x0028`
  - (e.g. module help context / type / readonly / private flags per MS-OVBA)

#### Unicode-vs-ANSI selection rule + Unicode payload extraction

For string-like fields that have both ANSI/MBCS and Unicode record variants:

- If the Unicode record exists, the ANSI record does **not** contribute.
- Unicode record payloads are UTF-16LE bytes. Some producers also embed an *internal* `u32le` length
  prefix before the UTF-16LE bytes (with the length interpreted as either code units or bytes).
  `formula-vba` strips that internal length prefix only when it is consistent with the record
  length; otherwise it treats the full payload as raw UTF-16LE bytes.

#### `MODULESTREAMNAME` reserved trimming

For the ANSI `MODULESTREAMNAME` record (`0x001A`), some producers append a trailing reserved `u16`
(`0x0000`). This helper trims this reserved `u16` before appending the bytes.

### `formula-vba` binding implementation (high level)

To bind the signature to the VBA project contents, `formula-vba`:

1. Extracts the signed digest bytes from the signature payload.
2. Computes the appropriate Contents Hash transcript:
   - v1/v2: per MS-OVBA §2.4.2.1/§2.4.2.2
   - v3: best-effort (see “`formula-vba` implementation notes (v3)” above)
3. Computes digest bytes for that transcript:
   - v1 (`DigitalSignature`): compute **MD5** of `ContentNormalizedData` (MS-OSHARED §4.3; ignore the
     `DigestInfo` OID for binding)
   - v2 (`DigitalSignatureEx`): compute **MD5** of (`ContentNormalizedData || FormsNormalizedData`)
     (MS-OSHARED §4.3; ignore the `DigestInfo` OID for binding)
   - v3 (`DigitalSignatureExt`): currently compute **SHA-256** via `formula_vba::contents_hash_v3`
     (best-effort; SHA-256 is common for `DigitalSignatureExt`, but not a spec-guaranteed constant)
   - When the signature stream kind is unknown (for example, a raw PKCS#7/CMS blob from
     `vbaProjectSignature.bin`), `formula-vba` best-effort attempts v3 binding first, then falls back
     to legacy binding.
4. Compares the computed digest bytes to the signed digest bytes.

Result interpretation:

- If CMS verification fails ⇒ signature invalid (`VbaSignatureVerification::SignedInvalid` /
  `SignedParseError`).
- If CMS verification succeeds and binding matches ⇒ signature is verified and bound
  (`VbaSignatureBinding::Bound`).
- If CMS verification succeeds and binding mismatches ⇒ signature is present but **not bound** to the
  current VBA project bytes (`VbaSignatureBinding::NotBound`).
- If CMS verification succeeds but binding cannot be evaluated (missing/unparseable data, incomplete
  project bytes, etc.) ⇒ binding is conservative/unknown (`VbaSignatureBinding::Unknown`).

## Repo implementation pointers

If you need to update or extend signature handling, start with:

- `crates/formula-vba/src/signature.rs`
  - signature stream discovery (`\x05DigitalSignature*`)
  - CMS verification and the `VbaDigitalSignature::binding` decision
- `crates/formula-vba/src/offcrypto.rs`
  - MS-OSHARED / MS-OFFCRYPTO DigSig wrapper parsing (deterministic CMS offset/length)
- `crates/formula-vba/src/authenticode.rs`
  - `SignedData.encapContentInfo.eContent` parsing and `SpcIndirectDataContent` → `DigestInfo`
- `crates/formula-vba/src/contents_hash.rs`
  - MS-OVBA normalized-data transcript builders
- `crates/formula-vba/src/project_digest.rs`
  - `compute_vba_project_digest` (hash over `ContentNormalizedData || FormsNormalizedData`; equivalent to v1 when `FormsNormalizedData` is empty; strict transcript-only, no raw-stream fallback)
  - `compute_vba_project_digest_v3` (hash over `formula-vba`'s current v3 transcript with a
    caller-selected algorithm; useful for debugging/compatibility; see v3 notes above)

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
- `crates/formula-vba/tests/contents_hash_v3.rs`, `crates/formula-vba/tests/signature_binding_v3.rs`,
  and `crates/formula-vba/tests/signature_binding_v3_ext.rs`
  - cover v3 transcript construction and `\x05DigitalSignatureExt` binding behavior (SHA-256
    `ContentsHashV3`), including the “ignore DigestInfo OID and compare digest bytes” rule.
- `crates/formula-vba/tests/digsig_blob.rs`
  - verifies that MS-OSHARED `DigSigBlob`-wrapped signatures are parsed deterministically (without
    relying on DER scanning heuristics).
- `crates/formula-vba/tests/wordsig_blob.rs`
  - verifies that MS-OSHARED `WordSigBlob`-wrapped signatures are parsed deterministically (without
    relying on DER scanning heuristics).

## Specs / references

- **MS-OVBA**: VBA project storage, signature streams, and Contents Hash (including `ContentsHashV3`)
  computation.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
- **MS-OSHARED**: Office shared structures; documents the MD5-always “VBA project source hash”
  behavior used by v1/v2 (legacy) VBA signatures.
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/40c8dab3-e8db-4c66-a6be-8cec06351b1e
- **MS-OFFCRYPTO**: Office cryptography structures (historical references; some real-world wrappers
  are MS-OSHARED).
  - https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
- **RFC 5652**: Cryptographic Message Syntax (CMS) (PKCS#7 `SignedData`).
  - https://www.rfc-editor.org/rfc/rfc5652
