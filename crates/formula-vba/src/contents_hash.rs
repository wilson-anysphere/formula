use encoding_rs::{Encoding, UTF_16LE, WINDOWS_1252};
use md5::Md5;
use sha2::Digest as _;
use sha2::Sha256;

use crate::forms_normalized_data;
use crate::{decompress_container, DirParseError, DirStream, OleFile, ParseError};

#[derive(Debug, Clone, Default)]
struct ModuleInfo {
    stream_name: String,
    text_offset: Option<usize>,
}

/// Build the MS-OVBA `ProjectNormalizedData` byte sequence (used by "Contents Hash V3").
///
/// This is derived from the decompressed `VBA/dir` stream by iterating records in stored order and
/// concatenating the **record data bytes** (excluding the 6-byte record header: `id` + `len`) for
/// the subset of `dir` record types specified by MS-OVBA.
///
/// Spec reference: MS-OVBA §2.4.2 "ProjectNormalizedData" (Contents Hash V3).
pub fn project_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Project information record IDs (MS-OVBA 2.3.4.2.1).
    //
    // Note: `VBA/dir` stores records as: u16 id, u32 len, `len` bytes of record data.
    const PROJECTSYSKIND: u16 = 0x0001;
    const PROJECTLCID: u16 = 0x0002;
    const PROJECTCODEPAGE: u16 = 0x0003;
    const PROJECTNAME: u16 = 0x0004;
    const PROJECTDOCSTRING: u16 = 0x0005;
    const PROJECTHELPFILEPATH: u16 = 0x0006;
    const PROJECTHELPCONTEXT: u16 = 0x0007;
    const PROJECTLIBFLAGS: u16 = 0x0008;
    const PROJECTVERSION: u16 = 0x0009;
    const PROJECTCONSTANTS: u16 = 0x000C;
    const PROJECTLCIDINVOKE: u16 = 0x0014;

    // Some producers emit a second, Unicode form record immediately following the ANSI form.
    // The MS-OVBA pseudocode prefers the Unicode form when present.
    //
    // These record IDs are not part of the minimal VBA project fixture used by this repo, but they
    // occur in real-world files.
    const PROJECTDOCSTRINGUNICODE: u16 = 0x0040;
    const PROJECTCONSTANTSUNICODE: u16 = 0x003C;

    let mut ole = OleFile::open(vba_project_bin)?;

    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;

    let mut out = Vec::new();

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        if offset + 6 > dir_decompressed.len() {
            return Err(DirParseError::Truncated.into());
        }

        let id = u16::from_le_bytes([dir_decompressed[offset], dir_decompressed[offset + 1]]);
        let len = u32::from_le_bytes([
            dir_decompressed[offset + 2],
            dir_decompressed[offset + 3],
            dir_decompressed[offset + 4],
            dir_decompressed[offset + 5],
        ]) as usize;
        offset += 6;
        if offset + len > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = &dir_decompressed[offset..offset + len];
        offset += len;

        let next_id = peek_next_record_id(&dir_decompressed, offset);

        match id {
            PROJECTSYSKIND
            | PROJECTLCID
            | PROJECTLCIDINVOKE
            | PROJECTCODEPAGE
            | PROJECTNAME
            | PROJECTHELPFILEPATH
            | PROJECTHELPCONTEXT
            | PROJECTLIBFLAGS
            | PROJECTVERSION => {
                out.extend_from_slice(data);
            }

            PROJECTDOCSTRING => {
                if next_id != Some(PROJECTDOCSTRINGUNICODE) {
                    out.extend_from_slice(data);
                }
            }
            PROJECTDOCSTRINGUNICODE => {
                out.extend_from_slice(data);
            }

            PROJECTCONSTANTS => {
                if next_id != Some(PROJECTCONSTANTSUNICODE) {
                    out.extend_from_slice(data);
                }
            }
            PROJECTCONSTANTSUNICODE => {
                out.extend_from_slice(data);
            }

            _ => {}
        }
    }

    Ok(out)
}

/// Build the MS-OVBA "ContentNormalizedData" byte sequence for a VBA project.
///
/// This is a building block used by MS-OVBA when computing the VBA signature binding digest
/// ("Contents Hash") that a `\x05DigitalSignature*` stream signs.
///
/// This implementation is intentionally focused on correctness for the trickier normalization
/// rules that are easy to regress:
///
/// - **Module ordering** comes from the stored order in `VBA/dir` (`PROJECTMODULES.Modules` order),
///   not alphabetical sorting and not OLE directory enumeration order.
/// - **Module source normalization** treats CR and lone-LF as line breaks, ignores the LF of CRLF,
///   and strips `Attribute ...` lines (case-insensitive, start-of-line match).
/// - **Project name and constants** are incorporated by appending the raw record payload bytes for
///   `PROJECTNAME.ProjectName` (0x0004) and `PROJECTCONSTANTS.Constants` (0x000C) in `VBA/dir`
///   record order.
/// - **Reference records** are incorporated for a subset of record types, matching the MS-OVBA
///   §2.4.2.1 pseudocode allowlist:
///   - `REFERENCEREGISTERED` (0x000D)
///   - `REFERENCEPROJECT` (0x000E)
///   - `REFERENCECONTROL` (0x002F)
///   - `REFERENCEEXTENDED` (0x0030)
///   - `REFERENCEORIGINAL` (0x0033)
///   Other reference-related records (e.g. `REFERENCENAME` (0x0016)) MUST NOT contribute.
///
/// Spec reference: MS-OVBA §2.4.2.1 "Content Normalized Data".
pub fn content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let project_stream_bytes = ole.read_stream_opt("PROJECT")?;
    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let encoding = project_stream_bytes
        .as_deref()
        .and_then(crate::detect_project_codepage)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    // Prefer parsing the decompressed `VBA/dir` stream using the MS-OVBA §2.3.4.2 structured record
    // layout. Many synthetic fixtures (including some in our test suite) use a simplified TLV-ish
    // encoding (`u16 id || u32 len || data`) that is not fully spec-accurate; for those we fall
    // back to the existing forgiving parser below.
    //
    // This keeps backwards compatibility for tests/fixtures while enabling correct Contents Hash
    // recomputation for real-world projects that follow the spec record layout (with module/dir
    // terminators, nested records, and reserved/unicode fields).
    if let Some(strict) = content_normalized_data_strict(&mut ole, &dir_decompressed, encoding) {
        return Ok(strict);
    }

    let mut out = Vec::new();
    let mut modules: Vec<ModuleInfo> = Vec::new();
    let mut current_module: Option<ModuleInfo> = None;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        if offset + 6 > dir_decompressed.len() {
            return Err(DirParseError::Truncated.into());
        }

        let id = u16::from_le_bytes([dir_decompressed[offset], dir_decompressed[offset + 1]]);
        let len = u32::from_le_bytes([
            dir_decompressed[offset + 2],
            dir_decompressed[offset + 3],
            dir_decompressed[offset + 4],
            dir_decompressed[offset + 5],
        ]) as usize;
        offset += 6;
        if offset + len > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = &dir_decompressed[offset..offset + len];
        offset += len;

        match id {
            // PROJECTNAME.ProjectName
            0x0004 => out.extend_from_slice(data),

            // PROJECTCONSTANTS.Constants
            0x000C => out.extend_from_slice(data),

            // MS-OVBA §2.4.2.1 ContentNormalizedData reference record allowlist.
            //
            // NOTE: The spec explicitly includes only some REFERENCE* record variants.
            //
            // REFERENCEREGISTERED (0x000D)
            0x000D => out.extend_from_slice(data),

            // REFERENCEPROJECT (0x000E)
            0x000E => out.extend_from_slice(&normalize_reference_project(data)?),

            // REFERENCECONTROL (0x002F)
            0x002F => out.extend_from_slice(&normalize_reference_control(data)?),

            // REFERENCEEXTENDED (0x0030)
            0x0030 => out.extend_from_slice(data),

            // REFERENCEORIGINAL (0x0033)
            0x0033 => out.extend_from_slice(&normalize_reference_original(data)?),

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    modules.push(m);
                }
                current_module = Some(ModuleInfo {
                    stream_name: decode_dir_string(data, encoding),
                    text_offset: None,
                });
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_string(trim_reserved_u16(data), encoding);
                }
            }

            // MODULETEXTOFFSET (u32 LE).
            0x0031 => {
                if let Some(m) = current_module.as_mut() {
                    if data.len() >= 4 {
                        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                        m.text_offset = Some(n);
                    }
                }
            }

            _ => {}
        }
    }

    if let Some(m) = current_module.take() {
        modules.push(m);
    }

    for module in modules {
        let stream_path = format!("VBA/{}", module.stream_name);
        let module_stream = ole
            .read_stream_opt(&stream_path)?
            .ok_or(ParseError::MissingStream("module stream"))?;

        let text_offset = module
            .text_offset
            .unwrap_or_else(|| guess_text_offset(&module_stream));
        let text_offset = text_offset.min(module_stream.len());
        let source_container = &module_stream[text_offset..];
        let source = decompress_container(source_container)?;
        out.extend_from_slice(&normalize_module_source(&source));
    }

    Ok(out)
}

#[derive(Debug, Clone)]
enum ReferenceForHash {
    Registered,
    Project {
        size_of_libid_absolute: u32,
        libid_absolute: Vec<u8>,
        size_of_libid_relative: u32,
        libid_relative: Vec<u8>,
        major_version: u32,
        minor_version: u16,
    },
}

#[derive(Debug, Clone)]
struct ModuleForHash {
    stream_name: String,
    text_offset: u32,
}

#[derive(Debug, Clone)]
struct DirForHash {
    project_name: Vec<u8>,
    project_constants: Vec<u8>,
    references: Vec<ReferenceForHash>,
    modules: Vec<ModuleForHash>,
}

/// Try to parse the decompressed `VBA/dir` stream according to the MS-OVBA §2.3.4.2 structured
/// record layout and compute ContentNormalizedData (MS-OVBA §2.4.2.1).
///
/// Returns `None` when the input does not look like a spec-compliant `dir` stream (common for
/// synthetic fixtures that use a simplified TLV encoding).
fn content_normalized_data_strict(
    ole: &mut OleFile,
    dir_decompressed: &[u8],
    encoding: &'static Encoding,
) -> Option<Vec<u8>> {
    let dir = parse_dir_for_hash_strict(dir_decompressed, encoding)?;

    // MS-OVBA §2.4.2.1 ContentNormalizedData
    let mut out = Vec::new();
    out.extend_from_slice(&dir.project_name);
    out.extend_from_slice(&dir.project_constants);

    for r in &dir.references {
        match r {
            ReferenceForHash::Registered => out.push(0x7B),
            ReferenceForHash::Project {
                size_of_libid_absolute,
                libid_absolute,
                size_of_libid_relative,
                libid_relative,
                major_version,
                minor_version,
            } => {
                // MS-OVBA §2.4.2.1 REFERENCEPROJECT normalization.
                let mut temp = Vec::new();
                temp.extend_from_slice(&size_of_libid_absolute.to_le_bytes());
                temp.extend_from_slice(libid_absolute);
                temp.extend_from_slice(&size_of_libid_relative.to_le_bytes());
                temp.extend_from_slice(libid_relative);
                temp.extend_from_slice(&major_version.to_le_bytes());
                temp.extend_from_slice(&minor_version.to_le_bytes());
                temp.push(0x00);

                for &b in &temp {
                    if b == 0x00 {
                        break;
                    }
                    out.push(b);
                }
            }
        }
    }

    for m in &dir.modules {
        let stream_path = format!("VBA/{}", m.stream_name);
        let module_stream = ole.read_stream_opt(&stream_path).ok().flatten()?;
        let text_offset = (m.text_offset as usize).min(module_stream.len());
        let source_container = &module_stream[text_offset..];
        let source = decompress_container(source_container).ok()?;
        out.extend_from_slice(&normalize_module_source_strict(&source));
    }

    Some(out)
}

#[derive(Debug, Clone, Copy)]
struct DirCursor<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> DirCursor<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.offset)
    }

    fn peek_u16(&self) -> Option<u16> {
        if self.remaining() < 2 {
            return None;
        }
        Some(u16::from_le_bytes([
            self.bytes[self.offset],
            self.bytes[self.offset + 1],
        ]))
    }

    fn read_u16(&mut self) -> Option<u16> {
        let v = self.peek_u16()?;
        self.offset += 2;
        Some(v)
    }

    fn read_u32(&mut self) -> Option<u32> {
        if self.remaining() < 4 {
            return None;
        }
        let v = u32::from_le_bytes([
            self.bytes[self.offset],
            self.bytes[self.offset + 1],
            self.bytes[self.offset + 2],
            self.bytes[self.offset + 3],
        ]);
        self.offset += 4;
        Some(v)
    }

    fn take(&mut self, n: usize) -> Option<&'a [u8]> {
        if self.remaining() < n {
            return None;
        }
        let out = &self.bytes[self.offset..self.offset + n];
        self.offset += n;
        Some(out)
    }

    fn skip(&mut self, n: usize) -> Option<()> {
        self.take(n).map(|_| ())
    }
}

fn parse_fixed_record(cur: &mut DirCursor<'_>, expected_id: u16, expected_size: u32) -> Option<()> {
    let id = cur.read_u16()?;
    let size = cur.read_u32()?;
    if id != expected_id || size != expected_size {
        return None;
    }
    cur.skip(size as usize)?;
    Some(())
}

fn parse_record_u16_u16(cur: &mut DirCursor<'_>, expected_id: u16, expected_size: u32) -> Option<u16> {
    let id = cur.read_u16()?;
    let size = cur.read_u32()?;
    if id != expected_id || size != expected_size {
        return None;
    }
    cur.read_u16()
}

fn parse_projectname_record(cur: &mut DirCursor<'_>) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != 0x0004 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    let name = cur.take(size)?.to_vec();
    Some(name)
}

fn parse_projectdocstring_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0005 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    // Reserved (0x0040) + SizeOfDocStringUnicode + DocStringUnicode
    let reserved = cur.read_u16()?;
    if reserved != 0x0040 {
        return None;
    }
    let size_unicode = cur.read_u32()? as usize;
    cur.skip(size_unicode)?;
    Some(())
}

fn parse_projecthelpfilepath_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0006 {
        return None;
    }
    let size1 = cur.read_u32()? as usize;
    cur.skip(size1)?;
    let reserved = cur.read_u16()?;
    if reserved != 0x003D {
        return None;
    }
    let size2 = cur.read_u32()? as usize;
    cur.skip(size2)?;
    Some(())
}

fn parse_projectversion_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0009 {
        return None;
    }
    let _reserved = cur.read_u32()?;
    let _major = cur.read_u32()?;
    let _minor = cur.read_u16()?;
    Some(())
}

fn parse_projectconstants_record(cur: &mut DirCursor<'_>) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != 0x000C {
        return None;
    }
    let size = cur.read_u32()? as usize;
    let constants = cur.take(size)?.to_vec();
    // Reserved (0x003C) + SizeOfConstantsUnicode + ConstantsUnicode
    let reserved = cur.read_u16()?;
    if reserved != 0x003C {
        return None;
    }
    let size_unicode = cur.read_u32()? as usize;
    cur.skip(size_unicode)?;
    Some(constants)
}

fn parse_referencename_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0016 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    let _reserved = cur.read_u16()?;
    let size_unicode = cur.read_u32()? as usize;
    cur.skip(size_unicode)?;
    Some(())
}

fn parse_reference_control(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x002F {
        return None;
    }
    let size_twiddled = cur.read_u32()? as usize;
    cur.skip(size_twiddled)?;

    // Optional NameRecordExtended (REFERENCENAME).
    if cur.peek_u16()? == 0x0016 {
        parse_referencename_record(cur)?;
    }

    let reserved3 = cur.read_u16()?;
    if reserved3 != 0x0030 {
        return None;
    }
    let size_extended = cur.read_u32()? as usize;
    cur.skip(size_extended)?;
    Some(())
}

fn parse_reference_for_hash(cur: &mut DirCursor<'_>) -> Option<Option<ReferenceForHash>> {
    // Optional NameRecord.
    if cur.peek_u16()? == 0x0016 {
        parse_referencename_record(cur)?;
    }

    let id = cur.peek_u16()?;
    match id {
        0x000D => {
            let _id = cur.read_u16()?;
            let size = cur.read_u32()? as usize;
            cur.skip(size)?;
            Some(Some(ReferenceForHash::Registered))
        }
        0x000E => {
            let _id = cur.read_u16()?;
            let size_total = cur.read_u32()? as usize;
            let start = cur.offset;

            let size_abs = cur.read_u32()?;
            let libid_abs = cur.take(size_abs as usize)?.to_vec();
            let size_rel = cur.read_u32()?;
            let libid_rel = cur.take(size_rel as usize)?.to_vec();
            let major = cur.read_u32()?;
            let minor = cur.read_u16()?;

            // Ensure we consumed exactly the expected number of bytes.
            let consumed = cur.offset.checked_sub(start)?;
            if consumed != size_total {
                return None;
            }

            Some(Some(ReferenceForHash::Project {
                size_of_libid_absolute: size_abs,
                libid_absolute: libid_abs,
                size_of_libid_relative: size_rel,
                libid_relative: libid_rel,
                major_version: major,
                minor_version: minor,
            }))
        }
        0x002F => {
            parse_reference_control(cur)?;
            Some(None)
        }
        0x0033 => {
            let _id = cur.read_u16()?;
            let size_libid = cur.read_u32()? as usize;
            cur.skip(size_libid)?;
            // Nested REFERENCECONTROL.
            parse_reference_control(cur)?;
            Some(None)
        }
        _ => None,
    }
}

fn parse_modulename_record(cur: &mut DirCursor<'_>, expected_id: u16) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != expected_id {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.take(size).map(|b| b.to_vec())
}

fn parse_module_stream_name(cur: &mut DirCursor<'_>, encoding: &'static Encoding) -> Option<String> {
    let id = cur.read_u16()?;
    if id != 0x001A {
        return None;
    }
    let size_name = cur.read_u32()? as usize;
    let raw_name = cur.take(size_name)?.to_vec();

    // Spec-compliant MODULESTREAMNAME includes Reserved (0x0032) + Unicode name, but many fixtures
    // omit those fields. Only parse them when the Reserved marker is present.
    let unicode_name = if cur.peek_u16() == Some(0x0032) {
        let _reserved = cur.read_u16()?;
        let size_unicode = cur.read_u32()? as usize;
        Some(cur.take(size_unicode)?.to_vec())
    } else {
        None
    };

    if let Some(unicode) = unicode_name {
        let (cow, _) = UTF_16LE.decode_without_bom_handling(&unicode);
        return Some(cow.into_owned());
    }

    Some(decode_dir_string(trim_reserved_u16(&raw_name), encoding))
}

fn parse_moduledocstring_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x001C {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    let reserved = cur.read_u16()?;
    if reserved != 0x0048 {
        return None;
    }
    let size_unicode = cur.read_u32()? as usize;
    cur.skip(size_unicode)?;
    Some(())
}

fn parse_moduleoffset_record(cur: &mut DirCursor<'_>) -> Option<u32> {
    let id = cur.read_u16()?;
    if id != 0x0031 {
        return None;
    }
    let size = cur.read_u32()?;
    if size != 0x0000_0004 {
        return None;
    }
    cur.read_u32()
}

fn parse_dir_for_hash_strict(dir_decompressed: &[u8], encoding: &'static Encoding) -> Option<DirForHash> {
    let mut cur = DirCursor::new(dir_decompressed);

    // PROJECTINFORMATION record (MS-OVBA §2.3.4.2.1)
    parse_fixed_record(&mut cur, 0x0001, 0x0000_0004)?; // PROJECTSYSKIND
    if cur.peek_u16()? == 0x004A {
        parse_fixed_record(&mut cur, 0x004A, 0x0000_0004)?; // PROJECTCOMPATVERSION (optional)
    }
    parse_fixed_record(&mut cur, 0x0002, 0x0000_0004)?; // PROJECTLCID
    parse_fixed_record(&mut cur, 0x0014, 0x0000_0004)?; // PROJECTLCIDINVOKE
    parse_record_u16_u16(&mut cur, 0x0003, 0x0000_0002)?; // PROJECTCODEPAGE
    let project_name = parse_projectname_record(&mut cur)?;
    parse_projectdocstring_record(&mut cur)?;
    parse_projecthelpfilepath_record(&mut cur)?;
    parse_fixed_record(&mut cur, 0x0007, 0x0000_0004)?; // PROJECTHELPCONTEXT
    parse_fixed_record(&mut cur, 0x0008, 0x0000_0004)?; // PROJECTLIBFLAGS
    parse_projectversion_record(&mut cur)?;
    let project_constants = if cur.peek_u16()? == 0x000C {
        parse_projectconstants_record(&mut cur)?
    } else {
        Vec::new()
    };

    // PROJECTREFERENCES record (MS-OVBA §2.3.4.2.2): variable array terminated by PROJECTMODULES (0x000F).
    let mut references = Vec::new();
    loop {
        let next = cur.peek_u16()?;
        if next == 0x000F {
            break;
        }
        if let Some(r) = parse_reference_for_hash(&mut cur)? {
            references.push(r);
        }
    }

    // PROJECTMODULES record (MS-OVBA §2.3.4.2.3)
    let modules_id = cur.read_u16()?;
    if modules_id != 0x000F {
        return None;
    }
    let modules_size = cur.read_u32()?;
    if modules_size != 0x0000_0002 {
        return None;
    }
    let module_count = cur.read_u16()? as usize;

    // PROJECTCOOKIE record (MS-OVBA §2.3.4.2.3.1)
    let cookie_id = cur.read_u16()?;
    if cookie_id != 0x0013 {
        return None;
    }
    let cookie_size = cur.read_u32()?;
    if cookie_size != 0x0000_0002 {
        return None;
    }
    cur.read_u16()?; // Cookie (ignored)

    let mut modules = Vec::new();
    for _ in 0..module_count {
        // MODULENAME
        parse_modulename_record(&mut cur, 0x0019)?;
        // Optional MODULENAMEUNICODE (0x0047)
        if cur.peek_u16()? == 0x0047 {
            parse_modulename_record(&mut cur, 0x0047)?;
        }
        // MODULESTREAMNAME
        let stream_name = parse_module_stream_name(&mut cur, encoding)?;
        // MODULEDOCSTRING
        parse_moduledocstring_record(&mut cur)?;
        // MODULEOFFSET
        let text_offset = parse_moduleoffset_record(&mut cur)?;
        // MODULEHELPCONTEXT
        parse_fixed_record(&mut cur, 0x001E, 0x0000_0004)?;
        // MODULECOOKIE
        parse_fixed_record(&mut cur, 0x002C, 0x0000_0002)?;
        // MODULETYPE (id 0x0021 or 0x0022, followed by reserved u32)
        let module_type_id = cur.read_u16()?;
        if module_type_id != 0x0021 && module_type_id != 0x0022 {
            return None;
        }
        cur.read_u32()?; // Reserved
        // Optional MODULEREADONLY (0x0025) and MODULEPRIVATE (0x0028)
        if cur.peek_u16()? == 0x0025 {
            cur.read_u16()?;
            cur.read_u32()?;
        }
        if cur.peek_u16()? == 0x0028 {
            cur.read_u16()?;
            cur.read_u32()?;
        }
        // Terminator + reserved
        let term = cur.read_u16()?;
        if term != 0x002B {
            return None;
        }
        let reserved = cur.read_u32()?;
        if reserved != 0 {
            return None;
        }

        modules.push(ModuleForHash {
            stream_name,
            text_offset,
        });
    }

    // Dir stream terminator + reserved (MS-OVBA §2.3.4.2)
    let end = cur.read_u16()?;
    if end != 0x0010 {
        return None;
    }
    let reserved = cur.read_u32()?;
    if reserved != 0 {
        return None;
    }
    if cur.offset != dir_decompressed.len() {
        return None;
    }

    Some(DirForHash {
        project_name,
        project_constants,
        references,
        modules,
    })
}

fn normalize_module_source_strict(bytes: &[u8]) -> Vec<u8> {
    // MS-OVBA §2.4.2.1: split into lines on CR and lone-LF; ignore the LF of CRLF.
    let mut lines: Vec<&[u8]> = Vec::new();
    let mut line_start = 0usize;
    let mut prev = 0u8;
    for (i, &ch) in bytes.iter().enumerate() {
        if ch == b'\r' {
            lines.push(&bytes[line_start..i]);
            line_start = i + 1;
        } else if ch == b'\n' {
            if prev != b'\r' {
                lines.push(&bytes[line_start..i]);
            }
            // Always advance past LF (whether it was a lone LF or part of CRLF).
            line_start = i + 1;
        }
        prev = ch;
    }
    lines.push(&bytes[line_start..]);

    let mut out = Vec::with_capacity(bytes.len());
    for line in lines {
        if starts_with_ascii_case_insensitive(line, b"attribute") {
            continue;
        }
        out.extend_from_slice(line);
    }
    out
}

fn starts_with_ascii_case_insensitive(haystack: &[u8], needle: &[u8]) -> bool {
    if haystack.len() < needle.len() {
        return false;
    }
    haystack[..needle.len()]
        .iter()
        .zip(needle.iter())
        .all(|(a, b)| a.to_ascii_lowercase() == b.to_ascii_lowercase())
}

/// Compute the MS-OVBA §2.4.2.3 **Content Hash** (v1) for a VBA project.
///
/// Per MS-OSHARED §4.3, the digest bytes used for VBA signature binding are **MD5 (16 bytes)** even
/// when the PKCS#7/CMS signature uses SHA-1/SHA-256 and even when the Authenticode `DigestInfo`
/// algorithm OID indicates SHA-256.
pub fn content_hash_md5(vba_project_bin: &[u8]) -> Result<[u8; 16], ParseError> {
    let normalized = content_normalized_data(vba_project_bin)?;
    Ok(Md5::digest(&normalized).into())
}

/// Compute the MS-OVBA §2.4.2.4 **Agile Content Hash** (v2) for a VBA project, if possible.
///
/// The Agile hash extends the legacy Content Hash by incorporating `FormsNormalizedData`
/// (designer/UserForm storages).
///
/// Returns `Ok(None)` when `FormsNormalizedData` cannot be computed (missing/unparseable data).
pub fn agile_content_hash_md5(vba_project_bin: &[u8]) -> Result<Option<[u8; 16]>, ParseError> {
    let content = content_normalized_data(vba_project_bin)?;
    let forms = match crate::forms_normalized_data(vba_project_bin) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };

    let mut h = Md5::new();
    h.update(&content);
    h.update(&forms);
    Ok(Some(h.finalize().into()))
}

#[derive(Debug, Clone, Default)]
struct ModuleInfoV3 {
    stream_name: String,
    text_offset: Option<usize>,
    // Bytes contributed to V3ContentNormalizedData before the module's normalized source code.
    //
    // We keep this as raw bytes (as stored in the decompressed `VBA/dir` record payloads) to match
    // the MS-OVBA transcript semantics and avoid codepage decoding concerns.
    transcript_prefix: Vec<u8>,
}

/// Build the MS-OVBA §2.4.2 V3 "V3ContentNormalizedData" byte sequence for a VBA project.
///
/// This is the transcript used by MS-OVBA "Contents Hash" version 3, commonly associated with the
/// `\x05DigitalSignatureExt` signature stream.
///
/// Compared to [`content_normalized_data`] (v1-ish), this includes additional metadata required by
/// the v3 transcript, notably:
/// - additional reference record types (e.g. control references), and
/// - module identity/metadata record payloads (module name/stream name/type) in `VBA/dir` order.
///
/// Spec reference: MS-OVBA §2.4.2 "Contents Hash" (version 3).
pub fn v3_content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let project_stream_bytes = ole.read_stream_opt("PROJECT")?;
    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let encoding = project_stream_bytes
        .as_deref()
        .and_then(crate::detect_project_codepage)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    let mut out = Vec::new();
    let mut current_module: Option<ModuleInfoV3> = None;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        if offset + 6 > dir_decompressed.len() {
            return Err(DirParseError::Truncated.into());
        }

        let id = u16::from_le_bytes([dir_decompressed[offset], dir_decompressed[offset + 1]]);
        let len = u32::from_le_bytes([
            dir_decompressed[offset + 2],
            dir_decompressed[offset + 3],
            dir_decompressed[offset + 4],
            dir_decompressed[offset + 5],
        ]) as usize;
        offset += 6;
        if offset + len > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = &dir_decompressed[offset..offset + len];
        offset += len;

        match id {
            // ---- References ----
            //
            // See MS-OVBA §2.4.2 for which reference record types are incorporated in each content
            // hash version.

            // REFERENCEREGISTERED
            0x000D => {
                out.extend_from_slice(data);
            }

            // REFERENCEPROJECT
            0x000E => {
                out.extend_from_slice(&normalize_reference_project(data)?);
            }

            // REFERENCECONTROL
            0x002F => {
                out.extend_from_slice(&normalize_reference_control(data)?);
            }

            // REFERENCEEXTENDED
            0x0030 => {
                out.extend_from_slice(data);
            }

            // REFERENCEORIGINAL
            0x0033 => {
                out.extend_from_slice(&normalize_reference_original(data)?);
            }

            // ---- Modules ----

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    append_v3_module(&mut out, &mut ole, &m)?;
                }
                let name = decode_dir_string(data, encoding);
                let mut transcript_prefix = Vec::new();
                transcript_prefix.extend_from_slice(data);
                current_module = Some(ModuleInfoV3 {
                    stream_name: name,
                    text_offset: None,
                    transcript_prefix,
                });
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    let trimmed = trim_reserved_u16(data);
                    m.stream_name = decode_dir_string(trimmed, encoding);
                    m.transcript_prefix.extend_from_slice(trimmed);
                }
            }

            // MODULETYPE
            //
            // MS-OVBA defines MODULETYPE records with an `Id` of either:
            // - 0x0021 (procedural module)
            // - 0x0022 (non-procedural module)
            //
            // For the v3 transcript (MS-OVBA §2.4.2.5), we only append the module's TypeRecord bytes
            // when `TypeRecord.Id == 0x0021`, and the bytes to append are:
            // `TypeRecord.Id (u16 LE) || TypeRecord.Reserved (u16 LE)`.
            0x0021 => {
                if let Some(m) = current_module.as_mut() {
                    m.transcript_prefix.extend_from_slice(&0x0021u16.to_le_bytes());
                    if data.len() >= 2 {
                        m.transcript_prefix.extend_from_slice(&data[..2]);
                    } else {
                        m.transcript_prefix.extend_from_slice(&0u16.to_le_bytes());
                    }
                }
            }
            0x0022 => {
                // Explicitly ignored: non-procedural module type records do not contribute to the
                // v3 transcript per MS-OVBA §2.4.2.5 pseudocode.
            }

            // MODULETEXTOFFSET (u32 LE).
            0x0031 => {
                if let Some(m) = current_module.as_mut() {
                    if data.len() >= 4 {
                        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                        m.text_offset = Some(n);
                    }
                }
            }

            _ => {}
        }
    }

    if let Some(m) = current_module.take() {
        append_v3_module(&mut out, &mut ole, &m)?;
    }

    Ok(out)
}

fn append_v3_module(
    out: &mut Vec<u8>,
    ole: &mut OleFile,
    module: &ModuleInfoV3,
) -> Result<(), ParseError> {
    out.extend_from_slice(&module.transcript_prefix);

    let stream_path = format!("VBA/{}", module.stream_name);
    let module_stream = ole
        .read_stream_opt(&stream_path)?
        .ok_or(ParseError::MissingStream("module stream"))?;

    let text_offset = module
        .text_offset
        .unwrap_or_else(|| guess_text_offset(&module_stream));
    let text_offset = text_offset.min(module_stream.len());
    let source_container = &module_stream[text_offset..];
    let source = decompress_container(source_container)?;
    out.extend_from_slice(&normalize_module_source(&source));
    Ok(())
}

/// Build the MS-OVBA §2.4.2 v3 `ProjectNormalizedData` byte sequence for a `vbaProject.bin`.
///
/// Spec reference: MS-OVBA §2.4.2 ("Contents Hash" version 3).
pub fn project_normalized_data_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut out = v3_content_normalized_data(vba_project_bin)?;
    let forms = forms_normalized_data(vba_project_bin)?;
    out.extend_from_slice(&forms);
    Ok(out)
}

/// Compute the MS-OVBA §2.4.2 "Contents Hash" v3 digest bytes for a `vbaProject.bin`.
///
/// This is the digest that the newest Office signature stream (`\x05DigitalSignatureExt`) binds
/// against.
///
/// MS-OVBA v3 uses **SHA-256** over:
/// `ProjectNormalizedData = V3ContentNormalizedData || FormsNormalizedData`.
pub fn contents_hash_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let normalized = project_normalized_data_v3(vba_project_bin)?;
    Ok(Sha256::digest(&normalized).to_vec())
}

fn normalize_reference_project(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for the fields used by the MS-OVBA normalization pseudocode.
    //
    // REFERENCEPROJECT (0x000E) contains two u32-len-prefixed strings followed by version integers.
    // The exact MS-OVBA `ContentNormalizedData` logic for this record type is subtle: it builds a
    // temporary byte buffer and then copies bytes until the first NUL byte.
    let mut cur = data;

    let libid_absolute = read_u32_len_prefixed_bytes(&mut cur)?;
    let libid_relative = read_u32_len_prefixed_bytes(&mut cur)?;
    if cur.len() < 6 {
        return Err(DirParseError::Truncated.into());
    }
    let major = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]);
    let minor = u16::from_le_bytes([cur[4], cur[5]]);

    // TempBuffer = LibidAbsolute || LibidRelative || MajorVersion(u32le) || MinorVersion(u16le)
    // Then copy bytes until NUL.
    let mut temp = Vec::new();
    temp.extend_from_slice(&libid_absolute);
    temp.extend_from_slice(&libid_relative);
    temp.extend_from_slice(&major.to_le_bytes());
    temp.extend_from_slice(&minor.to_le_bytes());

    Ok(copy_until_nul(&temp))
}

fn normalize_reference_control(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for the fields used by the MS-OVBA normalization pseudocode.
    //
    // REFERENCECONTROL (0x002F) contains a u32-len-prefixed libid (LibidTwiddled) followed by
    // reserved version integers. As with REFERENCEPROJECT, the MS-OVBA normalization builds a
    // TempBuffer then copies bytes until the first NUL byte.
    let mut cur = data;

    let libid_twiddled = read_u32_len_prefixed_bytes(&mut cur)?;
    if cur.len() < 6 {
        return Err(DirParseError::Truncated.into());
    }

    // Reserved1 (u32 LE) + Reserved2 (u16 LE)
    let reserved1 = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]);
    let reserved2 = u16::from_le_bytes([cur[4], cur[5]]);

    // TempBuffer = LibidTwiddled || Reserved1(u32le) || Reserved2(u16le)
    // Then copy bytes until NUL.
    let mut temp = Vec::new();
    temp.extend_from_slice(&libid_twiddled);
    temp.extend_from_slice(&reserved1.to_le_bytes());
    temp.extend_from_slice(&reserved2.to_le_bytes());

    Ok(copy_until_nul(&temp))
}

fn normalize_reference_original(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for MS-OVBA reference normalization.
    //
    // REFERENCEORIGINAL (0x0033) stores a u32-len-prefixed libid (LibidOriginal). The normalization
    // includes the libid bytes and stops at the first NUL byte.
    let mut cur = data;
    let libid_original = read_u32_len_prefixed_bytes(&mut cur)?;
    Ok(copy_until_nul(&libid_original))
}

fn read_u32_len_prefixed_bytes<'a>(cur: &mut &'a [u8]) -> Result<Vec<u8>, ParseError> {
    if cur.len() < 4 {
        return Err(DirParseError::Truncated.into());
    }
    let len = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]) as usize;
    *cur = &cur[4..];
    if cur.len() < len {
        return Err(DirParseError::Truncated.into());
    }
    let out = cur[..len].to_vec();
    *cur = &cur[len..];
    Ok(out)
}

fn copy_until_nul(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    for &b in bytes {
        if b == 0x00 {
            break;
        }
        out.push(b);
    }
    out
}

fn trim_reserved_u16(bytes: &[u8]) -> &[u8] {
    if bytes.len() >= 2 && bytes[bytes.len() - 2..] == [0x00, 0x00] {
        &bytes[..bytes.len() - 2]
    } else {
        bytes
    }
}

fn peek_next_record_id(bytes: &[u8], offset: usize) -> Option<u16> {
    if offset + 2 > bytes.len() {
        return None;
    }
    Some(u16::from_le_bytes([bytes[offset], bytes[offset + 1]]))
}

fn guess_text_offset(module_stream: &[u8]) -> usize {
    // CompressedContainer starts with 0x01, followed by a chunk header whose signature bits
    // must be 0b011 (MS-OVBA 2.4.1.3.5). This avoids false positives where 0x01 appears in the
    // module stream header.
    for idx in 0..module_stream.len().saturating_sub(3) {
        if module_stream[idx] != 0x01 {
            continue;
        }
        let header = u16::from_le_bytes([module_stream[idx + 1], module_stream[idx + 2]]);
        let signature_bits = (header & 0x7000) >> 12;
        if signature_bits == 0b011 {
            // Best-effort validation: module streams can contain header bytes before the compressed
            // source container, and those bytes can occasionally look like a container signature.
            // Attempt decompression; if it fails, keep scanning for a later candidate.
            if decompress_container(&module_stream[idx..]).is_ok() {
                return idx;
            }
        }
    }
    0
}

fn normalize_module_source(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes.len());

    let mut line_start = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\r' => {
                append_module_line(&bytes[line_start..i], &mut out);
                i += 1;
                if i < bytes.len() && bytes[i] == b'\n' {
                    // Ignore the LF of CRLF.
                    i += 1;
                }
                line_start = i;
            }
            b'\n' => {
                append_module_line(&bytes[line_start..i], &mut out);
                i += 1;
                line_start = i;
            }
            _ => i += 1,
        }
    }

    // If the module does not end with a newline, process the trailing line.
    if line_start < bytes.len() {
        append_module_line(&bytes[line_start..], &mut out);
    }

    out
}

fn append_module_line(line: &[u8], out: &mut Vec<u8>) {
    if is_attribute_line(line) {
        return;
    }
    out.extend_from_slice(line);
    out.extend_from_slice(b"\r\n");
}

fn is_attribute_line(line: &[u8]) -> bool {
    let keyword = b"attribute";
    if line.len() < keyword.len() {
        return false;
    }
    for (a, b) in line[..keyword.len()].iter().zip(keyword.iter()) {
        if a.to_ascii_lowercase() != *b {
            return false;
        }
    }
    if line.len() == keyword.len() {
        return true;
    }
    matches!(line[keyword.len()], b' ' | b'\t')
}
fn decode_dir_string(bytes: &[u8], encoding: &'static Encoding) -> String {
    // MS-OVBA dir strings are generally stored using the project codepage, but some records may
    // appear in UTF-16LE form. Use the same heuristic as the main `DirStream` parser so we can
    // reliably locate module streams with non-ASCII names.
    if looks_like_utf16le(bytes) {
        let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
        return cow.into_owned();
    }

    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || bytes.len() % 2 != 0 {
        return false;
    }

    // If a substantial portion of the high bytes are NUL, it's probably UTF-16LE for ASCII-range
    // characters (common for simple names).
    let high_bytes = bytes.iter().skip(1).step_by(2);
    let total = bytes.len() / 2;
    let nul_count = high_bytes.filter(|&&b| b == 0).count();
    // Use a ceiling half threshold. For very short inputs (e.g. 2 bytes), `total / 2` is 0 and
    // would incorrectly classify any 2-byte MBCS string as UTF-16LE.
    nul_count * 2 >= total
}
