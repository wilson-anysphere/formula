use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::PathBuf;

use formula_offcrypto::StandardEncryptionHeaderFlags;

/// OLE/CFB file signature.
///
/// See: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn main() {
    let mut args = std::env::args_os();
    let exe = args
        .next()
        .unwrap_or_else(|| std::ffi::OsString::from("ooxml-encryption-info"));

    let usage = || {
        eprintln!("usage: {} [--verbose|-v] <path>", exe.to_string_lossy());
        std::process::exit(2);
    };

    let mut verbose = false;
    let mut path: Option<PathBuf> = None;
    for arg in args {
        if arg == std::ffi::OsStr::new("--verbose") || arg == std::ffi::OsStr::new("-v") {
            verbose = true;
            continue;
        }
        if arg == std::ffi::OsStr::new("--help") || arg == std::ffi::OsStr::new("-h") {
            eprintln!("usage: {} [--verbose|-v] <path>", exe.to_string_lossy());
            std::process::exit(0);
        }

        if arg.to_string_lossy().starts_with('-') {
            eprintln!("error: unknown flag {}", arg.to_string_lossy());
            usage();
        }

        if path.is_some() {
            usage();
        }
        path = Some(PathBuf::from(arg));
    }
    let Some(path) = path else { usage() };

    let mut file = match File::open(&path) {
        Ok(f) => f,
        Err(err) => {
            eprintln!("error: failed to open {}: {err}", path.display());
            std::process::exit(1);
        }
    };

    let mut header = [0u8; 8];
    let n = match file.read(&mut header) {
        Ok(n) => n,
        Err(err) => {
            eprintln!("error: failed to read {}: {err}", path.display());
            std::process::exit(1);
        }
    };
    if n != OLE_MAGIC.len() || header != OLE_MAGIC {
        eprintln!(
            "not an OLE/CFB compound file (missing magic header): {}",
            path.display()
        );
        std::process::exit(1);
    }

    if let Err(err) = file.rewind() {
        eprintln!("error: failed to rewind {}: {err}", path.display());
        std::process::exit(1);
    }

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(err) => {
            eprintln!("error: failed to parse OLE/CFB compound file: {err}");
            std::process::exit(1);
        }
    };

    if !stream_exists(&mut ole, "EncryptionInfo") || !stream_exists(&mut ole, "EncryptedPackage") {
        eprintln!(
            "OLE/CFB file is not an OOXML encrypted container (missing EncryptionInfo/EncryptedPackage): {}",
            path.display()
        );
        std::process::exit(1);
    }

    let stream = match open_stream(&mut ole, "EncryptionInfo") {
        Ok(s) => s,
        Err(err) => {
            eprintln!("error: failed to open EncryptionInfo stream: {err}");
            std::process::exit(1);
        }
    };

    let mut encryption_info = Vec::new();
    // `EncryptionInfo` streams are small in practice (Agile XML descriptors are typically <100KiB).
    // Cap reads defensively to avoid accidentally loading multi-GB corrupted streams.
    const MAX_ENCRYPTION_INFO_BYTES: u64 = 4 * 1024 * 1024; // 4MiB
    let mut limited = stream.take(MAX_ENCRYPTION_INFO_BYTES + 1);
    if let Err(err) = limited.read_to_end(&mut encryption_info) {
        eprintln!("error: failed to read EncryptionInfo stream: {err}");
        std::process::exit(1);
    }
    if encryption_info.len() as u64 > MAX_ENCRYPTION_INFO_BYTES {
        eprintln!(
            "error: EncryptionInfo stream too large (>{} bytes); refusing to read full stream",
            MAX_ENCRYPTION_INFO_BYTES
        );
        std::process::exit(1);
    }

    let mut r = Reader::new(&encryption_info);
    let major = match r.read_u16_le("EncryptionVersionInfo.major") {
        Ok(v) => v,
        Err(err) => {
            eprintln!("error: {err}");
            std::process::exit(1);
        }
    };
    let minor = match r.read_u16_le("EncryptionVersionInfo.minor") {
        Ok(v) => v,
        Err(err) => {
            eprintln!("error: {err}");
            std::process::exit(1);
        }
    };
    let flags = match r.read_u32_le("EncryptionVersionInfo.flags") {
        Ok(v) => v,
        Err(err) => {
            eprintln!("error: {err}");
            std::process::exit(1);
        }
    };

    // MS-OFFCRYPTO identifies "Standard" encryption via `versionMinor == 2`, but real-world files
    // vary the major version across Office generations (2/3/4). Keep this diagnostic tool aligned
    // with `formula-offcrypto` so it correctly labels Standard-encrypted files.
    let kind = match (major, minor) {
        (4, 4) => "Agile",
        (major, 2) if (2..=4).contains(&major) => "Standard",
        (major, 3) if (3..=4).contains(&major) => "Extensible",
        _ => "Unknown",
    };

    let mut extra = String::new();
    if kind == "Agile" {
        if let Some(tag) = sniff_agile_xml_root_tag(&encryption_info[8..]) {
            extra.push_str(" xml_root=");
            extra.push_str(&tag);
        }
    } else if kind == "Standard" {
        match parse_standard_encryption_header_fixed_dwords(&encryption_info) {
            Ok(hdr) => {
                extra.push_str(&format!(
                    " hdr_flags=0x{:08x} fCryptoAPI={} fAES={} algId=0x{:08x} algIdHash=0x{:08x} keySize={}",
                    hdr.flags_raw,
                    bool_to_u8(hdr.flags.f_cryptoapi),
                    bool_to_u8(hdr.flags.f_aes),
                    hdr.alg_id,
                    hdr.alg_id_hash,
                    hdr.key_size
                ));
            }
            Err(_err) => {
                // Best-effort: if the Standard header can't be parsed, still print the
                // `EncryptionVersionInfo` line so callers can triage the scheme/version.
            }
        }
    }

    let stdout = std::io::stdout();
    let mut out = std::io::BufWriter::new(stdout.lock());
    macro_rules! outln {
        ($($arg:tt)*) => {{
            if let Err(err) = writeln!(&mut out, $($arg)*) {
                if err.kind() == std::io::ErrorKind::BrokenPipe {
                    // Allow piping output to tools like `head` without panicking.
                    return;
                }
                eprintln!("error: failed to write output: {err}");
                std::process::exit(1);
            }
        }};
    }

    // One-line summary; keep deterministic to help with fixture validation and corpus triage.
    outln!("{kind} ({major}.{minor}) flags=0x{flags:08x}{extra}");

    if !verbose {
        return;
    }
    if kind == "Standard" {
        match parse_standard_encryption_header_verbose(&encryption_info) {
            Ok(hdr) => {
                outln!("EncryptionInfo.headerSize={}", hdr.header_size);
                outln!(
                    "EncryptionHeader.flags=0x{:08x} fCryptoAPI={} fDocProps={} fExternal={} fAES={}",
                    hdr.flags_raw,
                    hdr.flags.f_cryptoapi,
                    hdr.flags.f_doc_props,
                    hdr.flags.f_external,
                    hdr.flags.f_aes
                );
                outln!("EncryptionHeader.algId=0x{:08x} ({})", hdr.alg_id, hdr.alg_id);
                outln!(
                    "EncryptionHeader.algIdHash=0x{:08x} ({})",
                    hdr.alg_id_hash, hdr.alg_id_hash
                );
                outln!(
                    "EncryptionHeader.keySize=0x{:08x} ({})",
                    hdr.key_size, hdr.key_size
                );
                outln!(
                    "EncryptionHeader.providerType=0x{:08x} ({})",
                    hdr.provider_type, hdr.provider_type
                );
                outln!("EncryptionHeader.CSPName=\"{}\"", hdr.csp_name);
            }
            Err(err) => {
                eprintln!("error: {err}");
                std::process::exit(1);
            }
        }
    }
}

fn stream_exists<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

fn open_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, String> {
    match ole.open_stream(name) {
        Ok(s) => Ok(s),
        Err(err1) => {
            let with_leading_slash = format!("/{name}");
            match ole.open_stream(&with_leading_slash) {
                Ok(s) => Ok(s),
                Err(err2) => Err(format!("{err1}; {err2}")),
            }
        }
    }
}

fn sniff_agile_xml_root_tag(buf: &[u8]) -> Option<String> {
    // The Agile EncryptionInfo stream payload is an XML document, typically small.
    // Read a bounded prefix to avoid accidentally decoding huge/corrupt streams.
    const MAX_XML_PREFIX_BYTES: usize = 128 * 1024; // 128KiB
    let buf = &buf[..buf.len().min(MAX_XML_PREFIX_BYTES)];
    let xml = decode_best_effort_xml(buf)?;
    xml_root_tag_name(&xml)
}

fn decode_best_effort_xml(buf: &[u8]) -> Option<String> {
    // UTF-8 BOM.
    let buf = buf.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(buf);

    // UTF-16 BOM.
    if let Some(rest) = buf.strip_prefix(&[0xFF, 0xFE]) {
        let (cow, _) = encoding_rs::UTF_16LE.decode_without_bom_handling(rest);
        return Some(cow.into_owned());
    }
    if let Some(rest) = buf.strip_prefix(&[0xFE, 0xFF]) {
        let (cow, _) = encoding_rs::UTF_16BE.decode_without_bom_handling(rest);
        return Some(cow.into_owned());
    }

    // If we see NUL bytes, best-effort guess UTF-16 endianness for ASCII-heavy XML.
    if buf.iter().any(|b| *b == 0) {
        let sample_len = buf.len().min(512);
        let sample_len = sample_len - (sample_len % 2);
        if sample_len >= 4 {
            let sample = &buf[..sample_len];
            let mut even_zero = 0usize;
            let mut odd_zero = 0usize;
            for (idx, b) in sample.iter().enumerate() {
                if *b == 0 {
                    if idx % 2 == 0 {
                        even_zero += 1;
                    } else {
                        odd_zero += 1;
                    }
                }
            }
            // For UTF-16LE ASCII, the high byte is typically 0, which lands at odd indexes.
            // For UTF-16BE ASCII, the high byte lands at even indexes.
            if odd_zero > even_zero.saturating_mul(3) {
                let (cow, _) = encoding_rs::UTF_16LE.decode_without_bom_handling(buf);
                return Some(cow.into_owned());
            } else if even_zero > odd_zero.saturating_mul(3) {
                let (cow, _) = encoding_rs::UTF_16BE.decode_without_bom_handling(buf);
                return Some(cow.into_owned());
            }
        }
    }

    std::str::from_utf8(buf).ok().map(|s| s.to_string())
}

fn xml_root_tag_name(mut xml: &str) -> Option<String> {
    xml = xml.trim_start();

    // Skip XML declaration / processing instructions / comments / doctypes until we find a start tag.
    loop {
        let start = xml.find('<')?;
        xml = &xml[start..];

        let after_lt = xml.get(1..)?;
        let mut chars = after_lt.chars();
        let first = chars.next()?;

        match first {
            '?' => {
                // Processing instruction (likely `<?xml ...?>`).
                let end = xml.find("?>")?;
                let next = end.checked_add(2)?;
                xml = xml.get(next..)?.trim_start();
                continue;
            }
            '!' => {
                if xml.starts_with("<!--") {
                    let end = xml.find("-->")?;
                    let next = end.checked_add(3)?;
                    xml = xml.get(next..)?.trim_start();
                    continue;
                }
                // DOCTYPE or other declaration; skip to the next '>'.
                let end = xml.find('>')?;
                let next = end.checked_add(1)?;
                xml = xml.get(next..)?.trim_start();
                continue;
            }
            '/' => {
                // Closing tag; skip.
                let end = xml.find('>')?;
                let next = end.checked_add(1)?;
                xml = xml.get(next..)?.trim_start();
                continue;
            }
            _ => {
                // Start tag.
                let name_start = 1;
                let mut end_idx = name_start;
                for (i, ch) in xml[name_start..].char_indices() {
                    if ch.is_whitespace() || ch == '>' || ch == '/' {
                        break;
                    }
                    end_idx = name_start + i + ch.len_utf8();
                }
                if end_idx <= name_start {
                    return None;
                }
                return Some(xml[name_start..end_idx].to_string());
            }
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct StandardHeaderFixedDwords {
    flags_raw: u32,
    flags: StandardEncryptionHeaderFlags,
    alg_id: u32,
    alg_id_hash: u32,
    key_size: u32,
}

fn parse_standard_encryption_header_fixed_dwords(
    encryption_info: &[u8],
) -> Result<StandardHeaderFixedDwords, &'static str> {
    // EncryptionInfo stream:
    //   EncryptionVersionInfo (8 bytes)
    //   headerSize (u32)
    //   EncryptionHeader (headerSize bytes)
    //
    // EncryptionHeader begins with 8 DWORDs (32 bytes) of fixed fields.
    const FIXED_HEADER_DWORDS_LEN: usize = 8 * 4;

    let header_size_bytes: [u8; 4] = encryption_info
        .get(8..12)
        .ok_or("truncated_before_header_size")?
        .try_into()
        .map_err(|_| "truncated_before_header_size")?;
    let header_size = u32::from_le_bytes(header_size_bytes);
    let header_size_usize = header_size as usize;
    if header_size_usize < FIXED_HEADER_DWORDS_LEN {
        return Err("invalid_header_size");
    }

    let header_start = 12usize;
    let needed_for_fixed = header_start
        .checked_add(FIXED_HEADER_DWORDS_LEN)
        .ok_or("truncated_in_header_fixed_fields")?;
    if encryption_info.len() < needed_for_fixed {
        return Err("truncated_in_header_fixed_fields");
    }

    // Parse the fixed DWORDs from the available bytes. Even if the declared header size
    // exceeds the remaining buffer, fixed fields are still readable (and useful for triage).
    let hdr = encryption_info
        .get(header_start..)
        .ok_or("truncated_in_header_fixed_fields")?;

    let flags_raw = u32::from_le_bytes(
        hdr.get(0..4)
            .and_then(|b| b.try_into().ok())
            .ok_or("truncated_in_header_fixed_fields")?,
    );
    // let size_extra = u32::from_le_bytes(
    //     hdr.get(4..8)
    //         .and_then(|b| b.try_into().ok())
    //         .ok_or("truncated_in_header_fixed_fields")?,
    // );
    let alg_id = u32::from_le_bytes(
        hdr.get(8..12)
            .and_then(|b| b.try_into().ok())
            .ok_or("truncated_in_header_fixed_fields")?,
    );
    let alg_id_hash = u32::from_le_bytes(
        hdr.get(12..16)
            .and_then(|b| b.try_into().ok())
            .ok_or("truncated_in_header_fixed_fields")?,
    );
    let key_size = u32::from_le_bytes(
        hdr.get(16..20)
            .and_then(|b| b.try_into().ok())
            .ok_or("truncated_in_header_fixed_fields")?,
    );

    Ok(StandardHeaderFixedDwords {
        flags_raw,
        flags: StandardEncryptionHeaderFlags::from_raw(flags_raw),
        alg_id,
        alg_id_hash,
        key_size,
    })
}

fn bool_to_u8(v: bool) -> u8 {
    if v { 1 } else { 0 }
}

#[derive(Debug, Clone)]
struct StandardEncryptionHeaderVerbose {
    header_size: u32,
    flags_raw: u32,
    flags: StandardEncryptionHeaderFlags,
    alg_id: u32,
    alg_id_hash: u32,
    key_size: u32,
    provider_type: u32,
    csp_name: String,
}

fn parse_standard_encryption_header_verbose(
    encryption_info: &[u8],
) -> Result<StandardEncryptionHeaderVerbose, String> {
    // Standard EncryptionInfo stream:
    //   EncryptionVersionInfo (8 bytes)
    //   headerSize (u32)
    //   EncryptionHeader (headerSize bytes)
    //
    // EncryptionHeader begins with 8 DWORDs (32 bytes) of fixed fields, followed by UTF-16LE CSPName.
    const ENCRYPTION_HEADER_FIXED_LEN: u32 = 8 * 4;
    const MAX_STANDARD_HEADER_SIZE: u32 = 64 * 1024;

    let mut r = Reader::new(encryption_info);
    // Skip `EncryptionVersionInfo`.
    r.read_u16_le("EncryptionVersionInfo.major")?;
    r.read_u16_le("EncryptionVersionInfo.minor")?;
    r.read_u32_le("EncryptionVersionInfo.flags")?;

    let header_size = r.read_u32_le("EncryptionInfo.headerSize")?;
    if header_size < ENCRYPTION_HEADER_FIXED_LEN {
        return Err(format!(
            "invalid Standard EncryptionHeader size {header_size}: must be at least {ENCRYPTION_HEADER_FIXED_LEN}"
        ));
    }
    if header_size > MAX_STANDARD_HEADER_SIZE {
        return Err(format!(
            "invalid Standard EncryptionHeader size {header_size}: exceeds max {MAX_STANDARD_HEADER_SIZE}"
        ));
    }

    let header_bytes = r.read_bytes(header_size as usize, "EncryptionHeader")?;
    debug_assert!(header_bytes.len() >= ENCRYPTION_HEADER_FIXED_LEN as usize);

    let flags_raw = u32::from_le_bytes(
        header_bytes
            .get(0..4)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.flags".to_string())?,
    );
    let _size_extra = u32::from_le_bytes(
        header_bytes
            .get(4..8)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.sizeExtra".to_string())?,
    );
    let alg_id = u32::from_le_bytes(
        header_bytes
            .get(8..12)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.algId".to_string())?,
    );
    let alg_id_hash = u32::from_le_bytes(
        header_bytes
            .get(12..16)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.algIdHash".to_string())?,
    );
    let key_size = u32::from_le_bytes(
        header_bytes
            .get(16..20)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.keySize".to_string())?,
    );
    let provider_type = u32::from_le_bytes(
        header_bytes
            .get(20..24)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.providerType".to_string())?,
    );
    let _reserved1 = u32::from_le_bytes(
        header_bytes
            .get(24..28)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.reserved1".to_string())?,
    );
    let _reserved2 = u32::from_le_bytes(
        header_bytes
            .get(28..32)
            .and_then(|b| b.try_into().ok())
            .ok_or_else(|| "truncated EncryptionHeader.reserved2".to_string())?,
    );

    let flags = StandardEncryptionHeaderFlags::from_raw(flags_raw);
    let csp_tail = header_bytes.get(32..).unwrap_or(&[]);
    let csp_name = decode_utf16le_nul_terminated_best_effort(csp_tail);

    Ok(StandardEncryptionHeaderVerbose {
        header_size,
        flags_raw,
        flags,
        alg_id,
        alg_id_hash,
        key_size,
        provider_type,
        csp_name,
    })
}

struct Reader<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn remaining(&self) -> usize {
        self.bytes.len().saturating_sub(self.pos)
    }

    fn read_bytes(&mut self, len: usize, context: &'static str) -> Result<&'a [u8], String> {
        let available = self.remaining();
        if available < len {
            return Err(format!(
                "truncated EncryptionInfo stream while reading {context}: needed {len} bytes, only {available} available"
            ));
        }
        let start = self.pos;
        let end = start.checked_add(len).ok_or_else(|| {
            format!(
                "truncated EncryptionInfo stream while reading {context}: needed {len} bytes, only {available} available"
            )
        })?;
        let out = self.bytes.get(start..end).ok_or_else(|| {
            format!(
                "truncated EncryptionInfo stream while reading {context}: needed {len} bytes, only {available} available"
            )
        })?;
        self.pos = end;
        Ok(out)
    }

    fn read_u16_le(&mut self, context: &'static str) -> Result<u16, String> {
        let b = self.read_bytes(2, context)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn read_u32_le(&mut self, context: &'static str) -> Result<u32, String> {
        let b = self.read_bytes(4, context)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }
}

fn decode_utf16le_nul_terminated_best_effort(bytes: &[u8]) -> String {
    // Best-effort decode:
    // - ignore trailing odd byte (UTF-16LE should be even-length)
    // - stop at the first NUL terminator
    // - strip trailing NULs if no terminator is present
    // - use lossy UTF-16 decoding for robustness on malformed inputs
    let len = bytes.len() - (bytes.len() % 2);
    let mut code_units = Vec::new();
    for chunk in bytes[..len].chunks_exact(2) {
        code_units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    if let Some(end) = code_units.iter().position(|u| *u == 0) {
        code_units.truncate(end);
    } else {
        while code_units.last() == Some(&0) {
            code_units.pop();
        }
    }
    String::from_utf16_lossy(&code_units)
}
