use std::fs::File;
use std::io::{Read, Seek};
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

    let Some(info_hdr) = encryption_info.get(..8) else {
        eprintln!(
            "error: failed to read EncryptionInfo header: expected 8 bytes, got {}",
            encryption_info.len()
        );
        std::process::exit(1);
    };
    let major = u16::from_le_bytes([info_hdr[0], info_hdr[1]]);
    let minor = u16::from_le_bytes([info_hdr[2], info_hdr[3]]);
    let flags = u32::from_le_bytes([info_hdr[4], info_hdr[5], info_hdr[6], info_hdr[7]]);

    // MS-OFFCRYPTO identifies "Standard" encryption via `versionMinor == 2`, but real-world files
    // vary the major version across Office generations (2/3/4). Keep this diagnostic tool aligned
    // with `formula-offcrypto` so it correctly labels Standard-encrypted files.
    let kind = match (major, minor) {
        (4, 4) => "Agile",
        (major, 2) if (2..=4).contains(&major) => "Standard",
        (major, 3) if (3..=4).contains(&major) => "Extensible",
        _ => "Unknown",
    };

    let standard_header = if verbose && kind == "Standard" {
        match parse_standard_encryption_header_prefix(&encryption_info) {
            Ok(hdr) => Some(hdr),
            Err(err) => {
                eprintln!("error: {err}");
                std::process::exit(1);
            }
        }
    } else {
        None
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

    // One-line summary; keep deterministic to help with fixture validation and corpus triage.
    println!("{kind} ({major}.{minor}) flags=0x{flags:08x}{extra}");

    if !verbose {
        return;
    }

    if let Some(hdr) = standard_header {
        let f_cryptoapi = hdr.flags & 0x0000_0004 != 0;
        let f_aes = hdr.flags & 0x0000_0020 != 0;

        println!(
            "EncryptionHeader.flags=0x{:08x} fCryptoAPI={f_cryptoapi} fAES={f_aes}",
            hdr.flags
        );
        println!("EncryptionHeader.algId=0x{:08x}", hdr.alg_id);
        println!("EncryptionHeader.algIdHash=0x{:08x}", hdr.alg_id_hash);
        println!("EncryptionHeader.keySize={}", hdr.key_size);
        println!("EncryptionHeader.providerType=0x{:08x}", hdr.provider_type);
    }
}

#[derive(Debug, Clone, Copy)]
struct StandardEncryptionHeaderPrefix {
    flags: u32,
    #[allow(dead_code)]
    size_extra: u32,
    alg_id: u32,
    alg_id_hash: u32,
    /// Key size in *bits*.
    key_size: u32,
    provider_type: u32,
    #[allow(dead_code)]
    reserved1: u32,
    #[allow(dead_code)]
    reserved2: u32,
}

fn parse_standard_encryption_header_prefix(
    encryption_info: &[u8],
) -> Result<StandardEncryptionHeaderPrefix, String> {
    const ENCRYPTION_HEADER_FIXED_LEN: usize = 8 * 4;

    let payload = encryption_info.get(8..).unwrap_or(&[]);
    if payload.len() < 4 {
        return Err(format!(
            "truncated Standard EncryptionInfo stream while reading headerSize: needed 4 bytes, only {} available",
            payload.len()
        ));
    }

    let header_size = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
    let header_size_usize = usize::try_from(header_size)
        .map_err(|_| format!("invalid Standard EncryptionInfo headerSize {header_size}: too large"))?;
    if header_size_usize < ENCRYPTION_HEADER_FIXED_LEN {
        return Err(format!(
            "invalid Standard EncryptionInfo headerSize {header_size}: must be at least {ENCRYPTION_HEADER_FIXED_LEN}"
        ));
    }

    let needed = 4usize
        .checked_add(header_size_usize)
        .ok_or_else(|| format!("invalid Standard EncryptionInfo headerSize {header_size}: too large"))?;
    if payload.len() < needed {
        return Err(format!(
            "truncated Standard EncryptionInfo stream while reading EncryptionHeader: headerSize={header_size} bytes, only {} available",
            payload.len().saturating_sub(4)
        ));
    }

    let header = &payload[4..4 + ENCRYPTION_HEADER_FIXED_LEN];
    let flags = u32::from_le_bytes([header[0], header[1], header[2], header[3]]);
    let size_extra = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
    let alg_id = u32::from_le_bytes([header[8], header[9], header[10], header[11]]);
    let alg_id_hash = u32::from_le_bytes([header[12], header[13], header[14], header[15]]);
    let key_size = u32::from_le_bytes([header[16], header[17], header[18], header[19]]);
    let provider_type = u32::from_le_bytes([header[20], header[21], header[22], header[23]]);
    let reserved1 = u32::from_le_bytes([header[24], header[25], header[26], header[27]]);
    let reserved2 = u32::from_le_bytes([header[28], header[29], header[30], header[31]]);

    Ok(StandardEncryptionHeaderPrefix {
        flags,
        size_extra,
        alg_id,
        alg_id_hash,
        key_size,
        provider_type,
        reserved1,
        reserved2,
    })
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
                xml = xml.get(end + 2..)?.trim_start();
                continue;
            }
            '!' => {
                if xml.starts_with("<!--") {
                    let end = xml.find("-->")?;
                    xml = xml.get(end + 3..)?.trim_start();
                    continue;
                }
                // DOCTYPE or other declaration; skip to the next '>'.
                let end = xml.find('>')?;
                xml = xml.get(end + 1..)?.trim_start();
                continue;
            }
            '/' => {
                // Closing tag; skip.
                let end = xml.find('>')?;
                xml = xml.get(end + 1..)?.trim_start();
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

    if encryption_info.len() < 8 + 4 {
        return Err("truncated_before_header_size");
    }
    let header_size = u32::from_le_bytes(
        encryption_info[8..12]
            .try_into()
            .expect("slice length checked"),
    );
    let header_size_usize = header_size as usize;
    if header_size_usize < FIXED_HEADER_DWORDS_LEN {
        return Err("invalid_header_size");
    }

    let header_start = 12usize;
    let needed_for_fixed = header_start + FIXED_HEADER_DWORDS_LEN;
    if encryption_info.len() < needed_for_fixed {
        return Err("truncated_in_header_fixed_fields");
    }

    // Parse the fixed DWORDs from the available bytes. Even if the declared header size
    // exceeds the remaining buffer, fixed fields are still readable (and useful for triage).
    let hdr = &encryption_info[header_start..];
    let flags_raw = u32::from_le_bytes(hdr[0..4].try_into().unwrap());
    // let size_extra = u32::from_le_bytes(hdr[4..8].try_into().unwrap());
    let alg_id = u32::from_le_bytes(hdr[8..12].try_into().unwrap());
    let alg_id_hash = u32::from_le_bytes(hdr[12..16].try_into().unwrap());
    let key_size = u32::from_le_bytes(hdr[16..20].try_into().unwrap());

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
