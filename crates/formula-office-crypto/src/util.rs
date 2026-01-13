use crate::error::OfficeCryptoError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum EncryptionInfoKind {
    Agile,
    Standard,
}

#[derive(Debug, Clone)]
pub(crate) struct EncryptionInfoHeader {
    pub(crate) version_major: u16,
    pub(crate) version_minor: u16,
    pub(crate) flags: u32,
    pub(crate) header_size: u32,
    pub(crate) kind: EncryptionInfoKind,
    pub(crate) header_offset: usize,
}

pub(crate) fn parse_encryption_info_header(
    bytes: &[u8],
) -> Result<EncryptionInfoHeader, OfficeCryptoError> {
    if bytes.len() < 12 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionInfo stream too short".to_string(),
        ));
    }
    let version_major = read_u16_le(bytes, 0)?;
    let version_minor = read_u16_le(bytes, 2)?;
    let flags = read_u32_le(bytes, 4)?;
    let header_size = read_u32_le(bytes, 8)?;

    // MS-OFFCRYPTO / ECMA-376 identifies "Standard" encryption via `versionMinor == 2`, but
    // real-world files vary `versionMajor` across Office generations (2/3/4).
    //
    // "Extensible" encryption uses `versionMinor == 3` with `versionMajor` 3 or 4.
    let kind = match (version_major, version_minor) {
        (4, 4) => EncryptionInfoKind::Agile,
        (major, 2) if (2..=4).contains(&major) => EncryptionInfoKind::Standard,
        _ => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported EncryptionInfo version {version_major}.{version_minor} (flags={flags:#x})"
            )))
        }
    };

    Ok(EncryptionInfoHeader {
        version_major,
        version_minor,
        flags,
        header_size,
        kind,
        header_offset: 12,
    })
}

pub(crate) fn read_u16_le(bytes: &[u8], offset: usize) -> Result<u16, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

pub(crate) fn read_u32_le(bytes: &[u8], offset: usize) -> Result<u32, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

pub(crate) fn read_u64_le(bytes: &[u8], offset: usize) -> Result<u64, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u64::from_le_bytes([
        b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
    ]))
}

pub(crate) fn decode_utf16le_nul_terminated(bytes: &[u8]) -> Result<String, OfficeCryptoError> {
    if bytes.len() % 2 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(
            "UTF-16LE string has odd length".to_string(),
        ));
    }

    let mut code_units: Vec<u16> = Vec::new();
    for pair in bytes.chunks_exact(2) {
        let cu = u16::from_le_bytes([pair[0], pair[1]]);
        if cu == 0 {
            break;
        }
        code_units.push(cu);
    }
    String::from_utf16(&code_units).map_err(|_| {
        OfficeCryptoError::InvalidFormat("invalid UTF-16LE in EncryptionInfo".to_string())
    })
}
