//! PNG <-> CF_DIB / CF_DIBV5 conversion helpers.
//!
//! CF_DIB (BITMAPINFOHEADER) and CF_DIBV5 (BITMAPV5HEADER) are Windows clipboard formats containing a
//! bitmap header followed by pixel data.
//!
//! For interoperability with apps that don't understand the registered "PNG" clipboard format
//! (notably some Office apps), we write both DIB flavors:
//! - **CF_DIBV5**: top-down (negative height), BGRA, preserves alpha.
//! - **CF_DIB**: bottom-up (positive height), 32bpp BI_RGB, forces opaque alpha for compatibility
//!   with consumers that treat the 4th byte as padding (BGRX).
//!
//! This module is platform-neutral: it operates on bytes only and does not call Win32 APIs.

use std::io::Cursor;

use png::{BitDepth, ColorType};

// `MAX_IMAGE_BYTES` bounds PNG payload sizes that cross the IPC boundary. A highly-compressible PNG
// can still expand to a *much* larger RGBA buffer. Guard conversion helpers against allocating
// unbounded pixel buffers by enforcing a cap on decoded RGBA bytes.
//
// This is intentionally larger than `MAX_IMAGE_BYTES` because DIBs are uncompressed, but still
// bounded to keep clipboard conversions from exhausting memory.
// Allow decoded RGBA buffers up to 4x the raw PNG size cap.
const MAX_DECODED_RGBA_BYTES: usize = 4 * super::MAX_IMAGE_BYTES;

const BITMAPINFOHEADER_SIZE: usize = 40;
const BITMAPV5HEADER_SIZE: usize = 124;
const BI_RGB: u32 = 0;
const BI_RLE8: u32 = 1;
const BI_RLE4: u32 = 2;
const BI_BITFIELDS: u32 = 3;
const BI_PNG: u32 = 5;
const BI_ALPHABITFIELDS: u32 = 6;

// 'sRGB' as a u32 in little-endian.
const LCS_SRGB: u32 = 0x7352_4742;
const LCS_GM_IMAGES: u32 = 4;

#[derive(Copy, Clone, Debug)]
struct MaskInfo {
    mask: u32,
    shift: u32,
    max: u32,
}

fn mask_info(mask: u32) -> Option<MaskInfo> {
    if mask == 0 {
        return None;
    }
    let shift = mask.trailing_zeros();
    let shifted = mask >> shift;
    // Only support contiguous masks (e.g. 0x00FF0000).
    if shifted == 0 || (shifted & (shifted + 1)) != 0 {
        return None;
    }
    Some(MaskInfo {
        mask,
        shift,
        max: shifted,
    })
}

fn extract_masked_u8(value: u32, info: MaskInfo) -> u8 {
    let raw = ((value & info.mask) >> info.shift) as u64;
    let max = info.max as u64;
    debug_assert!(max > 0);
    // Scale to 0..=255.
    ((raw * 255 + max / 2) / max) as u8
}

fn read_bitfield_masks(
    dib_bytes: &[u8],
    header_size: usize,
    compression: u32,
) -> Option<(u32, u32, u32, u32)> {
    // BITMAPV2INFOHEADER (52 bytes) and later embed masks in the header.
    if header_size >= 52 {
        let r = read_u32_le(dib_bytes, 40)?;
        let g = read_u32_le(dib_bytes, 44)?;
        let b = read_u32_le(dib_bytes, 48)?;
        let a = if header_size >= 56 {
            read_u32_le(dib_bytes, 52).unwrap_or(0)
        } else {
            0
        };
        return Some((r, g, b, a));
    }

    // BITMAPINFOHEADER (40 bytes) stores masks immediately after the header for BI_BITFIELDS /
    // BI_ALPHABITFIELDS.
    if header_size == BITMAPINFOHEADER_SIZE
        && (compression == BI_BITFIELDS || compression == BI_ALPHABITFIELDS)
    {
        let r = read_u32_le(dib_bytes, 40)?;
        let g = read_u32_le(dib_bytes, 44)?;
        let b = read_u32_le(dib_bytes, 48)?;
        let a = if compression == BI_ALPHABITFIELDS {
            read_u32_le(dib_bytes, 52)?
        } else {
            0
        };
        return Some((r, g, b, a));
    }

    None
}

fn decode_rle8(data: &[u8], width: usize, height: usize) -> Result<Vec<u8>, String> {
    let len = width
        .checked_mul(height)
        .ok_or_else(|| "rle8 decoded index buffer overflows".to_string())?;
    let mut out = vec![0u8; len];

    let mut x: usize = 0;
    let mut y: usize = 0;
    let mut i: usize = 0;

    while i + 1 < data.len() && y < height {
        let count = data[i] as usize;
        let value = data[i + 1];
        i += 2;

        if count == 0 {
            match value {
                0 => {
                    // End of line.
                    x = 0;
                    y = y.saturating_add(1);
                }
                1 => {
                    // End of bitmap.
                    break;
                }
                2 => {
                    // Delta: move by dx, dy.
                    if i + 1 >= data.len() {
                        return Err("rle8 delta is truncated".to_string());
                    }
                    let dx = data[i] as usize;
                    let dy = data[i + 1] as usize;
                    i += 2;

                    x = x.saturating_add(dx);
                    y = y.saturating_add(dy);

                    if x > width {
                        x = width;
                    }
                    if y > height {
                        y = height;
                    }
                }
                n => {
                    // Absolute mode: copy the next n bytes literally.
                    let n = n as usize;
                    if x == width {
                        // Some producers omit end-of-line markers; treat this as an implicit new line.
                        x = 0;
                        y = y.saturating_add(1);
                        if y >= height {
                            break;
                        }
                    }
                    if i + n > data.len() {
                        return Err("rle8 absolute mode is truncated".to_string());
                    }

                    let remaining = width.saturating_sub(x);
                    let take = remaining.min(n);
                    for j in 0..take {
                        out[y * width + x] = data[i + j];
                        x += 1;
                    }
                    // Discard any pixels that would exceed the row width.
                    if take < n {
                        x = width;
                    }
                    i += n;

                    // Pad to word boundary.
                    if n % 2 == 1 {
                        if i < data.len() {
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
        } else {
            if x == width {
                // Some producers omit end-of-line markers; treat this as an implicit new line.
                x = 0;
                y = y.saturating_add(1);
                if y >= height {
                    break;
                }
            }

            let remaining = width.saturating_sub(x);
            let take = remaining.min(count);
            for _ in 0..take {
                out[y * width + x] = value;
                x += 1;
            }
            // Discard any pixels that would exceed the row width.
            if take < count {
                x = width;
            }
        }
    }

    Ok(out)
}

fn decode_rle4(data: &[u8], width: usize, height: usize) -> Result<Vec<u8>, String> {
    let len = width
        .checked_mul(height)
        .ok_or_else(|| "rle4 decoded index buffer overflows".to_string())?;
    let mut out = vec![0u8; len];

    let mut x: usize = 0;
    let mut y: usize = 0;
    let mut i: usize = 0;

    while i + 1 < data.len() && y < height {
        let count = data[i] as usize;
        let value = data[i + 1];
        i += 2;

        if count == 0 {
            match value {
                0 => {
                    // End of line.
                    x = 0;
                    y = y.saturating_add(1);
                }
                1 => {
                    // End of bitmap.
                    break;
                }
                2 => {
                    // Delta: move by dx, dy.
                    if i + 1 >= data.len() {
                        return Err("rle4 delta is truncated".to_string());
                    }
                    let dx = data[i] as usize;
                    let dy = data[i + 1] as usize;
                    i += 2;

                    x = x.saturating_add(dx);
                    y = y.saturating_add(dy);

                    if x > width {
                        x = width;
                    }
                    if y > height {
                        y = height;
                    }
                }
                n => {
                    // Absolute mode: n pixels follow, packed 2 per byte (high nibble first).
                    let n = n as usize;
                    if x == width {
                        // Some producers omit end-of-line markers; treat this as an implicit new line.
                        x = 0;
                        y = y.saturating_add(1);
                        if y >= height {
                            break;
                        }
                    }

                    let bytes = (n + 1) / 2;
                    if i + bytes > data.len() {
                        return Err("rle4 absolute mode is truncated".to_string());
                    }

                    let remaining = width.saturating_sub(x);
                    let take = remaining.min(n);
                    for p in 0..take {
                        let b = data[i + p / 2];
                        let idx = if p % 2 == 0 { b >> 4 } else { b & 0x0F };
                        out[y * width + x] = idx;
                        x += 1;
                    }
                    // Discard any pixels that would exceed the row width.
                    if take < n {
                        x = width;
                    }
                    i += bytes;

                    // Pad to word boundary: if the number of bytes in the absolute segment is odd.
                    if bytes % 2 == 1 {
                        if i < data.len() {
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
        } else {
            if x == width {
                // Some producers omit end-of-line markers; treat this as an implicit new line.
                x = 0;
                y = y.saturating_add(1);
                if y >= height {
                    break;
                }
            }

            let hi = value >> 4;
            let lo = value & 0x0F;

            let remaining = width.saturating_sub(x);
            let take = remaining.min(count);
            for p in 0..take {
                let idx = if p % 2 == 0 { hi } else { lo };
                out[y * width + x] = idx;
                x += 1;
            }
            // Discard any pixels that would exceed the row width.
            if take < count {
                x = width;
            }
        }
    }

    Ok(out)
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
    let b = bytes.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
    let b = bytes.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn read_i32_le(bytes: &[u8], offset: usize) -> Option<i32> {
    let b = bytes.get(offset..offset + 4)?;
    Some(i32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn push_u16_le(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn push_u32_le(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn push_i32_le(out: &mut Vec<u8>, v: i32) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn decode_png_rgba8(png_bytes: &[u8]) -> Result<(u32, u32, Vec<u8>), String> {
    let mut decoder = png::Decoder::new(Cursor::new(png_bytes));
    decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
    let mut reader = decoder
        .read_info()
        .map_err(|e| format!("png decode header failed: {e}"))?;

    let output_size = reader.output_buffer_size();
    if output_size > MAX_DECODED_RGBA_BYTES {
        return Err(format!(
            "png decoded buffer exceeds maximum size ({output_size} > {MAX_DECODED_RGBA_BYTES} bytes)"
        ));
    }

    let mut buf = vec![0u8; output_size];
    let info = reader
        .next_frame(&mut buf)
        .map_err(|e| format!("png decode frame failed: {e}"))?;

    let bytes = &buf[..info.buffer_size()];

    let width = info.width;
    let height = info.height;

    if width == 0 || height == 0 {
        return Err("png has zero width/height".to_string());
    }

    let width_usize = usize::try_from(width).map_err(|_| "png width exceeds platform limits")?;
    let height_usize = usize::try_from(height).map_err(|_| "png height exceeds platform limits")?;
    let rgba_len = width_usize
        .checked_mul(height_usize)
        .and_then(|v| v.checked_mul(4))
        .ok_or_else(|| "png dimensions overflow".to_string())?;
    if rgba_len > MAX_DECODED_RGBA_BYTES {
        return Err(format!(
            "png decoded RGBA exceeds maximum size ({rgba_len} > {MAX_DECODED_RGBA_BYTES} bytes)"
        ));
    }

    let mut rgba: Vec<u8> = Vec::new();
    let _ = rgba.try_reserve(rgba_len);

    match (info.color_type, info.bit_depth) {
        (ColorType::Rgba, BitDepth::Eight) => rgba.extend_from_slice(bytes),
        (ColorType::Rgb, BitDepth::Eight) => {
            for chunk in bytes.chunks_exact(3) {
                rgba.extend_from_slice(&[chunk[0], chunk[1], chunk[2], 255]);
            }
        }
        (ColorType::Grayscale, BitDepth::Eight) => {
            for &g in bytes {
                rgba.extend_from_slice(&[g, g, g, 255]);
            }
        }
        (ColorType::GrayscaleAlpha, BitDepth::Eight) => {
            for chunk in bytes.chunks_exact(2) {
                let g = chunk[0];
                let a = chunk[1];
                rgba.extend_from_slice(&[g, g, g, a]);
            }
        }
        (ct, bd) => {
            return Err(format!(
                "unsupported png output format: color={ct:?} depth={bd:?}"
            ))
        }
    }

    Ok((width, height, rgba))
}

fn encode_png_rgba8(width: u32, height: u32, rgba: &[u8]) -> Result<Vec<u8>, String> {
    let mut out = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut out, width, height);
        encoder.set_color(ColorType::Rgba);
        encoder.set_depth(BitDepth::Eight);
        let mut writer = encoder
            .write_header()
            .map_err(|e| format!("png encode header failed: {e}"))?;
        writer
            .write_image_data(rgba)
            .map_err(|e| format!("png encode data failed: {e}"))?;
    }
    Ok(out)
}

fn rgba_to_bgra_bottom_up_opaque(
    width: usize,
    height: usize,
    rgba: &[u8],
) -> Result<Vec<u8>, String> {
    let row_bytes = width
        .checked_mul(4)
        .ok_or_else(|| "image width exceeds supported limits".to_string())?;
    let expected_len = row_bytes
        .checked_mul(height)
        .ok_or_else(|| "image dimensions exceed supported limits".to_string())?;
    if rgba.len() != expected_len {
        return Err("png pixel buffer length does not match dimensions".to_string());
    }

    let mut bgra = vec![0u8; rgba.len()];

    let rows = rgba.chunks_exact(row_bytes);
    if !rows.remainder().is_empty() {
        return Err("png pixel buffer length does not match dimensions".to_string());
    }

    for (y, src_row) in rows.enumerate() {
        let dst_y = height - 1 - y;
        let dst_start = dst_y
            .checked_mul(row_bytes)
            .ok_or_else(|| "image row offset overflows".to_string())?;
        let dst_end = dst_start
            .checked_add(row_bytes)
            .ok_or_else(|| "image row offset overflows".to_string())?;
        let dst_row = bgra
            .get_mut(dst_start..dst_end)
            .ok_or_else(|| "image row out of range".to_string())?;
        for (src_px, dst_px) in src_row
            .chunks_exact(4)
            .zip(dst_row.chunks_exact_mut(4))
        {
            let r = src_px[0];
            let g = src_px[1];
            let b = src_px[2];
            // Opaque alpha for maximum compatibility with consumers that treat BI_RGB 32bpp as
            // BGRX (unused 4th byte) or as BGRA alpha.
            dst_px.copy_from_slice(&[b, g, r, 255]);
        }
    }

    Ok(bgra)
}

fn bgra_top_down_to_dibv5(width_i32: i32, height_i32: i32, bgra: &[u8]) -> Result<Vec<u8>, String> {
    let size_image =
        u32::try_from(bgra.len()).map_err(|_| "DIB pixel buffer exceeds u32 limits".to_string())?;
    let out_capacity = BITMAPV5HEADER_SIZE
        .checked_add(bgra.len())
        .ok_or_else(|| "dib size overflow".to_string())?;
    let mut out: Vec<u8> = Vec::new();
    let _ = out.try_reserve(out_capacity);

    // BITMAPV5HEADER
    push_u32_le(&mut out, BITMAPV5HEADER_SIZE as u32); // bV5Size
    push_i32_le(&mut out, width_i32); // bV5Width
    push_i32_le(&mut out, -height_i32); // bV5Height (negative => top-down)
    push_u16_le(&mut out, 1); // bV5Planes
    push_u16_le(&mut out, 32); // bV5BitCount
    push_u32_le(&mut out, BI_BITFIELDS); // bV5Compression
    push_u32_le(&mut out, size_image); // bV5SizeImage
    push_i32_le(&mut out, 0); // bV5XPelsPerMeter
    push_i32_le(&mut out, 0); // bV5YPelsPerMeter
    push_u32_le(&mut out, 0); // bV5ClrUsed
    push_u32_le(&mut out, 0); // bV5ClrImportant

    // Color masks for BGRA.
    push_u32_le(&mut out, 0x00FF_0000); // bV5RedMask
    push_u32_le(&mut out, 0x0000_FF00); // bV5GreenMask
    push_u32_le(&mut out, 0x0000_00FF); // bV5BlueMask
    push_u32_le(&mut out, 0xFF00_0000); // bV5AlphaMask

    push_u32_le(&mut out, LCS_SRGB); // bV5CSType

    // bV5Endpoints (CIEXYZTRIPLE) - leave zeroed.
    out.extend_from_slice(&[0u8; 36]);

    // Gamma values.
    push_u32_le(&mut out, 0); // bV5GammaRed
    push_u32_le(&mut out, 0); // bV5GammaGreen
    push_u32_le(&mut out, 0); // bV5GammaBlue

    push_u32_le(&mut out, LCS_GM_IMAGES); // bV5Intent
    push_u32_le(&mut out, 0); // bV5ProfileData
    push_u32_le(&mut out, 0); // bV5ProfileSize
    push_u32_le(&mut out, 0); // bV5Reserved

    debug_assert_eq!(out.len(), BITMAPV5HEADER_SIZE);

    out.extend_from_slice(bgra);
    Ok(out)
}

fn bgra_bottom_up_to_dib(width_i32: i32, height_i32: i32, bgra: &[u8]) -> Result<Vec<u8>, String> {
    let size_image =
        u32::try_from(bgra.len()).map_err(|_| "DIB pixel buffer exceeds u32 limits".to_string())?;
    let out_capacity = BITMAPINFOHEADER_SIZE
        .checked_add(bgra.len())
        .ok_or_else(|| "dib size overflow".to_string())?;
    let mut out: Vec<u8> = Vec::new();
    let _ = out.try_reserve(out_capacity);

    // BITMAPINFOHEADER
    push_u32_le(&mut out, BITMAPINFOHEADER_SIZE as u32); // biSize
    push_i32_le(&mut out, width_i32); // biWidth
    push_i32_le(&mut out, height_i32); // biHeight (positive => bottom-up)
    push_u16_le(&mut out, 1); // biPlanes
    push_u16_le(&mut out, 32); // biBitCount
    push_u32_le(&mut out, BI_RGB); // biCompression
    push_u32_le(&mut out, size_image); // biSizeImage
    push_i32_le(&mut out, 0); // biXPelsPerMeter
    push_i32_le(&mut out, 0); // biYPelsPerMeter
    push_u32_le(&mut out, 0); // biClrUsed
    push_u32_le(&mut out, 0); // biClrImportant

    debug_assert_eq!(out.len(), BITMAPINFOHEADER_SIZE);

    out.extend_from_slice(bgra);
    Ok(out)
}

/// Convert PNG bytes into CF_DIBV5 bytes (BITMAPV5HEADER + BGRA pixels).
pub fn png_to_dibv5(png_bytes: &[u8]) -> Result<Vec<u8>, String> {
    let (width, height, mut rgba) = decode_png_rgba8(png_bytes)?;

    let width_i32 = i32::try_from(width).map_err(|_| "png width exceeds DIB limits".to_string())?;
    let height_i32 = i32::try_from(height).map_err(|_| "png height exceeds DIB limits".to_string())?;

    // Convert RGBA -> BGRA in place.
    for px in rgba.chunks_exact_mut(4) {
        px.swap(0, 2);
    }

    bgra_top_down_to_dibv5(width_i32, height_i32, &rgba)
}

/// Convert PNG bytes into CF_DIB bytes (BITMAPINFOHEADER + BGRA pixels).
///
/// This is a more widely supported clipboard format than CF_DIBV5, but it does not reliably
/// preserve alpha in consumers. We therefore force all pixels opaque for best compatibility.
pub fn png_to_dib(png_bytes: &[u8]) -> Result<Vec<u8>, String> {
    let (width, height, rgba) = decode_png_rgba8(png_bytes)?;

    let width_i32 = i32::try_from(width).map_err(|_| "png width exceeds DIB limits".to_string())?;
    let height_i32 =
        i32::try_from(height).map_err(|_| "png height exceeds DIB limits".to_string())?;

    let width_usize =
        usize::try_from(width).map_err(|_| "png width exceeds platform limits".to_string())?;
    let height_usize =
        usize::try_from(height).map_err(|_| "png height exceeds platform limits".to_string())?;
    let bgra = rgba_to_bgra_bottom_up_opaque(width_usize, height_usize, &rgba)?;
    bgra_bottom_up_to_dib(width_i32, height_i32, &bgra)
}

/// Convert PNG bytes into both CF_DIB and CF_DIBV5 payloads.
///
/// Returned as `(dib, dibv5)`.
pub fn png_to_dib_and_dibv5(png_bytes: &[u8]) -> Result<(Vec<u8>, Vec<u8>), String> {
    let (width, height, mut rgba) = decode_png_rgba8(png_bytes)?;

    let width_i32 = i32::try_from(width).map_err(|_| "png width exceeds DIB limits".to_string())?;
    let height_i32 =
        i32::try_from(height).map_err(|_| "png height exceeds DIB limits".to_string())?;

    let width_usize =
        usize::try_from(width).map_err(|_| "png width exceeds platform limits".to_string())?;
    let height_usize =
        usize::try_from(height).map_err(|_| "png height exceeds platform limits".to_string())?;

    // DIB (BITMAPINFOHEADER): bottom-up, opaque.
    let bgra_dib = rgba_to_bgra_bottom_up_opaque(width_usize, height_usize, &rgba)?;
    let dib = bgra_bottom_up_to_dib(width_i32, height_i32, &bgra_dib)?;

    // DIBV5: top-down with alpha.
    for px in rgba.chunks_exact_mut(4) {
        px.swap(0, 2);
    }
    let dibv5 = bgra_top_down_to_dibv5(width_i32, height_i32, &rgba)?;

    Ok((dib, dibv5))
}

/// Convert CF_DIBV5 bytes (BITMAPV5HEADER + pixels) into PNG bytes.
pub fn dibv5_to_png(dib_bytes: &[u8]) -> Result<Vec<u8>, String> {
    if dib_bytes.len() < BITMAPINFOHEADER_SIZE {
        return Err("dib is too small to contain BITMAPINFOHEADER".to_string());
    }

    let header_size = read_u32_le(dib_bytes, 0).ok_or("failed to read bV5Size")? as usize;
    if header_size < BITMAPINFOHEADER_SIZE {
        return Err(format!(
            "unsupported DIB header size: {header_size} (expected >= {BITMAPINFOHEADER_SIZE})"
        ));
    }
    if header_size > dib_bytes.len() {
        return Err("dib header size exceeds buffer length".to_string());
    }

    let width = read_i32_le(dib_bytes, 4).ok_or("failed to read bV5Width")?;
    let height = read_i32_le(dib_bytes, 8).ok_or("failed to read bV5Height")?;
    let planes = read_u16_le(dib_bytes, 12).ok_or("failed to read bV5Planes")?;
    let bit_count = read_u16_le(dib_bytes, 14).ok_or("failed to read bV5BitCount")?;
    let compression = read_u32_le(dib_bytes, 16).ok_or("failed to read bV5Compression")?;

    if compression == BI_PNG {
        // BI_PNG: the DIB payload contains a full PNG stream instead of raw pixels.
        //
        // This is used in BMP files and occasionally shows up in clipboard DIB payloads produced by
        // some apps. Treat it as a pass-through and return the embedded PNG bytes.
        const PNG_SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];
        let size_image = read_u32_le(dib_bytes, 20).unwrap_or(0) as usize;
        let start = header_size;
        let end = if size_image == 0 {
            dib_bytes.len()
        } else {
            start
                .checked_add(size_image)
                .ok_or_else(|| "dib embedded PNG size overflow".to_string())?
        };
        if end > dib_bytes.len() {
            return Err("dib embedded PNG exceeds buffer length".to_string());
        }
        if end <= start {
            return Err("dib embedded PNG is empty".to_string());
        }
        let png = &dib_bytes[start..end];
        if !png.starts_with(&PNG_SIGNATURE) {
            return Err("dib embedded BI_PNG data is not a PNG".to_string());
        }
        return Ok(png.to_vec());
    }

    if width <= 0 {
        return Err(format!("unsupported DIB width: {width}"));
    }
    if height == 0 {
        return Err("unsupported DIB height: 0".to_string());
    }
    if planes != 1 {
        return Err(format!("unsupported DIB planes: {planes}"));
    }

    let width_u32 = width as u32;
    let height_abs = height.unsigned_abs();
    let height_u32 = height_abs;

    let width_usize =
        usize::try_from(width_u32).map_err(|_| "dib width exceeds platform limits".to_string())?;
    let height_usize =
        usize::try_from(height_u32).map_err(|_| "dib height exceeds platform limits".to_string())?;

    // RGBA output buffer bound.
    let rgba_len = width_usize
        .checked_mul(height_usize)
        .and_then(|v| v.checked_mul(4))
        .ok_or_else(|| "dib dimensions overflow".to_string())?;
    if rgba_len > MAX_DECODED_RGBA_BYTES {
        return Err(format!(
            "dib decoded RGBA exceeds maximum size ({rgba_len} > {MAX_DECODED_RGBA_BYTES} bytes)"
        ));
    }

    let (row_bytes, stride) = match bit_count {
        32 => {
            if compression != BI_RGB && compression != BI_BITFIELDS && compression != BI_ALPHABITFIELDS
            {
                return Err(format!(
                    "unsupported DIB compression for 32bpp: {compression}"
                ));
            }
            let row = width_usize
                .checked_mul(4)
                .ok_or_else(|| "dib row size overflow".to_string())?;
            (row, row)
        }
        24 => {
            if compression != BI_RGB {
                return Err(format!(
                    "unsupported DIB compression for 24bpp: {compression}"
                ));
            }
            let row = width_usize
                .checked_mul(3)
                .ok_or_else(|| "dib row size overflow".to_string())?;
            let stride = row
                .checked_add(3)
                .ok_or_else(|| "dib stride overflow".to_string())?
                & !3;
            (row, stride)
        }
        16 => {
            if compression != BI_RGB && compression != BI_BITFIELDS && compression != BI_ALPHABITFIELDS
            {
                return Err(format!(
                    "unsupported DIB compression for 16bpp: {compression}"
                ));
            }
            let row = width_usize
                .checked_mul(2)
                .ok_or_else(|| "dib row size overflow".to_string())?;
            let stride = row
                .checked_add(3)
                .ok_or_else(|| "dib stride overflow".to_string())?
                & !3;
            (row, stride)
        }
        8 => {
            if compression != BI_RGB && compression != BI_RLE8 {
                return Err(format!(
                    "unsupported DIB compression for 8bpp: {compression}"
                ));
            }
            let row = width_usize;
            let stride = row
                .checked_add(3)
                .ok_or_else(|| "dib stride overflow".to_string())?
                & !3;
            (row, stride)
        }
        4 => {
            if compression != BI_RGB && compression != BI_RLE4 {
                return Err(format!(
                    "unsupported DIB compression for 4bpp: {compression}"
                ));
            }
            let row = width_usize
                .checked_add(1)
                .ok_or_else(|| "dib row size overflow".to_string())?
                / 2;
            let stride = row
                .checked_add(3)
                .ok_or_else(|| "dib stride overflow".to_string())?
                & !3;
            (row, stride)
        }
        1 => {
            if compression != BI_RGB {
                return Err(format!(
                    "unsupported DIB compression for 1bpp: {compression}"
                ));
            }
            let row = width_usize
                .checked_add(7)
                .ok_or_else(|| "dib row size overflow".to_string())?
                / 8;
            let stride = row
                .checked_add(3)
                .ok_or_else(|| "dib stride overflow".to_string())?
                & !3;
            (row, stride)
        }
        other => return Err(format!("unsupported DIB bit depth: {other}")),
    };

    // In a BITMAPINFOHEADER with BI_BITFIELDS compression, the color masks are stored immediately
    // after the 40-byte header (3 DWORDs, i.e. 12 bytes). With BI_ALPHABITFIELDS there are 4 masks
    // (16 bytes). For BITMAPV4/V5 headers the masks live inside the header itself.
    let mut pixel_offset = header_size;
    if compression == BI_BITFIELDS && header_size == BITMAPINFOHEADER_SIZE {
        pixel_offset = header_size
            .checked_add(12)
            .ok_or_else(|| "dib pixel offset overflow".to_string())?;
    } else if compression == BI_ALPHABITFIELDS && header_size == BITMAPINFOHEADER_SIZE {
        pixel_offset = header_size
            .checked_add(16)
            .ok_or_else(|| "dib pixel offset overflow".to_string())?;
    }

    // Palette-based formats (1/4/8bpp) store a color table between the header and pixel data.
    // This is part of the CF_DIB payload (not the `BITMAPV5HEADER` itself).
    let mut palette: Option<Vec<[u8; 4]>> = None;
    if matches!(bit_count, 1 | 4 | 8) {
        let clr_used = read_u32_le(dib_bytes, 32).unwrap_or(0) as usize;
        let max_colors = 1usize << bit_count;
        let colors = if clr_used == 0 {
            max_colors
        } else {
            clr_used.min(max_colors)
        };
        let table_bytes = colors
            .checked_mul(4)
            .ok_or_else(|| "dib color table size overflow".to_string())?;
        let table_end = pixel_offset
            .checked_add(table_bytes)
            .ok_or_else(|| "dib color table offset overflow".to_string())?;
        if table_end > dib_bytes.len() {
            return Err("dib does not contain full color table".to_string());
        }

        let mut table: Vec<[u8; 4]> = Vec::new();
        let _ = table.try_reserve(colors);
        for i in 0..colors {
            let off = pixel_offset
                .checked_add(i.checked_mul(4).ok_or_else(|| "dib palette offset overflow".to_string())?)
                .ok_or_else(|| "dib palette offset overflow".to_string())?;
            let end = off
                .checked_add(4)
                .ok_or_else(|| "dib palette entry offset overflow".to_string())?;
            let entry = dib_bytes
                .get(off..end)
                .ok_or_else(|| "dib palette entry out of range".to_string())?;
            let b = entry[0];
            let g = entry[1];
            let r = entry[2];
            // The 4th byte is reserved; treat as padding and force opaque.
            table.push([r, g, b, 255]);
        }
        palette = Some(table);
        pixel_offset = table_end;
    }

    if compression == BI_RLE8 || compression == BI_RLE4 {
        let table = palette
            .as_deref()
            .ok_or_else(|| "dib palette is missing".to_string())?;
        let size_image = read_u32_le(dib_bytes, 20).unwrap_or(0) as usize;
        let start = pixel_offset;
        let end = if size_image == 0 {
            dib_bytes.len()
        } else {
            start
                .checked_add(size_image)
                .ok_or_else(|| "dib RLE size overflow".to_string())?
        };
        if end > dib_bytes.len() {
            return Err("dib RLE data exceeds buffer length".to_string());
        }
        if end <= start {
            return Err("dib RLE data is empty".to_string());
        }
        let data = &dib_bytes[start..end];

        let indices = match compression {
            BI_RLE8 => {
                if bit_count != 8 {
                    return Err("unsupported BI_RLE8 bit depth".to_string());
                }
                decode_rle8(data, width_usize, height_usize)?
            }
            BI_RLE4 => {
                if bit_count != 4 {
                    return Err("unsupported BI_RLE4 bit depth".to_string());
                }
                decode_rle4(data, width_usize, height_usize)?
            }
            _ => return Err("unsupported DIB RLE compression".to_string()),
        };

        let bottom_up = height > 0;
        let mut rgba: Vec<u8> = Vec::new();
        let _ = rgba.try_reserve(rgba_len);
        for y in 0..height_usize {
            let src_y = if bottom_up {
                height_usize - 1 - y
            } else {
                y
            };
            let row_start = src_y
                .checked_mul(width_usize)
                .ok_or_else(|| "dib row offset overflows".to_string())?;
            let row_end = row_start
                .checked_add(width_usize)
                .ok_or_else(|| "dib row offset overflows".to_string())?;
            let row = indices
                .get(row_start..row_end)
                .ok_or_else(|| "dib rle indices out of range".to_string())?;
            for &idx in row {
                let color = table
                    .get(idx as usize)
                    .ok_or_else(|| format!("dib palette index out of range: {idx}"))?;
                rgba.extend_from_slice(color);
            }
        }
        return encode_png_rgba8(width_u32, height_u32, &rgba);
    }

    // For BI_BITFIELDS / BI_ALPHABITFIELDS we may have explicit channel masks. When present, use
    // them for decoding instead of assuming BGRA byte ordering.
    let mut masks = if bit_count == 32 || bit_count == 16 {
        read_bitfield_masks(dib_bytes, header_size, compression)
    } else {
        None
    };

    let pixel_data_bytes = stride
        .checked_mul(height_usize)
        .ok_or_else(|| "dib pixel data size overflow".to_string())?;

    // Some producers include BI_BITFIELDS-style masks even when `biCompression` is BI_RGB.
    // Detect common mask patterns and adjust the pixel offset accordingly.
    if compression == BI_RGB
        && header_size == BITMAPINFOHEADER_SIZE
        && masks.is_none()
        && (bit_count == 32 || bit_count == 16)
    {
        let r_mask = read_u32_le(dib_bytes, 40).unwrap_or(0);
        let g_mask = read_u32_le(dib_bytes, 44).unwrap_or(0);
        let b_mask = read_u32_le(dib_bytes, 48).unwrap_or(0);

        let matches_common = match bit_count {
            32 => matches!(
                (r_mask, g_mask, b_mask),
                (0x00FF_0000, 0x0000_FF00, 0x0000_00FF) | (0x0000_00FF, 0x0000_FF00, 0x00FF_0000)
            ),
            16 => matches!(
                (r_mask, g_mask, b_mask),
                (0x0000_F800, 0x0000_07E0, 0x0000_001F) | (0x0000_7C00, 0x0000_03E0, 0x0000_001F)
            ),
            _ => false,
        };

        if matches_common {
            // 3 DWORD masks (RGB) are always present.
            let mut mask_bytes = 12usize;
            let mut a_mask = 0u32;

            // Prefer 4 DWORD masks (RGBA) when present for 32bpp.
            if bit_count == 32 {
                let am = read_u32_le(dib_bytes, 52).unwrap_or(0);
                if am == 0xFF00_0000 {
                    mask_bytes = 16;
                    a_mask = am;
                }
            }

            let offset_with_masks = BITMAPINFOHEADER_SIZE
                .checked_add(mask_bytes)
                .ok_or_else(|| "dib mask offset overflow".to_string())?;
            let needed_with_masks = offset_with_masks
                .checked_add(pixel_data_bytes)
                .ok_or_else(|| "dib total size overflow".to_string())?;

            if needed_with_masks <= dib_bytes.len() {
                masks = Some((r_mask, g_mask, b_mask, a_mask));
                pixel_offset = offset_with_masks;
            }
        }
    }

    // BI_RGB 16bpp defaults to 5-5-5 (no alpha).
    if bit_count == 16 && compression == BI_RGB && masks.is_none() {
        masks = Some((0x0000_7C00, 0x0000_03E0, 0x0000_001F, 0));
    }
    let needed = pixel_offset
        .checked_add(pixel_data_bytes)
        .ok_or_else(|| "dib total size overflow".to_string())?;
    if dib_bytes.len() < needed {
        return Err("dib does not contain full pixel data".to_string());
    }
    let pixels = &dib_bytes[pixel_offset..needed];

    let bitfield_decoder: Option<(MaskInfo, MaskInfo, MaskInfo, Option<MaskInfo>, Option<MaskInfo>)> =
        masks.and_then(|(r_mask, g_mask, b_mask, a_mask)| {
            if r_mask == 0 || g_mask == 0 || b_mask == 0 {
                return None;
            }
            // Ensure RGB masks don't overlap.
            if (r_mask & g_mask) != 0 || (r_mask & b_mask) != 0 || (g_mask & b_mask) != 0 {
                return None;
            }
            let r = mask_info(r_mask)?;
            let g = mask_info(g_mask)?;
            let b = mask_info(b_mask)?;
            let a = mask_info(a_mask);
            let candidate_a = if bit_count == 32 && a.is_none() {
                mask_info(0xFFFF_FFFFu32 & !(r_mask | g_mask | b_mask))
            } else {
                None
            };
            Some((r, g, b, a, candidate_a))
        });

    // BI_RGB 32bpp is commonly used for BGRX (padding byte), but some producers treat the 4th byte
    // as alpha. Likewise, some BI_BITFIELDS producers omit an explicit alpha mask but still store
    // alpha in otherwise-unused bits.
    //
    // Heuristic: if alpha varies across pixels, or is a constant value other than 0/max, treat it
    // as alpha. If alpha is constant 0 or max, treat it as padding and force opaque.
    let has_alpha = match bit_count {
        32 => {
            // For bitfields formats, prefer examining the masked alpha bits when possible.
            if let Some((_r, _g, _b, a, candidate_a)) = bitfield_decoder {
                if a.is_some() {
                    true
                } else if let Some(mask) = candidate_a {
                    let mut first: Option<u32> = None;
                    let mut saw_variation = false;
                    for px in pixels.chunks_exact(4) {
                        let value = u32::from_le_bytes([px[0], px[1], px[2], px[3]]);
                        let raw = (value & mask.mask) >> mask.shift;
                        match first {
                            None => first = Some(raw),
                            Some(v) if v != raw => {
                                saw_variation = true;
                                break;
                            }
                            _ => {}
                        }
                    }
                    if saw_variation {
                        true
                    } else {
                        let v = first.unwrap_or(0);
                        v != 0 && v != mask.max
                    }
                } else {
                    false
                }
            } else {
                // Heuristic for BI_RGB / fallback:
                //
                // - If alpha varies across pixels, treat it as alpha.
                // - If alpha is constant 0 or 255, treat it as padding and force opaque (common for
                //   BGRX).
                // - If alpha is some other constant, treat it as alpha.
                let mut first: Option<u8> = None;
                let mut saw_variation = false;
                for px in pixels.chunks_exact(4) {
                    let a = px[3];
                    match first {
                        None => first = Some(a),
                        Some(v) if v != a => {
                            saw_variation = true;
                            break;
                        }
                        _ => {}
                    }
                }
                if saw_variation {
                    true
                } else {
                    let v = first.unwrap_or(0);
                    v != 0 && v != 255
                }
            }
        }
        16 => bitfield_decoder.is_some_and(|(_r, _g, _b, a, _candidate)| a.is_some()),
        _ => false,
    };

    if bit_count == 16 && bitfield_decoder.is_none() {
        return Err("unsupported 16bpp DIB bitfield masks".to_string());
    }

    let bottom_up = height > 0;

    let mut rgba: Vec<u8> = Vec::new();
    let _ = rgba.try_reserve(rgba_len);
    for y in 0..height_usize {
        let src_y = if bottom_up {
            height_usize - 1 - y
        } else {
            y
        };
        let row_start = src_y
            .checked_mul(stride)
            .ok_or_else(|| "dib row offset overflows".to_string())?;
        let row_end = row_start
            .checked_add(row_bytes)
            .ok_or_else(|| "dib row offset overflows".to_string())?;
        let row = pixels
            .get(row_start..row_end)
            .ok_or_else(|| "dib does not contain full pixel data".to_string())?;
        match bit_count {
            32 => {
                if let Some((r_mask, g_mask, b_mask, a_mask, candidate_a_mask)) = bitfield_decoder {
                    let alpha_mask = if has_alpha {
                        a_mask.or(candidate_a_mask)
                    } else {
                        None
                    };
                    for px in row.chunks_exact(4) {
                        let value = u32::from_le_bytes([px[0], px[1], px[2], px[3]]);
                        let r = extract_masked_u8(value, r_mask);
                        let g = extract_masked_u8(value, g_mask);
                        let b = extract_masked_u8(value, b_mask);
                        let a = alpha_mask
                            .map(|mask| extract_masked_u8(value, mask))
                            .unwrap_or(255);
                        rgba.extend_from_slice(&[r, g, b, a]);
                    }
                } else {
                    // Fallback: treat as BGRA (or BGRX when `has_alpha` is false).
                    for px in row.chunks_exact(4) {
                        let b = px[0];
                        let g = px[1];
                        let r = px[2];
                        let a = if has_alpha { px[3] } else { 255 };
                        rgba.extend_from_slice(&[r, g, b, a]);
                    }
                }
            }
            24 => {
                for px in row.chunks_exact(3) {
                    let b = px[0];
                    let g = px[1];
                    let r = px[2];
                    rgba.extend_from_slice(&[r, g, b, 255]);
                }
            }
            16 => {
                let Some((r_mask, g_mask, b_mask, a_mask, _candidate_a_mask)) = bitfield_decoder
                else {
                    return Err("unsupported 16bpp DIB bitfield masks".to_string());
                };
                for px in row.chunks_exact(2) {
                    let value = u16::from_le_bytes([px[0], px[1]]) as u32;
                    let r = extract_masked_u8(value, r_mask);
                    let g = extract_masked_u8(value, g_mask);
                    let b = extract_masked_u8(value, b_mask);
                    let a = if has_alpha {
                        a_mask.map(|m| extract_masked_u8(value, m)).unwrap_or(255)
                    } else {
                        255
                    };
                    rgba.extend_from_slice(&[r, g, b, a]);
                }
            }
            8 => {
                let table = palette
                    .as_deref()
                    .ok_or_else(|| "dib palette is missing".to_string())?;
                for &idx in row.iter().take(width_usize) {
                    let color = table
                        .get(idx as usize)
                        .ok_or_else(|| format!("dib palette index out of range: {idx}"))?;
                    rgba.extend_from_slice(color);
                }
            }
            4 => {
                let table = palette
                    .as_deref()
                    .ok_or_else(|| "dib palette is missing".to_string())?;
                for x in 0..width_usize {
                    let b = row
                        .get(x / 2)
                        .ok_or_else(|| "dib row is too short for 4bpp".to_string())?;
                    let idx = if x % 2 == 0 { b >> 4 } else { b & 0x0F };
                    let color = table
                        .get(idx as usize)
                        .ok_or_else(|| format!("dib palette index out of range: {idx}"))?;
                    rgba.extend_from_slice(color);
                }
            }
            1 => {
                let table = palette
                    .as_deref()
                    .ok_or_else(|| "dib palette is missing".to_string())?;
                for x in 0..width_usize {
                    let b = row
                        .get(x / 8)
                        .ok_or_else(|| "dib row is too short for 1bpp".to_string())?;
                    let shift = 7 - (x % 8);
                    let idx = (b >> shift) & 1;
                    let color = table
                        .get(idx as usize)
                        .ok_or_else(|| format!("dib palette index out of range: {idx}"))?;
                    rgba.extend_from_slice(color);
                }
            }
            _ => return Err("unsupported DIB bit depth".to_string()),
        }
    }

    encode_png_rgba8(width_u32, height_u32, &rgba)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn crc32(bytes: &[u8]) -> u32 {
        // PNG uses CRC-32 (IEEE) over the chunk type + chunk data.
        let mut crc: u32 = 0xFFFF_FFFF;
        for &b in bytes {
            crc ^= b as u32;
            for _ in 0..8 {
                if crc & 1 != 0 {
                    crc = (crc >> 1) ^ 0xEDB8_8320;
                } else {
                    crc >>= 1;
                }
            }
        }
        !crc
    }

    fn push_png_chunk(out: &mut Vec<u8>, kind: &[u8; 4], data: &[u8]) {
        out.extend_from_slice(&(data.len() as u32).to_be_bytes());
        out.extend_from_slice(kind);
        out.extend_from_slice(data);

        let mut crc_bytes: Vec<u8> = Vec::new();
        let _ = crc_bytes.try_reserve(kind.len().saturating_add(data.len()));
        crc_bytes.extend_from_slice(kind);
        crc_bytes.extend_from_slice(data);
        out.extend_from_slice(&crc32(&crc_bytes).to_be_bytes());
    }

    fn build_minimal_rgba_png(width: u32, height: u32) -> Vec<u8> {
        // Minimal PNG: signature + IHDR + IDAT (zlib stream for empty payload) + IEND.
        //
        // This is *not* a valid image for non-zero dimensions because the decompressed scanline
        // data is empty. That's fine for our purposes because the conversion routines should
        // reject oversized dimensions before attempting to decode the frame.
        let mut png = Vec::new();
        png.extend_from_slice(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]);

        let mut ihdr: Vec<u8> = Vec::new();
        let _ = ihdr.try_reserve(13);
        ihdr.extend_from_slice(&width.to_be_bytes());
        ihdr.extend_from_slice(&height.to_be_bytes());
        ihdr.extend_from_slice(&[
            8, // bit depth
            6, // color type (RGBA)
            0, // compression
            0, // filter
            0, // interlace
        ]);
        push_png_chunk(&mut png, b"IHDR", &ihdr);

        // zlib-compressed empty payload: `zlib.compress(b"")`.
        let idat = [0x78, 0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x01];
        push_png_chunk(&mut png, b"IDAT", &idat);

        push_png_chunk(&mut png, b"IEND", &[]);
        png
    }

    fn encode_test_png() -> Vec<u8> {
        // 2x2 RGBA image:
        // (0,0) red opaque
        // (1,0) green 50% alpha
        // (0,1) blue opaque
        // (1,1) yellow transparent
        let pixels: [u8; 16] = [
            255, 0, 0, 255, // red
            0, 255, 0, 128, // green
            0, 0, 255, 255, // blue
            255, 255, 0, 0, // yellow
        ];
        encode_png_rgba8(2, 2, &pixels).expect("png encode failed")
    }

    fn decode_png(png_bytes: &[u8]) -> (u32, u32, Vec<u8>) {
        let mut decoder = png::Decoder::new(Cursor::new(png_bytes));
        decoder.set_transformations(png::Transformations::EXPAND | png::Transformations::STRIP_16);
        let mut reader = decoder.read_info().expect("png read_info");
        let mut buf = vec![0u8; reader.output_buffer_size()];
        let info = reader.next_frame(&mut buf).expect("png next_frame");
        let bytes = &buf[..info.buffer_size()];
        let mut rgba = Vec::new();
        match (info.color_type, info.bit_depth) {
            (ColorType::Rgba, BitDepth::Eight) => rgba.extend_from_slice(bytes),
            (ColorType::Rgb, BitDepth::Eight) => {
                for chunk in bytes.chunks_exact(3) {
                    rgba.extend_from_slice(&[chunk[0], chunk[1], chunk[2], 255]);
                }
            }
            (ct, bd) => panic!("unexpected decoded png format: {ct:?} {bd:?}"),
        }
        (info.width, info.height, rgba)
    }

    #[test]
    fn png_dibv5_png_roundtrip_preserves_pixels() {
        let png = encode_test_png();
        let dib = png_to_dibv5(&png).expect("png_to_dibv5 failed");

        // Header sanity checks.
        assert_eq!(read_u32_le(&dib, 0), Some(BITMAPV5HEADER_SIZE as u32));
        assert_eq!(read_i32_le(&dib, 4), Some(2));
        assert_eq!(read_i32_le(&dib, 8), Some(-2), "DIB should be top-down (negative height)");
        assert_eq!(read_u16_le(&dib, 12), Some(1));
        assert_eq!(read_u16_le(&dib, 14), Some(32));
        assert_eq!(read_u32_le(&dib, 16), Some(BI_BITFIELDS));
        assert_eq!(read_u32_le(&dib, 40), Some(0x00FF_0000));
        assert_eq!(read_u32_le(&dib, 44), Some(0x0000_FF00));
        assert_eq!(read_u32_le(&dib, 48), Some(0x0000_00FF));
        assert_eq!(read_u32_le(&dib, 52), Some(0xFF00_0000));

        let png2 = dibv5_to_png(&dib).expect("dibv5_to_png failed");

        let (w1, h1, px1) = decode_png(&png);
        let (w2, h2, px2) = decode_png(&png2);

        assert_eq!((w1, h1), (w2, h2));
        assert_eq!(px1, px2);
    }

    #[test]
    fn png_dib_png_roundtrip_forces_opaque() {
        let png = encode_test_png();
        let dib = png_to_dib(&png).expect("png_to_dib failed");

        // Header sanity checks.
        assert_eq!(read_u32_le(&dib, 0), Some(BITMAPINFOHEADER_SIZE as u32));
        assert_eq!(read_i32_le(&dib, 4), Some(2));
        assert_eq!(read_i32_le(&dib, 8), Some(2), "DIB should be bottom-up (positive height)");
        assert_eq!(read_u16_le(&dib, 12), Some(1));
        assert_eq!(read_u16_le(&dib, 14), Some(32));
        assert_eq!(read_u32_le(&dib, 16), Some(BI_RGB));

        let png2 = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png2);

        // Alpha should be forced to opaque.
        assert_eq!(
            px,
            vec![
                255, 0, 0, 255,     // red
                0, 255, 0, 255,     // green (was 128)
                0, 0, 255, 255,     // blue
                255, 255, 0, 255,   // yellow (was 0)
            ]
        );
    }

    #[test]
    fn png_to_dib_and_dibv5_produces_expected_headers() {
        let png = encode_test_png();
        let (dib, dibv5) = png_to_dib_and_dibv5(&png).expect("png_to_dib_and_dibv5 failed");

        assert_eq!(read_u32_le(&dib, 0), Some(BITMAPINFOHEADER_SIZE as u32));
        assert_eq!(read_i32_le(&dib, 4), Some(2));
        assert_eq!(read_i32_le(&dib, 8), Some(2));
        assert_eq!(read_u16_le(&dib, 14), Some(32));

        assert_eq!(read_u32_le(&dibv5, 0), Some(BITMAPV5HEADER_SIZE as u32));
        assert_eq!(read_i32_le(&dibv5, 4), Some(2));
        assert_eq!(read_i32_le(&dibv5, 8), Some(-2));
        assert_eq!(read_u16_le(&dibv5, 14), Some(32));
        assert_eq!(read_u32_le(&dibv5, 16), Some(BI_BITFIELDS));
    }

    #[test]
    fn dib32_bi_rgb_is_treated_as_opaque() {
        // Minimal BITMAPINFOHEADER (40 bytes) + 2 pixels BGRA (bottom-up).
        //
        // Many BI_RGB 32bpp DIBs use BGRX where the 4th byte is padding (often 0). If we treat it
        // as alpha we end up with a fully transparent image.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Pixel data starts immediately after the 40-byte header.
        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);
        // Two pixels: red and green. Alpha/padding byte is 0 to simulate BGRX.
        dib.extend_from_slice(&[
            0, 0, 255, 0, // red (BGRA/BGRX)
            0, 255, 0, 0, // green
        ]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 255, 0, 255, 0, 255]);
    }

    #[test]
    fn dib32_bi_rgb_preserves_alpha_when_nontrivial() {
        // Same as `dib32_bi_rgb_is_treated_as_opaque`, but include a non-trivial alpha value so our
        // heuristic treats BI_RGB 32bpp as BGRA.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);
        // Two pixels: red at 50% alpha, green opaque.
        dib.extend_from_slice(&[
            0, 0, 255, 128, // red (BGRA)
            0, 255, 0, 255, // green
        ]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 128, 0, 255, 0, 255]);
    }

    #[test]
    fn dib32_bi_alphabitfields_preserves_alpha() {
        // Some producers use BI_ALPHABITFIELDS with a BITMAPINFOHEADER and explicit color masks.
        // Treat this like BI_BITFIELDS with alpha.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_ALPHABITFIELDS); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Color masks (BGRA).
        push_u32_le(&mut dib, 0x00FF_0000); // red
        push_u32_le(&mut dib, 0x0000_FF00); // green
        push_u32_le(&mut dib, 0x0000_00FF); // blue
        push_u32_le(&mut dib, 0xFF00_0000); // alpha

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE + 16);

        // One BGRA pixel: red at 50% alpha.
        dib.extend_from_slice(&[0, 0, 255, 128]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 128]);
    }

    #[test]
    fn dib32_bi_bitfields_without_alpha_mask_preserves_nontrivial_alpha() {
        // BITMAPINFOHEADER (40 bytes) + 3 color masks (no alpha mask) + 2 pixels (bottom-up).
        //
        // Some producers store alpha in the 4th byte even without an explicit alpha mask. Use a
        // heuristic to preserve it when it looks non-trivial.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_BITFIELDS); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Color masks (BGR).
        push_u32_le(&mut dib, 0x00FF_0000); // red
        push_u32_le(&mut dib, 0x0000_FF00); // green
        push_u32_le(&mut dib, 0x0000_00FF); // blue

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE + 12);

        // Two pixels BGRA: red at 50% alpha, green opaque.
        dib.extend_from_slice(&[
            0, 0, 255, 128, // red
            0, 255, 0, 255, // green
        ]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 128, 0, 255, 0, 255]);
    }

    #[test]
    fn dib32_bi_alphabitfields_respects_channel_masks() {
        // BITMAPINFOHEADER (40 bytes) + 4 masks + 1 pixel (bottom-up).
        //
        // Use non-BGRA masks to ensure we honor the masks instead of assuming BGRA byte order.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_ALPHABITFIELDS); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Masks for RGBA byte order (little endian u32 = [R, G, B, A]).
        push_u32_le(&mut dib, 0x0000_00FF); // red
        push_u32_le(&mut dib, 0x0000_FF00); // green
        push_u32_le(&mut dib, 0x00FF_0000); // blue
        push_u32_le(&mut dib, 0xFF00_0000); // alpha

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE + 16);

        // One RGBA pixel: red at 50% alpha.
        dib.extend_from_slice(&[255, 0, 0, 128]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 128]);
    }

    #[test]
    fn dib32_bi_bitfields_respects_channel_masks() {
        // BITMAPINFOHEADER (40 bytes) + 3 masks + 2 pixels (bottom-up).
        //
        // Use non-BGRA masks to ensure we honor the masks instead of assuming BGRA byte order.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_BITFIELDS); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Masks for RGBx byte order (little endian u32 = [R, G, B, X]).
        push_u32_le(&mut dib, 0x0000_00FF); // red
        push_u32_le(&mut dib, 0x0000_FF00); // green
        push_u32_le(&mut dib, 0x00FF_0000); // blue

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE + 12);

        // Two pixels in RGBX byte order with a padding byte of 0 (X=0).
        // Our alpha heuristic should treat X as padding and force opaque.
        dib.extend_from_slice(&[
            255, 0, 0, 0, // red
            0, 255, 0, 0, // green
        ]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 255, 0, 255, 0, 255]);
    }

    #[test]
    fn dib16_bi_rgb_555_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + 1 pixel (bottom-up).
        //
        // BI_RGB 16bpp defaults to 5-5-5 BGR (no masks in header).
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 16); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // One pixel: 0b0RRRRRGGGGGBBBBB with R=31, G=0, B=0 => 0x7C00 (red).
        // Rows are padded to a 4-byte boundary; include 2 padding bytes.
        dib.extend_from_slice(&[0x00, 0x7C, 0x00, 0x00]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 255]);
    }

    #[test]
    fn dib16_bi_bitfields_565_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + 3 masks + 1 pixel (bottom-up).
        //
        // 16bpp 5-6-5 BGR is commonly encoded using BI_BITFIELDS masks.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 16); // biBitCount
        push_u32_le(&mut dib, BI_BITFIELDS); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        // Masks for 5-6-5.
        push_u32_le(&mut dib, 0x0000_F800); // red
        push_u32_le(&mut dib, 0x0000_07E0); // green
        push_u32_le(&mut dib, 0x0000_001F); // blue

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE + 12);

        // One pixel: green max (63) => 0x07E0.
        // Include 2 padding bytes for 4-byte row alignment.
        dib.extend_from_slice(&[0xE0, 0x07, 0x00, 0x00]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![0, 255, 0, 255]);
    }

    #[test]
    fn dib32_bi_rgb_with_embedded_masks_decodes_correctly() {
        // Some producers include BI_BITFIELDS-style masks after a BITMAPINFOHEADER even when
        // `biCompression` is BI_RGB.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression (even though masks are present)
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // BGRA masks (DWORDs).
        push_u32_le(&mut dib, 0x00FF_0000); // red
        push_u32_le(&mut dib, 0x0000_FF00); // green
        push_u32_le(&mut dib, 0x0000_00FF); // blue
        push_u32_le(&mut dib, 0xFF00_0000); // alpha

        // One pixel: red at 50% alpha in BGRA byte order.
        dib.extend_from_slice(&[0, 0, 255, 128]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 128]);
    }

    #[test]
    fn dib16_bi_rgb_with_embedded_masks_decodes_correctly() {
        // Some producers include 16bpp BI_BITFIELDS masks after a BITMAPINFOHEADER even when
        // `biCompression` is BI_RGB.
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 1); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 16); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression (even though masks are present)
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // 5-6-5 masks (DWORDs).
        push_u32_le(&mut dib, 0x0000_F800); // red
        push_u32_le(&mut dib, 0x0000_07E0); // green
        push_u32_le(&mut dib, 0x0000_001F); // blue

        // One pixel: green max (63) => 0x07E0.
        // Include 2 padding bytes for 4-byte row alignment.
        dib.extend_from_slice(&[0xE0, 0x07, 0x00, 0x00]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![0, 255, 0, 255]);
    }

    #[test]
    fn dib8_bi_rgb_palette_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + palette + 2 pixels (bottom-up).
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 8); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 2); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // Palette (RGBQUAD): [B, G, R, 0]
        // 0 => red
        dib.extend_from_slice(&[0, 0, 255, 0]);
        // 1 => green
        dib.extend_from_slice(&[0, 255, 0, 0]);

        // Pixel indices + row padding to 4 bytes.
        dib.extend_from_slice(&[0, 1, 0, 0]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 255, 0, 255, 0, 255]);
    }

    #[test]
    fn dib4_bi_rgb_palette_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + palette + 2 pixels packed into 1 byte (bottom-up).
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 4); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 2); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // Palette: 0 => blue, 1 => yellow
        dib.extend_from_slice(&[255, 0, 0, 0]); // blue
        dib.extend_from_slice(&[0, 255, 255, 0]); // yellow

        // Two pixels: index 0 then 1 => 0x01 (high nibble first)
        // Row padded to 4 bytes.
        dib.extend_from_slice(&[0x01, 0, 0, 0]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![0, 0, 255, 255, 255, 255, 0, 255]);
    }

    #[test]
    fn dib1_bi_rgb_palette_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + palette + 2 pixels packed into 1 byte (bottom-up).
        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 1); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 2); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // Palette: 0 => black, 1 => white
        dib.extend_from_slice(&[0, 0, 0, 0]); // black
        dib.extend_from_slice(&[255, 255, 255, 0]); // white

        // Two pixels: 0 then 1 => bits 7..6 = 0b01 => 0x40.
        // Row padded to 4 bytes.
        dib.extend_from_slice(&[0x40, 0, 0, 0]);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![0, 0, 0, 255, 255, 255, 255, 255]);
    }

    #[test]
    fn dib8_bi_rle8_palette_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + palette + RLE8 data (bottom-up).
        let rle = vec![1, 0, 1, 1, 0, 0, 0, 1]; // 2 pixels + EOL + EOB

        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 8); // biBitCount
        push_u32_le(&mut dib, BI_RLE8); // biCompression
        push_u32_le(&mut dib, rle.len() as u32); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 2); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // Palette: 0 => red, 1 => green.
        dib.extend_from_slice(&[0, 0, 255, 0]);
        dib.extend_from_slice(&[0, 255, 0, 0]);

        dib.extend_from_slice(&rle);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![255, 0, 0, 255, 0, 255, 0, 255]);
    }

    #[test]
    fn dib4_bi_rle4_palette_decodes_correctly() {
        // BITMAPINFOHEADER (40 bytes) + palette + RLE4 data (bottom-up).
        let rle = vec![2, 0x01, 0, 0, 0, 1]; // 2 pixels + EOL + EOB

        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 4); // biBitCount
        push_u32_le(&mut dib, BI_RLE4); // biCompression
        push_u32_le(&mut dib, rle.len() as u32); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 2); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        // Palette: 0 => blue, 1 => yellow.
        dib.extend_from_slice(&[255, 0, 0, 0]); // blue
        dib.extend_from_slice(&[0, 255, 255, 0]); // yellow

        dib.extend_from_slice(&rle);

        let png = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        let (_w, _h, px) = decode_png(&png);
        assert_eq!(px, vec![0, 0, 255, 255, 255, 255, 0, 255]);
    }

    #[test]
    fn dib_bi_png_is_passed_through() {
        // Some producers embed a full PNG stream in a DIB header using BI_PNG compression.
        // Ensure we can extract the PNG without attempting to interpret raw pixels.
        let png = encode_test_png();

        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, 2); // biWidth
        push_i32_le(&mut dib, 2); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 0); // biBitCount (ignored for BI_PNG)
        push_u32_le(&mut dib, BI_PNG); // biCompression
        push_u32_le(&mut dib, png.len() as u32); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);
        dib.extend_from_slice(&png);

        let out = dibv5_to_png(&dib).expect("dibv5_to_png failed");
        assert_eq!(out, png);
    }

    #[test]
    fn png_to_dibv5_rejects_oversized_images_without_allocating() {
        // Construct a PNG whose decoded RGBA buffer would exceed our cap.
        let width = (MAX_DECODED_RGBA_BYTES / 4) as u32 + 1;
        let png = build_minimal_rgba_png(width, 1);
        let err = png_to_dibv5(&png).expect_err("expected oversized PNG to be rejected");
        assert!(
            err.contains("png decoded buffer exceeds maximum size"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn dibv5_to_png_rejects_oversized_dimensions() {
        let width = (MAX_DECODED_RGBA_BYTES / 4) as i32 + 1;

        let mut dib = Vec::new();
        push_u32_le(&mut dib, BITMAPINFOHEADER_SIZE as u32); // biSize
        push_i32_le(&mut dib, width); // biWidth
        push_i32_le(&mut dib, 1); // biHeight (bottom-up)
        push_u16_le(&mut dib, 1); // biPlanes
        push_u16_le(&mut dib, 32); // biBitCount
        push_u32_le(&mut dib, BI_RGB); // biCompression
        push_u32_le(&mut dib, 0); // biSizeImage
        push_i32_le(&mut dib, 0); // biXPelsPerMeter
        push_i32_le(&mut dib, 0); // biYPelsPerMeter
        push_u32_le(&mut dib, 0); // biClrUsed
        push_u32_le(&mut dib, 0); // biClrImportant

        debug_assert_eq!(dib.len(), BITMAPINFOHEADER_SIZE);

        let err = dibv5_to_png(&dib).expect_err("expected oversized DIB to be rejected");
        assert!(
            err.contains("dib decoded RGBA exceeds maximum size"),
            "unexpected error: {err}"
        );
    }
}
