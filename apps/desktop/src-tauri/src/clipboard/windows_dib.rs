//! PNG <-> CF_DIBV5 conversion helpers.
//!
//! CF_DIBV5 is a Windows clipboard format containing a `BITMAPV5HEADER` followed by pixel data.
//! For interoperability with apps that don't understand the registered "PNG" clipboard format
//! (notably some Office apps), we write both formats.
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
const BI_BITFIELDS: u32 = 3;

// 'sRGB' as a u32 in little-endian.
const LCS_SRGB: u32 = 0x7352_4742;
const LCS_GM_IMAGES: u32 = 4;

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

    let mut rgba = Vec::with_capacity(rgba_len);

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

    for y in 0..height {
        let src_row = &rgba[y * row_bytes..(y + 1) * row_bytes];
        let dst_y = height - 1 - y;
        let dst_row = &mut bgra[dst_y * row_bytes..(dst_y + 1) * row_bytes];
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
    let mut out = Vec::with_capacity(out_capacity);

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
    let mut out = Vec::with_capacity(out_capacity);

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
            if compression != BI_RGB && compression != BI_BITFIELDS {
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
        other => return Err(format!("unsupported DIB bit depth: {other}")),
    };

    // In a BITMAPINFOHEADER with BI_BITFIELDS compression, the color masks are stored immediately
    // after the 40-byte header (3 DWORDs, i.e. 12 bytes). For BITMAPV4/V5 headers the masks live
    // inside the header itself.
    let mut pixel_offset = header_size;
    if compression == BI_BITFIELDS && header_size == BITMAPINFOHEADER_SIZE {
        pixel_offset = header_size
            .checked_add(12)
            .ok_or_else(|| "dib pixel offset overflow".to_string())?;
    }

    // For BI_BITFIELDS, treat the 4th byte as alpha only when a non-zero alpha mask is present
    // (BITMAPV4/V5 headers).
    let alpha_mask = if bit_count == 32 && compression == BI_BITFIELDS && header_size >= 56 {
        read_u32_le(dib_bytes, 52).unwrap_or(0)
    } else {
        0
    };
    let pixel_data_bytes = stride
        .checked_mul(height_usize)
        .ok_or_else(|| "dib pixel data size overflow".to_string())?;
    let needed = pixel_offset
        .checked_add(pixel_data_bytes)
        .ok_or_else(|| "dib total size overflow".to_string())?;
    if dib_bytes.len() < needed {
        return Err("dib does not contain full pixel data".to_string());
    }
    let pixels = &dib_bytes[pixel_offset..needed];

    // BI_RGB 32bpp is commonly used for BGRX (padding byte), but some producers treat the 4th byte
    // as alpha. Heuristic: if alpha varies across pixels, or is a constant value other than 0/255,
    // treat it as alpha. If alpha is constant 0 or 255, treat it as padding and force opaque.
    let has_alpha = if bit_count == 32 {
        match compression {
            BI_BITFIELDS => alpha_mask != 0,
            BI_RGB => {
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
            _ => false,
        }
    } else {
        false
    };

    let bottom_up = height > 0;

    let mut rgba = Vec::with_capacity(rgba_len);
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
                for px in row.chunks_exact(4) {
                    let b = px[0];
                    let g = px[1];
                    let r = px[2];
                    let a = if has_alpha { px[3] } else { 255 };
                    rgba.extend_from_slice(&[r, g, b, a]);
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
            _ => unreachable!("validated above"),
        }
    }

    encode_png_rgba8(width_u32, height_u32, &rgba)
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
