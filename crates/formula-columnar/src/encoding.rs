#![forbid(unsafe_code)]

use crate::bitmap::BitVec;
use crate::bitpacking::{bit_width_u32, bit_width_u64, pack_u32, pack_u64, unpack_u32, unpack_u64};
use std::sync::Arc;

#[derive(Clone, Debug, PartialEq)]
pub struct RleEncodedU64 {
    pub values: Vec<u64>,
    /// Cumulative end offsets (exclusive), same length as `values`.
    pub ends: Vec<u32>,
}

impl RleEncodedU64 {
    pub fn encode(values: &[u64]) -> Self {
        if values.is_empty() {
            return Self {
                values: Vec::new(),
                ends: Vec::new(),
            };
        }

        let mut out_values: Vec<u64> = Vec::new();
        let mut out_ends: Vec<u32> = Vec::new();

        let mut current = values[0];
        let mut run_len: u32 = 1;
        let mut pos: u32 = 0;
        for &v in values.iter().skip(1) {
            if v == current && run_len < u32::MAX {
                run_len += 1;
                continue;
            }

            pos = pos.saturating_add(run_len);
            out_values.push(current);
            out_ends.push(pos);
            current = v;
            run_len = 1;
        }

        pos = pos.saturating_add(run_len);
        out_values.push(current);
        out_ends.push(pos);

        Self {
            values: out_values,
            ends: out_ends,
        }
    }

    pub fn len(&self) -> usize {
        self.ends.last().copied().unwrap_or(0) as usize
    }

    pub fn runs(&self) -> usize {
        self.values.len()
    }

    pub fn get(&self, index: usize) -> u64 {
        let idx_u32 = index as u32;
        match self.ends.binary_search(&idx_u32.saturating_add(1)) {
            Ok(pos) | Err(pos) => self.values[pos],
        }
    }

    pub fn decode(&self) -> Vec<u64> {
        let mut out = Vec::new();
        let _ = out.try_reserve_exact(self.len());
        let mut start: u32 = 0;
        for (value, end) in self.values.iter().copied().zip(self.ends.iter().copied()) {
            let count = end.saturating_sub(start);
            out.extend(std::iter::repeat(value).take(count as usize));
            start = end;
        }
        out
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct RleEncodedU32 {
    pub values: Vec<u32>,
    pub ends: Vec<u32>,
}

impl RleEncodedU32 {
    pub fn encode(values: &[u32]) -> Self {
        if values.is_empty() {
            return Self {
                values: Vec::new(),
                ends: Vec::new(),
            };
        }

        let mut out_values: Vec<u32> = Vec::new();
        let mut out_ends: Vec<u32> = Vec::new();

        let mut current = values[0];
        let mut run_len: u32 = 1;
        let mut pos: u32 = 0;
        for &v in values.iter().skip(1) {
            if v == current && run_len < u32::MAX {
                run_len += 1;
                continue;
            }

            pos = pos.saturating_add(run_len);
            out_values.push(current);
            out_ends.push(pos);
            current = v;
            run_len = 1;
        }

        pos = pos.saturating_add(run_len);
        out_values.push(current);
        out_ends.push(pos);

        Self {
            values: out_values,
            ends: out_ends,
        }
    }

    pub fn len(&self) -> usize {
        self.ends.last().copied().unwrap_or(0) as usize
    }

    pub fn runs(&self) -> usize {
        self.values.len()
    }

    pub fn get(&self, index: usize) -> u32 {
        let idx_u32 = index as u32;
        match self.ends.binary_search(&idx_u32.saturating_add(1)) {
            Ok(pos) | Err(pos) => self.values[pos],
        }
    }

    pub fn decode(&self) -> Vec<u32> {
        let mut out = Vec::new();
        let _ = out.try_reserve_exact(self.len());
        let mut start: u32 = 0;
        for (value, end) in self.values.iter().copied().zip(self.ends.iter().copied()) {
            let count = end.saturating_sub(start);
            out.extend(std::iter::repeat(value).take(count as usize));
            start = end;
        }
        out
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum U64SequenceEncoding {
    Bitpacked { bit_width: u8, data: Vec<u8> },
    Rle(RleEncodedU64),
}

impl U64SequenceEncoding {
    pub fn encode(values: &[u64]) -> Self {
        let max = values.iter().copied().max().unwrap_or(0);
        let bit_width = bit_width_u64(max);
        let bitpacked_size = (values.len().saturating_mul(bit_width as usize) + 7) / 8;

        let rle = RleEncodedU64::encode(values);
        let rle_size = rle
            .runs()
            .saturating_mul(std::mem::size_of::<u64>() + std::mem::size_of::<u32>());

        if rle_size * 10 < bitpacked_size * 8 {
            // Use RLE when it is meaningfully smaller (<80%).
            return Self::Rle(rle);
        }

        Self::Bitpacked {
            bit_width,
            data: pack_u64(values, bit_width),
        }
    }

    pub fn get(&self, index: usize) -> u64 {
        match self {
            Self::Bitpacked { bit_width, data } => {
                crate::bitpacking::get_u64_at(data, *bit_width, index)
            }
            Self::Rle(rle) => rle.get(index),
        }
    }

    pub fn decode(&self, count: usize) -> Vec<u64> {
        match self {
            Self::Bitpacked { bit_width, data } => unpack_u64(data, *bit_width, count),
            Self::Rle(rle) => rle.decode(),
        }
    }

    pub fn compressed_size_bytes(&self) -> usize {
        match self {
            Self::Bitpacked { data, .. } => data.len(),
            Self::Rle(rle) => {
                rle.values.len() * std::mem::size_of::<u64>()
                    + rle.ends.len() * std::mem::size_of::<u32>()
            }
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum U32SequenceEncoding {
    Bitpacked { bit_width: u8, data: Vec<u8> },
    Rle(RleEncodedU32),
}

impl U32SequenceEncoding {
    pub fn encode(values: &[u32]) -> Self {
        let max = values.iter().copied().max().unwrap_or(0);
        let bit_width = bit_width_u32(max);
        let bitpacked_size = (values.len().saturating_mul(bit_width as usize) + 7) / 8;

        let rle = RleEncodedU32::encode(values);
        let rle_size = rle
            .runs()
            .saturating_mul(std::mem::size_of::<u32>() + std::mem::size_of::<u32>());

        if rle_size * 10 < bitpacked_size * 8 {
            return Self::Rle(rle);
        }

        Self::Bitpacked {
            bit_width,
            data: pack_u32(values, bit_width),
        }
    }

    pub fn get(&self, index: usize) -> u32 {
        match self {
            Self::Bitpacked { bit_width, data } => {
                crate::bitpacking::get_u64_at(data, *bit_width, index) as u32
            }
            Self::Rle(rle) => rle.get(index),
        }
    }

    pub fn decode(&self, count: usize) -> Vec<u32> {
        match self {
            Self::Bitpacked { bit_width, data } => unpack_u32(data, *bit_width, count),
            Self::Rle(rle) => rle.decode(),
        }
    }

    pub fn compressed_size_bytes(&self) -> usize {
        match self {
            Self::Bitpacked { data, .. } => data.len(),
            Self::Rle(rle) => {
                rle.values.len() * std::mem::size_of::<u32>()
                    + rle.ends.len() * std::mem::size_of::<u32>()
            }
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct ValueEncodedChunk {
    pub min: i64,
    pub len: usize,
    pub offsets: U64SequenceEncoding,
    pub validity: Option<BitVec>,
}

impl ValueEncodedChunk {
    pub fn compressed_size_bytes(&self) -> usize {
        let validity_bytes = self
            .validity
            .as_ref()
            .map(|v| v.as_words().len() * std::mem::size_of::<u64>())
            .unwrap_or(0);
        validity_bytes
            + self.offsets.compressed_size_bytes()
            + std::mem::size_of::<i64>()
            + std::mem::size_of::<usize>()
    }

    pub fn get_i64(&self, index: usize) -> Option<i64> {
        if let Some(validity) = &self.validity {
            if !validity.get(index) {
                return None;
            }
        }
        let offset = self.offsets.get(index);
        let value_i128 = self.min as i128 + offset as i128;
        Some(value_i128 as i64)
    }

    pub fn decode_i64(&self) -> Vec<i64> {
        let offsets = self.offsets.decode(self.len);
        offsets
            .into_iter()
            .map(|o| (self.min as i128 + o as i128) as i64)
            .collect()
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct DictionaryEncodedChunk {
    pub len: usize,
    pub indices: U32SequenceEncoding,
    pub validity: Option<BitVec>,
}

impl DictionaryEncodedChunk {
    pub fn compressed_size_bytes(&self) -> usize {
        let validity_bytes = self
            .validity
            .as_ref()
            .map(|v| v.as_words().len() * std::mem::size_of::<u64>())
            .unwrap_or(0);
        validity_bytes + self.indices.compressed_size_bytes() + std::mem::size_of::<usize>()
    }

    pub fn get_index(&self, index: usize) -> Option<u32> {
        if let Some(validity) = &self.validity {
            if !validity.get(index) {
                return None;
            }
        }
        Some(self.indices.get(index))
    }

    pub fn decode_indices(&self) -> Vec<u32> {
        self.indices.decode(self.len)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct BoolChunk {
    pub len: usize,
    pub data: Vec<u8>,
    pub validity: Option<BitVec>,
}

impl BoolChunk {
    pub fn compressed_size_bytes(&self) -> usize {
        let validity_bytes = self
            .validity
            .as_ref()
            .map(|v| v.as_words().len() * std::mem::size_of::<u64>())
            .unwrap_or(0);
        validity_bytes + self.data.len() + std::mem::size_of::<usize>()
    }

    pub fn get_bool(&self, index: usize) -> Option<bool> {
        if let Some(validity) = &self.validity {
            if !validity.get(index) {
                return None;
            }
        }
        let byte = self.data[index / 8];
        let bit = index % 8;
        Some(((byte >> bit) & 1) == 1)
    }

    pub fn decode_bools(&self) -> BitVec {
        let mut out = BitVec::with_capacity_bits(self.len);
        for i in 0..self.len {
            let b = self.get_bool(i).unwrap_or(false);
            out.push(b);
        }
        out
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct FloatChunk {
    pub values: Vec<f64>,
    pub validity: Option<BitVec>,
}

impl FloatChunk {
    pub fn len(&self) -> usize {
        self.values.len()
    }

    pub fn compressed_size_bytes(&self) -> usize {
        let validity_bytes = self
            .validity
            .as_ref()
            .map(|v| v.as_words().len() * std::mem::size_of::<u64>())
            .unwrap_or(0);
        validity_bytes + self.values.len() * std::mem::size_of::<f64>()
    }

    pub fn get_f64(&self, index: usize) -> Option<f64> {
        if let Some(validity) = &self.validity {
            if !validity.get(index) {
                return None;
            }
        }
        Some(self.values[index])
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum EncodedChunk {
    Int(ValueEncodedChunk),
    Dict(DictionaryEncodedChunk),
    Bool(BoolChunk),
    Float(FloatChunk),
}

impl EncodedChunk {
    pub fn len(&self) -> usize {
        match self {
            Self::Int(c) => c.len,
            Self::Dict(c) => c.len,
            Self::Bool(c) => c.len,
            Self::Float(c) => c.len(),
        }
    }

    pub fn compressed_size_bytes(&self) -> usize {
        match self {
            Self::Int(c) => c.compressed_size_bytes(),
            Self::Dict(c) => c.compressed_size_bytes(),
            Self::Bool(c) => c.compressed_size_bytes(),
            Self::Float(c) => c.compressed_size_bytes(),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum DecodedChunk {
    Int {
        values: Vec<i64>,
        validity: Option<BitVec>,
    },
    Dict {
        indices: Vec<u32>,
        validity: Option<BitVec>,
        dictionary: Arc<Vec<Arc<str>>>,
    },
    Bool {
        values: BitVec,
        validity: Option<BitVec>,
    },
    Float {
        values: Vec<f64>,
        validity: Option<BitVec>,
    },
}

impl DecodedChunk {
    pub fn get_i64(&self, index: usize) -> Option<i64> {
        match self {
            Self::Int { values, validity } => {
                if validity.as_ref().is_some_and(|v| !v.get(index)) {
                    return None;
                }
                Some(values[index])
            }
            _ => None,
        }
    }

    pub fn get_f64(&self, index: usize) -> Option<f64> {
        match self {
            Self::Float { values, validity } => {
                if validity.as_ref().is_some_and(|v| !v.get(index)) {
                    return None;
                }
                Some(values[index])
            }
            _ => None,
        }
    }

    pub fn get_bool(&self, index: usize) -> Option<bool> {
        match self {
            Self::Bool { values, validity } => {
                if validity.as_ref().is_some_and(|v| !v.get(index)) {
                    return None;
                }
                Some(values.get(index))
            }
            _ => None,
        }
    }

    pub fn get_string(&self, index: usize) -> Option<Arc<str>> {
        match self {
            Self::Dict {
                indices,
                validity,
                dictionary,
            } => {
                if validity.as_ref().is_some_and(|v| !v.get(index)) {
                    return None;
                }
                let idx = *indices.get(index)? as usize;
                dictionary.get(idx).cloned()
            }
            _ => None,
        }
    }
}
