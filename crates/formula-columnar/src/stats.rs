#![forbid(unsafe_code)]

use crate::types::{ColumnType, Value};
use std::collections::HashSet;

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ColumnStats {
    pub column_type: ColumnType,
    pub distinct_count: u64,
    pub null_count: u64,
    pub min: Option<Value>,
    pub max: Option<Value>,
    pub sum: Option<f64>,
    pub avg_length: Option<f64>,
}

#[derive(Clone, Debug)]
pub(crate) struct HyperLogLog {
    p: u8,
    registers: Vec<u8>,
}

impl HyperLogLog {
    pub fn with_precision(p: u8) -> Self {
        debug_assert!((4..=16).contains(&p));
        Self {
            p,
            registers: vec![0u8; 1 << p],
        }
    }

    pub fn insert_hash(&mut self, hash: u64) {
        let idx = (hash >> (64 - self.p)) as usize;
        let w = hash << self.p;
        let rank = (w.leading_zeros() + 1) as u8;
        self.registers[idx] = self.registers[idx].max(rank);
    }

    pub fn estimate(&self) -> u64 {
        let m = self.registers.len() as f64;
        let alpha = match self.registers.len() {
            16 => 0.673,
            32 => 0.697,
            64 => 0.709,
            _ => 0.7213 / (1.0 + 1.079 / m),
        };

        let mut inv_sum = 0.0;
        let mut zeros = 0u32;
        for &r in &self.registers {
            inv_sum += 2f64.powi(-(r as i32));
            if r == 0 {
                zeros += 1;
            }
        }

        let raw = alpha * m * m / inv_sum;

        // Small range correction.
        if raw <= 2.5 * m && zeros > 0 {
            let z = zeros as f64;
            return (m * (m / z).ln()).round().max(0.0) as u64;
        }

        raw.round().max(0.0) as u64
    }
}

fn splitmix64(mut x: u64) -> u64 {
    // A fast 64-bit hash suitable for HLL / distinct counting.
    x = x.wrapping_add(0x9E3779B97F4A7C15);
    let mut z = x;
    z = (z ^ (z >> 30)).wrapping_mul(0xBF58476D1CE4E5B9);
    z = (z ^ (z >> 27)).wrapping_mul(0x94D049BB133111EB);
    z ^ (z >> 31)
}

#[derive(Clone, Debug)]
pub(crate) enum DistinctCounter {
    Exact(HashSet<u64>),
    Hll(HyperLogLog),
}

impl DistinctCounter {
    pub fn new() -> Self {
        Self::Exact(HashSet::new())
    }

    pub fn insert_hash(&mut self, hash: u64) {
        match self {
            Self::Exact(set) => {
                const THRESH: usize = 2048;
                if set.len() >= THRESH && !set.contains(&hash) {
                    // Spill into HLL once the set becomes "large enough" that exact counting
                    // would become a liability (memory + CPU).
                    let mut hll = HyperLogLog::with_precision(10); // 1024 registers
                    for &h in set.iter() {
                        hll.insert_hash(h);
                    }
                    hll.insert_hash(hash);
                    *self = Self::Hll(hll);
                } else {
                    set.insert(hash);
                }
            }
            Self::Hll(hll) => hll.insert_hash(hash),
        }
    }

    pub fn insert_i64(&mut self, v: i64) {
        self.insert_hash(splitmix64(v as u64));
    }

    pub fn insert_bool(&mut self, v: bool) {
        self.insert_hash(splitmix64(if v { 1 } else { 0 }));
    }

    pub fn insert_str(&mut self, s: &str) {
        // FNV-1a for stable hashing across runs (not cryptographic).
        let mut h: u64 = 0xcbf29ce484222325;
        for b in s.as_bytes() {
            h ^= *b as u64;
            h = h.wrapping_mul(0x100000001b3);
        }
        self.insert_hash(splitmix64(h));
    }

    pub fn estimate(&self) -> u64 {
        match self {
            Self::Exact(set) => set.len() as u64,
            Self::Hll(hll) => hll.estimate(),
        }
    }
}
