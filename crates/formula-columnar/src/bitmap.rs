#![forbid(unsafe_code)]

/// A compact bit vector used for validity and boolean storage.
///
/// Bits are stored little-endian within each `u64` word:
/// - bit 0 is the LSB of word 0
/// - bit 63 is the MSB of word 0
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BitVec {
    words: Vec<u64>,
    len: usize,
    ones: usize,
}

impl BitVec {
    pub fn new() -> Self {
        Self {
            words: Vec::new(),
            len: 0,
            ones: 0,
        }
    }

    pub fn with_capacity_bits(bits: usize) -> Self {
        let words = (bits + 63) / 64;
        Self {
            words: Vec::with_capacity(words),
            len: 0,
            ones: 0,
        }
    }

    pub fn len(&self) -> usize {
        self.len
    }

    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    pub fn push(&mut self, value: bool) {
        let bit = self.len % 64;
        if bit == 0 {
            self.words.push(0);
        }

        if value {
            let word = self.len / 64;
            self.words[word] |= 1u64 << bit;
            self.ones += 1;
        }

        self.len += 1;
    }

    pub fn get(&self, index: usize) -> bool {
        debug_assert!(index < self.len, "BitVec index out of bounds");
        let word = self.words[index / 64];
        let bit = index % 64;
        ((word >> bit) & 1) == 1
    }

    pub fn set(&mut self, index: usize, value: bool) {
        debug_assert!(index < self.len, "BitVec index out of bounds");
        let word_idx = index / 64;
        let bit = index % 64;
        let mask = 1u64 << bit;
        let was_set = (self.words[word_idx] & mask) != 0;

        match (was_set, value) {
            (true, false) => {
                self.words[word_idx] &= !mask;
                self.ones -= 1;
            }
            (false, true) => {
                self.words[word_idx] |= mask;
                self.ones += 1;
            }
            _ => {}
        }
    }

    pub fn count_ones(&self) -> usize {
        self.ones
    }

    pub fn all_true(&self) -> bool {
        self.ones == self.len
    }

    pub fn as_words(&self) -> &[u64] {
        &self.words
    }
}

impl Default for BitVec {
    fn default() -> Self {
        Self::new()
    }
}
