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

    pub fn with_len_all_true(bits: usize) -> Self {
        if bits == 0 {
            return Self::new();
        }

        let word_len = (bits + 63) / 64;
        let mut words = vec![u64::MAX; word_len];
        let rem = bits % 64;
        if rem != 0 {
            let mask = (1u64 << rem) - 1;
            if let Some(last) = words.last_mut() {
                *last = mask;
            }
        }

        Self {
            words,
            len: bits,
            ones: bits,
        }
    }

    pub fn with_len_all_false(bits: usize) -> Self {
        if bits == 0 {
            return Self::new();
        }
        let word_len = (bits + 63) / 64;
        Self {
            words: vec![0u64; word_len],
            len: bits,
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

    pub fn and_inplace(&mut self, other: &BitVec) {
        debug_assert_eq!(self.len, other.len, "BitVec length mismatch");
        let len = self.len;
        let full_words = len / 64;
        let rem_bits = len % 64;

        let mut ones: usize = 0;
        for i in 0..full_words {
            let w = self.words.get_mut(i).expect("BitVec word missing");
            *w &= other.words.get(i).copied().unwrap_or(0);
            ones = ones.saturating_add(w.count_ones() as usize);
        }

        if rem_bits > 0 {
            let mask = (1u64 << rem_bits) - 1;
            if let Some(last) = self.words.get_mut(full_words) {
                *last &= other.words.get(full_words).copied().unwrap_or(0);
                *last &= mask;
                ones = ones.saturating_add(last.count_ones() as usize);
            }
        } else if full_words < self.words.len() {
            // Ensure any extra allocated words (shouldn't happen) are cleared.
            for w in self.words.iter_mut().skip(full_words) {
                *w = 0;
            }
        }

        self.ones = ones;
    }

    pub fn or_inplace(&mut self, other: &BitVec) {
        debug_assert_eq!(self.len, other.len, "BitVec length mismatch");
        let len = self.len;
        let full_words = len / 64;
        let rem_bits = len % 64;

        let mut ones: usize = 0;
        for i in 0..full_words {
            let w = self.words.get_mut(i).expect("BitVec word missing");
            *w |= other.words.get(i).copied().unwrap_or(0);
            ones = ones.saturating_add(w.count_ones() as usize);
        }

        if rem_bits > 0 {
            let mask = (1u64 << rem_bits) - 1;
            if let Some(last) = self.words.get_mut(full_words) {
                *last |= other.words.get(full_words).copied().unwrap_or(0);
                *last &= mask;
                ones = ones.saturating_add(last.count_ones() as usize);
            }
        } else if full_words < self.words.len() {
            for w in self.words.iter_mut().skip(full_words) {
                *w = 0;
            }
        }

        self.ones = ones;
    }

    pub fn not_inplace(&mut self) {
        if self.len == 0 {
            return;
        }

        for w in &mut self.words {
            *w = !*w;
        }

        let rem_bits = self.len % 64;
        if rem_bits != 0 {
            let mask = (1u64 << rem_bits) - 1;
            if let Some(last) = self.words.last_mut() {
                *last &= mask;
            }
        }

        self.ones = self.len.saturating_sub(self.ones);
    }

    /// Reconstruct a [`BitVec`] from a raw word buffer and a bit length.
    ///
    /// This is primarily used by persistence layers that store the `u64` words
    /// directly (e.g. SQLite blobs) and want to avoid per-bit rebuilds.
    pub fn from_words(words: Vec<u64>, len: usize) -> Self {
        if len == 0 {
            return Self {
                words,
                len: 0,
                ones: 0,
            };
        }

        let full_words = len / 64;
        let rem_bits = len % 64;
        let mut ones: usize = 0;

        for w in words.iter().take(full_words) {
            ones = ones.saturating_add(w.count_ones() as usize);
        }

        if rem_bits > 0 {
            if let Some(last) = words.get(full_words) {
                let mask = if rem_bits == 64 {
                    u64::MAX
                } else {
                    (1u64 << rem_bits) - 1
                };
                ones = ones.saturating_add((last & mask).count_ones() as usize);
            }
        }

        Self { words, len, ones }
    }
}

impl Default for BitVec {
    fn default() -> Self {
        Self::new()
    }
}
