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
        let mut word_vec = Vec::new();
        let _ = word_vec.try_reserve_exact(words);
        Self {
            words: word_vec,
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
        if index >= self.len {
            return false;
        }
        let word_idx = index / 64;
        let Some(&word) = self.words.get(word_idx) else {
            debug_assert!(false, "BitVec word missing (get)");
            return false;
        };
        let bit = index % 64;
        ((word >> bit) & 1) == 1
    }

    pub fn set(&mut self, index: usize, value: bool) {
        debug_assert!(index < self.len, "BitVec index out of bounds");
        if index >= self.len {
            return;
        }
        let word_idx = index / 64;
        let bit = index % 64;
        let mask = 1u64 << bit;
        let Some(word) = self.words.get_mut(word_idx) else {
            debug_assert!(false, "BitVec word missing (set)");
            return;
        };
        let was_set = (*word & mask) != 0;

        match (was_set, value) {
            (true, false) => {
                *word &= !mask;
                self.ones -= 1;
            }
            (false, true) => {
                *word |= mask;
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

        let required_words = (len + 63) / 64;
        if self.words.len() < required_words {
            let missing = required_words - self.words.len();
            if self.words.try_reserve_exact(missing).is_err() {
                debug_assert!(false, "allocation failed (BitVec.and_inplace repair)");
                return;
            }
            self.words.resize(required_words, 0);
        }

        let mut ones: usize = 0;
        for i in 0..full_words {
            let w = &mut self.words[i];
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

        let required_words = (len + 63) / 64;
        if self.words.len() < required_words {
            let missing = required_words - self.words.len();
            if self.words.try_reserve_exact(missing).is_err() {
                debug_assert!(false, "allocation failed (BitVec.or_inplace repair)");
                return;
            }
            self.words.resize(required_words, 0);
        }

        let mut ones: usize = 0;
        for i in 0..full_words {
            let w = &mut self.words[i];
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

    /// Iterate the indices of set bits (true values) in increasing order.
    pub fn iter_ones(&self) -> impl Iterator<Item = usize> + '_ {
        struct OnesIter<'a> {
            words: &'a [u64],
            len: usize,
            word_idx: usize,
            current_word: u64,
            base: usize,
        }

        impl<'a> Iterator for OnesIter<'a> {
            type Item = usize;

            fn next(&mut self) -> Option<Self::Item> {
                loop {
                    if self.current_word != 0 {
                        let bit = self.current_word.trailing_zeros() as usize;
                        // Clear lowest set bit.
                        self.current_word &= self.current_word - 1;
                        let idx = self.base + bit;
                        if idx < self.len {
                            return Some(idx);
                        }
                        continue;
                    }

                    if self.word_idx >= self.words.len() {
                        return None;
                    }

                    self.current_word = self.words[self.word_idx];
                    self.base = self.word_idx * 64;
                    self.word_idx += 1;
                }
            }
        }

        OnesIter {
            words: &self.words,
            len: self.len,
            word_idx: 0,
            current_word: 0,
            base: 0,
        }
    }

    pub fn extend_constant(&mut self, value: bool, mut count: usize) {
        if count == 0 {
            return;
        }

        // Fill any remaining bits in the current word first.
        let bit = self.len % 64;
        if bit != 0 {
            let available = 64 - bit;
            let take = count.min(available);
            if value {
                let word_idx = self.len / 64;
                let mask = ((1u64 << take) - 1) << bit;
                self.words[word_idx] |= mask;
                self.ones = self.ones.saturating_add(take);
            }
            self.len += take;
            count -= take;
        }

        // Add full words.
        while count >= 64 {
            self.words.push(if value { u64::MAX } else { 0 });
            self.len += 64;
            if value {
                self.ones = self.ones.saturating_add(64);
            }
            count -= 64;
        }

        // Add remaining partial word.
        if count > 0 {
            let word = if value { (1u64 << count) - 1 } else { 0 };
            self.words.push(word);
            self.len += count;
            if value {
                self.ones = self.ones.saturating_add(count);
            }
        }
    }

    /// Reconstruct a [`BitVec`] from a raw word buffer and a bit length.
    ///
    /// This is primarily used by persistence layers that store the `u64` words
    /// directly (e.g. SQLite blobs) and want to avoid per-bit rebuilds.
    pub fn from_words(words: Vec<u64>, len: usize) -> Self {
        if len == 0 {
            return Self {
                words: Vec::new(),
                len: 0,
                ones: 0,
            };
        }

        let full_words = len / 64;
        let rem_bits = len % 64;
        let required_words = (len + 63) / 64;
        let mut words = if words.len() == required_words {
            words
        } else if words.len() > required_words {
            words.into_iter().take(required_words).collect()
        } else {
            // If the persistence layer stored a shorter buffer, treat missing bits as 0.
            let mut out = words;
            out.resize(required_words, 0);
            out
        };
        let mut ones: usize = 0;

        for w in words.iter().take(full_words) {
            ones = ones.saturating_add(w.count_ones() as usize);
        }

        if rem_bits > 0 {
            let mask = (1u64 << rem_bits) - 1;
            if let Some(last) = words.get_mut(full_words) {
                *last &= mask;
                ones = ones.saturating_add(last.count_ones() as usize);
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

#[cfg(test)]
mod tests {
    use super::BitVec;

    #[test]
    fn extend_constant_matches_push() {
        let mut a = BitVec::new();
        for _ in 0..3 {
            a.push(false);
        }
        for _ in 0..70 {
            a.push(true);
        }
        for _ in 0..5 {
            a.push(false);
        }

        let mut b = BitVec::new();
        b.extend_constant(false, 3);
        b.extend_constant(true, 70);
        b.extend_constant(false, 5);

        assert_eq!(a, b);
        assert_eq!(b.count_ones(), 70);
    }

    #[test]
    fn from_words_masks_out_unused_bits() {
        // BitVec length is 3, but the stored word has extra bits set.
        let mut v = BitVec::from_words(vec![0xFFFF_FFFF_FFFF_FFFFu64], 3);
        assert_eq!(v.len(), 3);
        assert_eq!(v.count_ones(), 3);

        // Extending with false should not "inherit" the previously set trailing bits.
        v.extend_constant(false, 5); // len = 8
        assert_eq!(v.len(), 8);
        assert_eq!(v.count_ones(), 3);

        let bits: Vec<bool> = (0..v.len()).map(|i| v.get(i)).collect();
        assert_eq!(bits, vec![true, true, true, false, false, false, false, false]);
    }
}
