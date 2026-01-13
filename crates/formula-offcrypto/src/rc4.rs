/// Minimal RC4 implementation (KSA + PRGA) used for CryptoAPI standard encryption.
///
/// This is intentionally small and self-contained to avoid pulling in extra
/// cipher crates for a legacy format.
pub(crate) struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    pub(crate) fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");
        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }
        let mut j: u8 = 0;
        for i in 0..256u16 {
            let si = s[i as usize];
            j = j
                .wrapping_add(si)
                .wrapping_add(key[(i as usize) % key.len()]);
            s.swap(i as usize, j as usize);
        }
        Rc4 { s, i: 0, j: 0 }
    }

    pub(crate) fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[idx as usize];
            *b ^= k;
        }
    }
}
