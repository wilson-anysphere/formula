#![forbid(unsafe_code)]

use std::collections::{HashMap, VecDeque};
use std::hash::Hash;

#[derive(Debug, Clone, Copy)]
pub struct PageCacheConfig {
    pub max_entries: usize,
}

impl Default for PageCacheConfig {
    fn default() -> Self {
        Self { max_entries: 64 }
    }
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct CacheStats {
    pub hits: u64,
    pub misses: u64,
    pub evictions: u64,
}

/// A small, dependency-free LRU-ish cache with amortized O(1) operations.
///
/// Implementation strategy:
/// - A `HashMap` stores the current value + a monotonic "token".
/// - An `order` queue stores (key, token) entries. Old entries become stale when
///   a key is re-accessed and given a new token.
/// - Eviction pops from the front until it finds a non-stale entry.
#[derive(Debug)]
pub struct LruCache<K, V> {
    cap: usize,
    next_token: u64,
    map: HashMap<K, (V, u64)>,
    order: VecDeque<(K, u64)>,
    stats: CacheStats,
}

impl<K, V> LruCache<K, V>
where
    K: Eq + Hash + Clone,
    V: Clone,
{
    pub fn new(cap: usize) -> Self {
        Self {
            cap,
            next_token: 1,
            map: HashMap::new(),
            order: VecDeque::new(),
            stats: CacheStats::default(),
        }
    }

    pub fn stats(&self) -> CacheStats {
        self.stats
    }

    pub fn remove(&mut self, key: &K) -> Option<V> {
        self.map.remove(key).map(|(value, _)| value)
    }

    pub fn remove_if<F>(&mut self, mut predicate: F)
    where
        F: FnMut(&K) -> bool,
    {
        if self.map.is_empty() {
            return;
        }

        let keys: Vec<K> = self
            .map
            .keys()
            .filter(|k| predicate(k))
            .cloned()
            .collect();

        for key in keys {
            self.map.remove(&key);
        }
    }

    pub fn get(&mut self, key: &K) -> Option<V> {
        if self.cap == 0 {
            self.stats.misses += 1;
            return None;
        }

        match self.map.get_mut(key) {
            Some((value, token)) => {
                self.stats.hits += 1;
                let new_token = self.next_token;
                self.next_token = self.next_token.wrapping_add(1);
                *token = new_token;
                self.order.push_back((key.clone(), new_token));
                Some(value.clone())
            }
            None => {
                self.stats.misses += 1;
                None
            }
        }
    }

    pub fn insert(&mut self, key: K, value: V) {
        if self.cap == 0 {
            return;
        }

        let token = self.next_token;
        self.next_token = self.next_token.wrapping_add(1);
        self.order.push_back((key.clone(), token));
        self.map.insert(key, (value, token));

        self.evict_if_needed();
        self.maybe_compact_order();
    }

    fn evict_if_needed(&mut self) {
        while self.map.len() > self.cap {
            let Some((candidate_key, candidate_token)) = self.order.pop_front() else {
                break;
            };

            let Some((_, current_token)) = self.map.get(&candidate_key) else {
                continue;
            };

            if *current_token == candidate_token {
                self.map.remove(&candidate_key);
                self.stats.evictions += 1;
            }
        }
    }

    fn maybe_compact_order(&mut self) {
        // Avoid unbounded growth when the same small set of keys are hit repeatedly.
        // This is a best-effort compaction; correctness does not depend on it.
        let max_len = self.cap.saturating_mul(16).max(1024);
        if self.order.len() <= max_len {
            return;
        }

        let mut new_order = VecDeque::new();
        let _ = new_order.try_reserve(self.map.len());
        for (key, (_, token)) in self.map.iter() {
            new_order.push_back((key.clone(), *token));
        }
        self.order = new_order;
    }
}
