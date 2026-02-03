use dashmap::{DashMap, DashSet};
use std::sync::Arc;

use crate::state::RunMetadata;

pub trait RunStore: Send + Sync {
    fn insert(&self, run: RunMetadata);
    fn get(&self, run_id: &str) -> Option<RunMetadata>;
    fn contains(&self, run_id: &str) -> bool;
}

pub trait MetricRegistry: Send + Sync {
    fn register(&self, run_id: &str, metric_name: &str);
    fn count(&self, run_id: &str) -> usize;
    fn list_limited(&self, run_id: &str, limit: usize) -> Vec<String>;
}

pub trait LatestValueStore: Send + Sync {
    fn set_latest(&self, run_id: &str, metric_name: &str, value: f64);
    fn get_latest(&self, run_id: &str, metric_name: &str) -> Option<f64>;
    fn count(&self, run_id: &str) -> usize;
    fn list_latest_limited(&self, run_id: &str, limit: usize) -> Vec<(String, f64)>;
}

#[derive(Clone, Default)]
pub struct InMemoryRunStore {
    runs: Arc<DashMap<String, RunMetadata>>,
}

impl InMemoryRunStore {
    pub fn insert(&self, run: RunMetadata) {
        self.runs.insert(run.id.clone(), run);
    }

    pub fn get(&self, run_id: &str) -> Option<RunMetadata> {
        self.runs.get(run_id).map(|run| run.clone())
    }

    pub fn contains(&self, run_id: &str) -> bool {
        self.runs.contains_key(run_id)
    }
}

#[derive(Clone, Default)]
pub struct InMemoryMetricRegistry {
    metrics_by_run: Arc<DashMap<String, Arc<DashSet<String>>>>,
}

impl InMemoryMetricRegistry {
    pub fn register(&self, run_id: &str, metric_name: &str) {
        let set = self
            .metrics_by_run
            .entry(run_id.to_string())
            .or_insert_with(|| Arc::new(DashSet::new()))
            .clone();
        set.insert(metric_name.to_string());
    }

    pub fn count(&self, run_id: &str) -> usize {
        self.metrics_by_run
            .get(run_id)
            .map(|set| set.len())
            .unwrap_or(0)
    }

    pub fn list_limited(&self, run_id: &str, limit: usize) -> Vec<String> {
        self.metrics_by_run
            .get(run_id)
            .map(|set| set.iter().take(limit).map(|name| name.clone()).collect())
            .unwrap_or_default()
    }
}

#[derive(Clone, Default)]
pub struct InMemoryLatestValueStore {
    latest_by_run: Arc<DashMap<String, Arc<DashMap<String, f64>>>>,
}

impl InMemoryLatestValueStore {
    pub fn set_latest(&self, run_id: &str, metric_name: &str, value: f64) {
        let metrics = self
            .latest_by_run
            .entry(run_id.to_string())
            .or_insert_with(|| Arc::new(DashMap::new()))
            .clone();
        metrics.insert(metric_name.to_string(), value);
    }

    pub fn get_latest(&self, run_id: &str, metric_name: &str) -> Option<f64> {
        self.latest_by_run
            .get(run_id)
            .and_then(|metrics| metrics.get(metric_name).map(|value| *value))
    }

    pub fn count(&self, run_id: &str) -> usize {
        self.latest_by_run
            .get(run_id)
            .map(|metrics| metrics.len())
            .unwrap_or(0)
    }

    pub fn list_latest_limited(&self, run_id: &str, limit: usize) -> Vec<(String, f64)> {
        self.latest_by_run
            .get(run_id)
            .map(|metrics| {
                metrics
                    .iter()
                    .take(limit)
                    .map(|entry| (entry.key().clone(), *entry.value()))
                    .collect()
            })
            .unwrap_or_default()
    }
}

impl RunStore for InMemoryRunStore {
    fn insert(&self, run: RunMetadata) {
        InMemoryRunStore::insert(self, run);
    }

    fn get(&self, run_id: &str) -> Option<RunMetadata> {
        InMemoryRunStore::get(self, run_id)
    }

    fn contains(&self, run_id: &str) -> bool {
        InMemoryRunStore::contains(self, run_id)
    }
}

impl MetricRegistry for InMemoryMetricRegistry {
    fn register(&self, run_id: &str, metric_name: &str) {
        InMemoryMetricRegistry::register(self, run_id, metric_name);
    }

    fn count(&self, run_id: &str) -> usize {
        InMemoryMetricRegistry::count(self, run_id)
    }

    fn list_limited(&self, run_id: &str, limit: usize) -> Vec<String> {
        InMemoryMetricRegistry::list_limited(self, run_id, limit)
    }
}

impl LatestValueStore for InMemoryLatestValueStore {
    fn set_latest(&self, run_id: &str, metric_name: &str, value: f64) {
        InMemoryLatestValueStore::set_latest(self, run_id, metric_name, value);
    }

    fn get_latest(&self, run_id: &str, metric_name: &str) -> Option<f64> {
        InMemoryLatestValueStore::get_latest(self, run_id, metric_name)
    }

    fn count(&self, run_id: &str) -> usize {
        InMemoryLatestValueStore::count(self, run_id)
    }

    fn list_latest_limited(&self, run_id: &str, limit: usize) -> Vec<(String, f64)> {
        InMemoryLatestValueStore::list_latest_limited(self, run_id, limit)
    }
}
