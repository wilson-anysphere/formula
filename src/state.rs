use serde::{Deserialize, Serialize};
use std::sync::Arc;

use crate::stores::{
    InMemoryLatestValueStore, InMemoryMetricRegistry, InMemoryRunStore, LatestValueStore,
    MetricRegistry, RunStore,
};

pub type RunId = String;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct RunMetadata {
    pub id: RunId,
    pub name: String,
    pub created_at_ms: i64,
}

impl RunMetadata {
    pub fn new(id: impl Into<RunId>, name: impl Into<String>, created_at_ms: i64) -> Self {
        Self {
            id: id.into(),
            name: name.into(),
            created_at_ms,
        }
    }
}

#[derive(Debug, Clone)]
pub struct Limits {
    pub max_metrics: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self { max_metrics: 1000 }
    }
}

#[derive(Clone)]
pub struct AppState {
    pub(crate) runs: Arc<dyn RunStore>,
    pub(crate) metrics: Arc<dyn MetricRegistry>,
    pub(crate) latest: Arc<dyn LatestValueStore>,
    pub(crate) limits: Limits,
}

impl AppState {
    pub fn new(limits: Limits) -> Self {
        Self::with_stores(
            limits,
            Arc::new(InMemoryRunStore::default()),
            Arc::new(InMemoryMetricRegistry::default()),
            Arc::new(InMemoryLatestValueStore::default()),
        )
    }

    pub fn with_stores(
        limits: Limits,
        runs: Arc<dyn RunStore>,
        metrics: Arc<dyn MetricRegistry>,
        latest: Arc<dyn LatestValueStore>,
    ) -> Self {
        Self {
            runs,
            metrics,
            latest,
            limits,
        }
    }

    pub fn create_run(&self, run: RunMetadata) {
        self.runs.insert(run);
    }

    pub fn register_metric(&self, run_id: &str, metric_name: &str) {
        self.metrics.register(run_id, metric_name);
    }

    /// Minimal "ingest" API for MVP: updates latest scalar value and registers the metric.
    pub fn ingest_scalar(&self, run_id: &str, metric_name: &str, value: f64) {
        self.metrics.register(run_id, metric_name);
        self.latest.set_latest(run_id, metric_name, value);
    }
}
