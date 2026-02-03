use axum::{
    extract::{Path, Query, State},
    http::{header::HeaderName, HeaderMap, HeaderValue, StatusCode},
    response::{IntoResponse, Response},
    routing::get,
    Json, Router,
};
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

use crate::state::{AppState, RunMetadata};

const HEADER_METRICS_TOTAL: HeaderName = HeaderName::from_static("x-metrics-total");
const HEADER_METRICS_TRUNCATED: HeaderName = HeaderName::from_static("x-metrics-truncated");

#[derive(Debug, thiserror::Error)]
pub enum ApiError {
    #[error("run not found")]
    RunNotFound,
}

#[derive(Debug, Serialize)]
struct ErrorResponse {
    error: String,
}

#[derive(Debug, Default, Deserialize)]
struct MetricsQuery {
    limit: Option<usize>,
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        let (status, message) = match self {
            ApiError::RunNotFound => (StatusCode::NOT_FOUND, self.to_string()),
        };
        (status, Json(ErrorResponse { error: message })).into_response()
    }
}

pub fn router(state: AppState) -> Router {
    Router::new()
        .route("/api/runs/:run_id", get(get_run))
        .route("/api/runs/:run_id/metrics", get(list_metrics))
        .route("/api/runs/:run_id/latest", get(get_latest))
        .with_state(state)
}

async fn get_run(
    Path(run_id): Path<String>,
    State(state): State<AppState>,
) -> Result<Json<RunMetadata>, ApiError> {
    let run = state.runs.get(&run_id).ok_or(ApiError::RunNotFound)?;
    Ok(Json(run))
}

async fn list_metrics(
    Path(run_id): Path<String>,
    State(state): State<AppState>,
    Query(query): Query<MetricsQuery>,
) -> Result<Response, ApiError> {
    if !state.runs.contains(&run_id) {
        return Err(ApiError::RunNotFound);
    }

    let limit = query
        .limit
        .unwrap_or(state.limits.max_metrics)
        .min(state.limits.max_metrics);
    let total = state.metrics.count(&run_id);
    let truncated = total > limit;

    let mut metrics = state.metrics.list_limited(&run_id, limit);
    metrics.sort();

    Ok(with_metrics_headers(total, truncated, Json(metrics)).into_response())
}

async fn get_latest(
    Path(run_id): Path<String>,
    State(state): State<AppState>,
    Query(query): Query<MetricsQuery>,
) -> Result<Response, ApiError> {
    if !state.runs.contains(&run_id) {
        return Err(ApiError::RunNotFound);
    }

    let limit = query
        .limit
        .unwrap_or(state.limits.max_metrics)
        .min(state.limits.max_metrics);
    let total = state.latest.count(&run_id);
    let truncated = total > limit;

    let mut latest = state.latest.list_latest_limited(&run_id, limit);
    latest.sort_by(|(a, _), (b, _)| a.cmp(b));

    let latest: BTreeMap<String, f64> = latest.into_iter().collect();

    Ok(with_metrics_headers(total, truncated, Json(latest)).into_response())
}

fn with_metrics_headers<T: IntoResponse>(
    total: usize,
    truncated: bool,
    inner: T,
) -> (HeaderMap, T) {
    let mut headers = HeaderMap::new();
    headers.insert(
        HEADER_METRICS_TOTAL,
        HeaderValue::from_str(&total.to_string()).unwrap_or(HeaderValue::from_static("0")),
    );
    headers.insert(
        HEADER_METRICS_TRUNCATED,
        HeaderValue::from_str(&truncated.to_string()).unwrap_or(HeaderValue::from_static("false")),
    );
    (headers, inner)
}
