use axum::{
    body::Body,
    http::{Request, StatusCode},
};
use http_body_util::BodyExt;
use std::collections::BTreeMap;
use tower::ServiceExt;
use uuid::Uuid;

use swarm_wandb::{http, AppState, Limits, RunMetadata};

async fn read_json<T: serde::de::DeserializeOwned>(response: axum::response::Response) -> T {
    let body = response
        .into_body()
        .collect()
        .await
        .expect("collect body")
        .to_bytes();
    serde_json::from_slice(&body).expect("valid json")
}

#[tokio::test]
async fn query_api_returns_run_metrics_and_latest_values() {
    let state = AppState::new(Limits { max_metrics: 100 });

    let run_id = Uuid::new_v4().to_string();
    let run = RunMetadata::new(run_id.clone(), "my-run", 0);
    state.create_run(run.clone());

    state.register_metric(&run_id, "loss");
    state.register_metric(&run_id, "accuracy");

    state.ingest_scalar(&run_id, "loss", 1.23);
    state.ingest_scalar(&run_id, "accuracy", 0.9);

    let app = http::router(state);

    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let body: RunMetadata = read_json(response).await;
    assert_eq!(body, run);

    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/metrics"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let metrics: Vec<String> = read_json(response).await;
    assert_eq!(metrics, vec!["accuracy".to_string(), "loss".to_string()]);

    let response = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/latest"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let latest: BTreeMap<String, f64> = read_json(response).await;
    assert_eq!(latest.get("loss").copied(), Some(1.23));
    assert_eq!(latest.get("accuracy").copied(), Some(0.9));
}

#[tokio::test]
async fn metrics_and_latest_are_bounded_by_max_metrics() {
    let state = AppState::new(Limits { max_metrics: 2 });
    let run_id = Uuid::new_v4().to_string();
    state.create_run(RunMetadata::new(run_id.clone(), "run", 0));

    for (name, value) in [("a", 1.0), ("b", 2.0), ("c", 3.0)] {
        state.register_metric(&run_id, name);
        state.ingest_scalar(&run_id, name, value);
    }

    let app = http::router(state);

    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/metrics"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    assert_eq!(
        response.headers().get("x-metrics-total").unwrap(),
        "3",
        "reports total metrics"
    );
    assert_eq!(
        response.headers().get("x-metrics-truncated").unwrap(),
        "true",
        "reports truncation"
    );
    let metrics: Vec<String> = read_json(response).await;
    assert_eq!(metrics.len(), 2);

    let response = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/latest"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let latest: BTreeMap<String, f64> = read_json(response).await;
    assert_eq!(latest.len(), 2);
}

#[tokio::test]
async fn limit_query_param_is_respected_but_never_exceeds_max_metrics() {
    let state = AppState::new(Limits { max_metrics: 2 });
    let run_id = Uuid::new_v4().to_string();
    state.create_run(RunMetadata::new(run_id.clone(), "run", 0));

    for (name, value) in [("a", 1.0), ("b", 2.0), ("c", 3.0)] {
        state.register_metric(&run_id, name);
        state.ingest_scalar(&run_id, name, value);
    }

    let app = http::router(state);

    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/metrics?limit=1"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    assert_eq!(response.headers().get("x-metrics-total").unwrap(), "3");
    assert_eq!(
        response.headers().get("x-metrics-truncated").unwrap(),
        "true"
    );
    let metrics: Vec<String> = read_json(response).await;
    assert_eq!(metrics.len(), 1);

    let response = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/latest?limit=99"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let latest: BTreeMap<String, f64> = read_json(response).await;
    assert_eq!(latest.len(), 2, "cannot exceed max_metrics=2");
}

#[tokio::test]
async fn latest_endpoint_omits_metrics_without_values() {
    let state = AppState::new(Limits { max_metrics: 100 });
    let run_id = Uuid::new_v4().to_string();
    state.create_run(RunMetadata::new(run_id.clone(), "run", 0));

    state.register_metric(&run_id, "no_value_yet");
    state.ingest_scalar(&run_id, "has_value", 42.0);

    let app = http::router(state);

    let response = app
        .clone()
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/metrics"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let metrics: Vec<String> = read_json(response).await;
    assert_eq!(
        metrics,
        vec!["has_value".to_string(), "no_value_yet".to_string()]
    );

    let response = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}/latest"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::OK);
    let latest: BTreeMap<String, f64> = read_json(response).await;
    assert_eq!(latest.get("has_value").copied(), Some(42.0));
    assert!(
        !latest.contains_key("no_value_yet"),
        "missing values are omitted"
    );
}

#[tokio::test]
async fn unknown_run_returns_404() {
    let state = AppState::new(Limits { max_metrics: 10 });
    let app = http::router(state);

    let run_id = Uuid::new_v4().to_string();
    let response = app
        .oneshot(
            Request::builder()
                .uri(format!("/api/runs/{run_id}"))
                .body(Body::empty())
                .unwrap(),
        )
        .await
        .unwrap();
    assert_eq!(response.status(), StatusCode::NOT_FOUND);
}
