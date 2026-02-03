use swarm_wandb::{http, AppState, Limits};
use tower_http::trace::TraceLayer;
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt};

#[tokio::main]
async fn main() {
    tracing_subscriber::registry()
        .with(tracing_subscriber::EnvFilter::from_default_env())
        .with(tracing_subscriber::fmt::layer())
        .init();

    let state = AppState::new(Limits::default());
    let app = http::router(state).layer(TraceLayer::new_for_http());

    let listener = tokio::net::TcpListener::bind("127.0.0.1:8080")
        .await
        .expect("bind");
    tracing::info!(
        "listening on {}",
        listener.local_addr().expect("local_addr")
    );

    axum::serve(listener, app).await.expect("serve");
}
