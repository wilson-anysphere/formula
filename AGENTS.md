# Goal

Build a performant weights and biases clone using rust as a server and python as an API client. Choice of browser framework and language is up to the agents, but performance is the most important aspect. On charts, scrubbing and zooming should be instant. selecting, deselecting, and searching for runs should also be instant.
The clone should support connecting to a wandb server and calling wandb.init() on a run, then submitting per step metrics.
The client app should support all features of weights and biases, and performantly render multiple line charts per run.
It should work in a distributed training setup in a "baterries-included" manner and the python client should never crash the run or consume material resources on the same process.

Use puppeteer or other software to test browser features.
Include extensive testing.

# Perf-First Experiment Tracking (Single Tenant) — Concise Plan

## Goals
- **Extremely low training-loop overhead**: `log()` should be near-constant time, non-blocking, minimal allocations.
- **Incredibly fast charting** across pan/zoom/scroll with many runs via **multi-resolution tiles**.
- **High ingest throughput** with durability: at-least-once + idempotent frames.
- **Single tenant/user**: simplify auth/multi-tenancy, focus on performance-critical paths.

---

## Major Components

### 1) Training Client (Python)
**Purpose:** capture metrics/config/logs with minimal overhead and ship in batches.

- **Hot path (`log()`)**
  - Append events into a **bounded ring buffer** (preallocated chunks).
  - Maintain client-side **monotonic step** + `wall_time_ns`.
  - Optional metric registration (names/types) to improve encoding/compressibility.

- **Background flusher**
  - Batches events into **frames** (e.g., 256–4096 events).
  - Encodes frames in **binary** (protobuf/flatbuffers/msgpack) + optional **zstd**.
  - Streams frames to ingest server (prefer **gRPC stream**).

- **Local spool (WAL)**
  - Append frames to a local **append-only** file first.
  - Delete/compact once server ACKs (durability without blocking training).

---

### 2) Ingest Server (Rust, Stateless)
**Purpose:** accept frames quickly, durably persist, ACK fast.

- Implemented in **Rust** for:
  - predictable latency
  - zero/low-copy IO
  - efficient async networking
  - strong memory safety under load
- Exposes **gRPC bidirectional stream**: client sends `BatchFrame`, server returns `AckFrame`.
- Performs minimal validation + optional run auto-create.
- Writes frames to **durable append-only log**.
- ACKs **as soon as persisted** (not after indexing).

**Key property:** **Idempotency**
---

### 3) Durable Event Log
**Purpose:** single source of truth for ingestion; enables async indexing.

- For single-tenant MVP: **append-only files + checkpoints** (fast to build, fewer dependencies)
- Partition key: `run_id` (keeps ordering per run).

---

### 4) Indexer / Materializer (Async Workers)
**Purpose:** consume the event log and build read-optimized views.

Produces:
- **Run State Store**
  - status (running/finished), start/end, tags
  - config blob
  - summary blob
  - **latest metric values** map

- **Time Series Stores**
  - Raw series chunks (optional or limited retention)
  - **Multi-resolution pyramid** for fast charting (see Charting section)

Indexing can lag by seconds; UI reads from materialized stores.

---

### 5) Storage
**Run metadata**
- Simple KV / SQLite / Postgres (single tenant: simplest that works).
- Small and mutable.

**Time series + pyramid**
- Preferred: **ClickHouse** (fast range scans, compression, materialized views) or DuckDb
**Blob store (optional)**
- Local disk / object storage for config/summary blobs and WAL backups.

---

### 6) Web / Query Server (Rust)
**Purpose:** serve UI and programmatic reads with minimal latency.

- Implemented in **Rust** for:
  - high fan-out tile queries
  - efficient binary responses
  - predictable tail latency under load
- Handles:
  - run listing & metadata
  - latest metrics
  - **time-series tile API** for charts
- Speaks HTTP/JSON for metadata and **binary** (Arrow/custom) for tiles.

---
## Charting: “Map Tiles for Time Series” (Fast Pan/Zoom/Many Runs)

### Core idea
Never ship or render raw points when zoomed out. Use a **resolution pyramid** + **tile addressing** to make requests cacheable and rendering cheap.

### 1) Multi-resolution pyramid per (run, metric)
Levels L0..LN where each level aggregates buckets:
- Each bucket stores: `min`, `max`, `first`, `last` (+ optional argmin/argmax for tooltips).
- Preserves spikes better than averages and is extremely render-friendly.

### 2) Time-series tiles
Store pyramid data in fixed-size tiles:
- Tile key: `(metric_id, run_id, level, tile_index)`
- Tile holds e.g. **2048 buckets**.
- Query picks level based on viewport width (`px`) and returns intersecting tiles for [x0, x1].

Benefits:
- Stable cache keys (server + browser + optional CDN).
- Smooth scrolling: reuse adjacent tiles; prefetch beyond viewport.
- Backend work is O(tiles), not O(points).

### 3) Rendering strategy (WebGL-first)
- Render min/max buckets as **vertical segments** (envelope).
- For many runs:
  - Selected/hovered runs: draw as lines/envelopes.
  - Others: **quantile bands** or **density heatmap** (precomputed group tiles).

### 4) Interaction performance
- Tooltips use direct bucket indexing (x → tile → bucket).
- Progressive loading:
  - reproject existing tiles immediately on pan
  - fetch missing tiles async
  - LRU cache of decoded tiles
  - prefetch beyond viewport edges

### 5) Smoothing
- Client-side only, and only when zoomed-in (few thousand visible samples).
- Otherwise ignored or applied to already-downsampled data.

---

## How It Fits Together (Data Flow)

1. Training code calls `log()` → events appended to ring buffer (cheap).
2. Background flusher batches → encodes frame → appends to local WAL → streams to **Rust ingest server**.
3. Ingest server persists frame to durable log → ACKs.
4. Indexer consumes durable log → updates run state + writes raw chunks and pyramid tiles.
5. UI queries **Rust web/query server**:
   - run list/state from run store
   - charts via `GetSeriesTiles(...)` returning binary tiles
6. Frontend renders via WebGL with caching + progressive/predictive tile loading.

---

## MVP Cut (Single Tenant)
- No multi-tenant auth: one API key or local token.
- Minimal entities: runs + metrics; projects optional.
- Keep only: config, summary, scalar metrics (maybe text logs).
- Skip artifacts/lineage/sweeps initially; add later as event types.
