/// <reference lib="webworker" />

type EngineRequest = {
  id: number;
  method: "init" | "ping";
  params?: unknown;
};

type EngineResponse =
  | { id: number; ok: true; result: unknown }
  | { id: number; ok: false; error: string };

let initialized = false;

async function handle(method: EngineRequest["method"], _params?: unknown) {
  switch (method) {
    case "init": {
      // Placeholder for WASM initialization. The intention is that this Worker
      // will `WebAssembly.instantiate` the Rust/WASM engine and keep it alive for
      // the lifetime of the document/app.
      initialized = true;
      return;
    }
    case "ping": {
      return initialized ? "pong" : "not-initialized";
    }
  }
}

self.onmessage = async (event: MessageEvent<EngineRequest>) => {
  const { id, method, params } = event.data;

  try {
    const result = await handle(method, params);
    const response: EngineResponse = { id, ok: true, result };
    self.postMessage(response);
  } catch (error) {
    const response: EngineResponse = {
      id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    };
    self.postMessage(response);
  }
};

