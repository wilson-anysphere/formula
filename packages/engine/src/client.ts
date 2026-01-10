export type EngineClient = {
  init: () => Promise<void>;
  ping: () => Promise<string>;
  terminate: () => void;
};

type EngineRequest = {
  id: number;
  method: "init" | "ping";
  params?: unknown;
};

type EngineResponse =
  | { id: number; ok: true; result: unknown }
  | { id: number; ok: false; error: string };

export function createEngineClient(): EngineClient {
  // Vite supports Worker construction via `new URL(..., import.meta.url)` and will
  // bundle the Worker entrypoint correctly for both dev and production builds.
  const worker = new Worker(new URL("./engine.worker.ts", import.meta.url), {
    type: "module",
  });

  let nextId = 1;
  const pending = new Map<
    number,
    { resolve: (value: unknown) => void; reject: (error: Error) => void }
  >();

  worker.onmessage = (event) => {
    const message = event.data as EngineResponse;
    const handler = pending.get(message.id);
    if (!handler) return;

    pending.delete(message.id);

    if (message.ok) {
      handler.resolve(message.result);
      return;
    }

    handler.reject(new Error(message.error));
  };

  worker.onerror = (event) => {
    const error = new Error(event.message);
    for (const handler of pending.values()) handler.reject(error);
    pending.clear();
  };

  function call<T>(method: EngineRequest["method"], params?: unknown) {
    const id = nextId++;
    const request: EngineRequest = { id, method, params };
    worker.postMessage(request);

    return new Promise<T>((resolve, reject) => {
      pending.set(id, { resolve, reject });
    });
  }

  return {
    init: () => call<void>("init"),
    ping: () => call<string>("ping"),
    terminate: () => worker.terminate(),
  };
}

