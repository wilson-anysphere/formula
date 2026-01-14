import type { TauriEmit, TauriListen } from "./api";

export type { TauriEmit, TauriListen } from "./api";

export function installOpenFileIpc({
  listen,
  emit,
  onOpenPath,
}: {
  listen: TauriListen;
  emit: TauriEmit | null;
  onOpenPath: (path: string) => void;
}): void {
  const openFileListener = listen("open-file", (event) => {
    const payload = (event as any)?.payload;
    if (!Array.isArray(payload)) return;
    const paths = payload
      .map((p) => (typeof p === "string" ? p.trim() : ""))
      .filter((p) => p !== "");
    if (paths.length === 0) return;

    for (const path of paths) {
      onOpenPath(path);
    }
  });

  // Signal that we're ready to receive (and flush any queued) open-file requests.
  void openFileListener
    .then(() => {
      if (!emit) return;
      return Promise.resolve(emit("open-file-ready"));
    })
    .catch((err) => {
      console.error("Failed to signal open-file readiness:", err);
    });
}
