self.addEventListener("message", () => {
  void (async () => {
    try {
      const wasmModuleUrl = new URL("/engine/formula_wasm.js", self.location.origin).toString();
      const mod = await import(wasmModuleUrl);
      await mod.default?.();

      const wb = new mod.WasmWorkbook();
      wb.setCell("A1", 1);
      wb.setCell("A2", "=A1*2");
      wb.recalculate();
      const cell = wb.getCell("A2");

      self.postMessage({ ok: true, value: cell?.value ?? null });
    } catch (err) {
      self.postMessage({
        ok: false,
        error: err instanceof Error ? err.message : String(err)
      });
    }
  })().catch(() => {
    // ignore: worker message handlers swallow returned promises; avoid unhandled rejections.
  });
});
