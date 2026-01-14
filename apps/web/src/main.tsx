import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

// The extension test harness is only used by Playwright e2e tests and manual debugging.
// Load it lazily to keep the default web preview bundle small.
if (typeof window !== "undefined") {
  const params = new URLSearchParams(window.location.search);
  if (params.has("extTest")) {
    void import("./extensionTestHarness")
      .then(({ setupExtensionTestHarness }) => setupExtensionTestHarness())
      .catch(() => {
        // Best-effort: the extension test harness is only used by e2e tests / debugging and should
        // never crash the main app if it fails to load.
      });
  }
}
