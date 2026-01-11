import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { setupExtensionTestHarness } from "./extensionTestHarness";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

void setupExtensionTestHarness();
