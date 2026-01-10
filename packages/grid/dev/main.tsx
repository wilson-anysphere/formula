import React from "react";
import ReactDOM from "react-dom/client";
import { GridPerformanceHarness } from "../src/dev/GridPerformanceHarness";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <GridPerformanceHarness />
  </React.StrictMode>
);

