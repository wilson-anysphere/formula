import React from "react";
import ReactDOM from "react-dom/client";
import { GridPerformanceHarness, MergedCellsDemo } from "../src/dev";

const params = new URLSearchParams(window.location.search);
const demo = params.get("demo");

const Root = demo === "perf" ? GridPerformanceHarness : MergedCellsDemo;

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
