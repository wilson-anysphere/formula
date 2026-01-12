import React from "react";
import ReactDOM from "react-dom/client";
import { CellFormattingDemo, GridPerformanceHarness, MergedCellsDemo } from "../src/dev";

const params = new URLSearchParams(window.location.search);
const demo = params.get("demo");

const DEMOS: Record<string, React.ComponentType> = {
  perf: GridPerformanceHarness,
  merged: MergedCellsDemo,
  style: CellFormattingDemo,
  formatting: CellFormattingDemo
};

const Root = (demo ? DEMOS[demo] : undefined) ?? MergedCellsDemo;

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
