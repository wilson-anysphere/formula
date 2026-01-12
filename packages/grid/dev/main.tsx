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

function UnknownDemo(props: { demo: string }): React.ReactElement {
  return (
    <div style={{ padding: 16, fontFamily: "system-ui, sans-serif" }}>
      <div style={{ fontWeight: 700, marginBottom: 8 }}>Unknown demo</div>
      <div style={{ marginBottom: 12 }}>
        No dev demo named <code>{props.demo}</code>.
      </div>
      <div>Available demos:</div>
      <ul style={{ marginTop: 6 }}>
        <li>
          <a href="?demo=style">
            <code>?demo=style</code>
          </a>{" "}
          (cell formatting)
        </li>
        <li>
          <a href="?demo=merged">
            <code>?demo=merged</code>
          </a>{" "}
          (merged cells)
        </li>
        <li>
          <a href="?demo=perf">
            <code>?demo=perf</code>
          </a>{" "}
          (performance)
        </li>
      </ul>
    </div>
  );
}

const Root: React.ComponentType =
  demo == null ? MergedCellsDemo : DEMOS[demo] ?? (() => <UnknownDemo demo={demo} />);

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
