import React from "react";
import { createRoot, type Root } from "react-dom/client";

import type { RibbonActions } from "./ribbonSchema.js";
import { Ribbon } from "./Ribbon.js";

export function mountRibbon(container: HTMLElement, actions: RibbonActions): () => void {
  const root: Root = createRoot(container);
  root.render(<Ribbon actions={actions} />);
  return () => root.unmount();
}

