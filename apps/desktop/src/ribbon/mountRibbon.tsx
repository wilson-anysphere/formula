import React from "react";
import { createRoot, type Root } from "react-dom/client";
import { flushSync } from "react-dom";

import type { RibbonActions, RibbonSchema } from "./ribbonSchema.js";
import { Ribbon } from "./Ribbon.js";

export interface MountRibbonOptions {
  schema?: RibbonSchema;
  initialTabId?: string;
}

export function mountRibbon(container: HTMLElement, actions: RibbonActions, options: MountRibbonOptions = {}): () => void {
  const root: Root = createRoot(container);
  const element = <Ribbon actions={actions} schema={options.schema} initialTabId={options.initialTabId} />;
  flushSync(() => {
    root.render(element);
  });
  return () => root.unmount();
}
