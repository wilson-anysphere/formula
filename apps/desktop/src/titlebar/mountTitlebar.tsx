import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { Titlebar, type TitlebarProps } from "./Titlebar.js";

export function mountTitlebar(container: HTMLElement, props: TitlebarProps): () => void {
  container.replaceChildren();

  const reactHost = document.createElement("div");
  container.appendChild(reactHost);

  let root: Root | null = createRoot(reactHost);
  root.render(
    <React.StrictMode>
      <Titlebar {...props} />
    </React.StrictMode>,
  );

  return () => {
    root?.unmount();
    root = null;
    reactHost.remove();
  };
}

