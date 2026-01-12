import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { Titlebar, type TitlebarProps } from "./Titlebar.js";

export function mountTitlebar(container: HTMLElement, props: TitlebarProps): () => void {
  // Some existing shells already apply `formula-titlebar` styling to the container itself.
  // Since `Titlebar` renders its own `.formula-titlebar` root, avoid ending up with a nested
  // `.formula-titlebar` (double padding/height) by temporarily removing it from the container.
  const containerHadTitlebarClass = container.classList.contains("formula-titlebar");
  if (containerHadTitlebarClass) container.classList.remove("formula-titlebar");

  container.replaceChildren();

  let root: Root | null = createRoot(container);
  root.render(
    <React.StrictMode>
      <Titlebar {...props} />
    </React.StrictMode>,
  );

  return () => {
    root?.unmount();
    root = null;
    container.replaceChildren();
    if (containerHadTitlebarClass) container.classList.add("formula-titlebar");
  };
}
