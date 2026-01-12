import React from "react";
import { createRoot, type Root } from "react-dom/client";

import { Titlebar, type TitlebarProps } from "./Titlebar.js";

export type MountedTitlebar = {
  update: (props: TitlebarProps) => void;
  dispose: () => void;
};

export function mountTitlebar(container: HTMLElement, props: TitlebarProps): MountedTitlebar {
  // Some shells may already apply `formula-titlebar` styling to the container itself.
  // Since `Titlebar` renders its own `.formula-titlebar` root, avoid ending up with a nested
  // `.formula-titlebar` (double padding/height) by temporarily removing it from the container.
  const containerHadTitlebarClass = container.classList.contains("formula-titlebar");
  if (containerHadTitlebarClass) container.classList.remove("formula-titlebar");

  container.replaceChildren();

  let root: Root | null = createRoot(container);
  const render = (nextProps: TitlebarProps) => {
    root?.render(
      <React.StrictMode>
        <Titlebar {...nextProps} />
      </React.StrictMode>,
    );
  };
  render(props);

  return {
    update: render,
    dispose: () => {
      root?.unmount();
      root = null;
      container.replaceChildren();
      if (containerHadTitlebarClass) container.classList.add("formula-titlebar");
    },
  };
}

