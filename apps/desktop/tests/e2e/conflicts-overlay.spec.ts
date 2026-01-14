import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("conflict overlay styling", () => {
  test("conflict toast/dialog use tokenized radii and compact spacing", async ({ page }) => {
    await gotoDesktop(page, "/", { waitForIdle: false, waitForContextMenu: false });
    await expect(page.locator("#grid")).toHaveCount(1);

    const styles = await page.evaluate(() => {
      const grid = document.querySelector<HTMLElement>("#grid") ?? document.body;

      const root = document.documentElement;
      const computedRoot = getComputedStyle(root);
      const prevRadius = computedRoot.getPropertyValue("--radius").trim();
      const prevRadiusSm = computedRoot.getPropertyValue("--radius-sm").trim();

      // Override tokens to prove the conflict UI uses CSS variables rather than hardcoded radii.
      root.style.setProperty("--radius", "17px");
      root.style.setProperty("--radius-sm", "9px");

      const overlay = document.createElement("div");
      overlay.className = "conflict-ui-overlay";

      const toastRoot = document.createElement("div");
      toastRoot.className = "conflict-ui-toast-root";

      const toast = document.createElement("div");
      toast.appendChild(document.createTextNode("Conflict detected"));
      const toastButton = document.createElement("button");
      toastButton.textContent = "Resolve";
      toast.appendChild(toastButton);

      toastRoot.appendChild(toast);
      overlay.appendChild(toastRoot);

      const dialog = document.createElement("div");
      dialog.className = "conflict-ui-dialog-root";
      dialog.textContent = "Dialog";
      overlay.appendChild(dialog);

      grid.appendChild(overlay);

      const toastStyle = getComputedStyle(toast);
      const buttonStyle = getComputedStyle(toastButton);
      const dialogStyle = getComputedStyle(dialog);

      const result = {
        toast: {
          paddingTop: toastStyle.paddingTop,
          paddingRight: toastStyle.paddingRight,
          paddingBottom: toastStyle.paddingBottom,
          paddingLeft: toastStyle.paddingLeft,
          borderWidth: toastStyle.borderWidth,
          borderStyle: toastStyle.borderStyle,
          borderRadius: toastStyle.borderTopLeftRadius,
        },
        button: {
          borderRadius: buttonStyle.borderTopLeftRadius,
        },
        dialog: {
          paddingTop: dialogStyle.paddingTop,
          paddingRight: dialogStyle.paddingRight,
          paddingBottom: dialogStyle.paddingBottom,
          paddingLeft: dialogStyle.paddingLeft,
          borderWidth: dialogStyle.borderWidth,
          borderStyle: dialogStyle.borderStyle,
          borderRadius: dialogStyle.borderTopLeftRadius,
        },
      };

      overlay.remove();
      // Restore any mutated theme tokens so subsequent tests remain isolated.
      if (prevRadius) root.style.setProperty("--radius", prevRadius);
      else root.style.removeProperty("--radius");
      if (prevRadiusSm) root.style.setProperty("--radius-sm", prevRadiusSm);
      else root.style.removeProperty("--radius-sm");

      return result;
    });

    expect(styles.toast.borderWidth).toBe("1px");
    expect(styles.toast.borderStyle).toBe("solid");
    expect(styles.toast.borderRadius).toBe("17px");

    expect(styles.button.borderRadius).toBe("9px");

    // Spreadsheet-density spacing: compact padding for toast surfaces.
    expect(styles.toast.paddingTop).toBe("8px");
    expect(styles.toast.paddingRight).toBe("10px");
    expect(styles.toast.paddingBottom).toBe("8px");
    expect(styles.toast.paddingLeft).toBe("10px");

    expect(styles.dialog.borderWidth).toBe("1px");
    expect(styles.dialog.borderStyle).toBe("solid");
    expect(styles.dialog.borderRadius).toBe("17px");
  });

  test("conflict surfaces remain visibly outlined in forced-colors mode", async ({ page }) => {
    await page.emulateMedia({ forcedColors: "active" });
    await gotoDesktop(page, "/", { waitForIdle: false, waitForContextMenu: false });
    await expect(page.locator("#grid")).toHaveCount(1);

    const border = await page.evaluate(() => {
      const grid = document.querySelector<HTMLElement>("#grid") ?? document.body;
      const overlay = document.createElement("div");
      overlay.className = "conflict-ui-overlay";
      const toastRoot = document.createElement("div");
      toastRoot.className = "conflict-ui-toast-root";
      const toast = document.createElement("div");
      toast.textContent = "Conflict detected";
      toastRoot.appendChild(toast);
      overlay.appendChild(toastRoot);
      grid.appendChild(overlay);

      const style = getComputedStyle(toast);
      const result = {
        borderWidth: style.borderWidth,
        borderStyle: style.borderStyle,
        backgroundColor: style.backgroundColor,
      };

      overlay.remove();
      return result;
    });

    expect(border.borderWidth).toBe("1px");
    expect(border.borderStyle).toBe("solid");
    expect(border.backgroundColor).not.toBe("");
  });
});
