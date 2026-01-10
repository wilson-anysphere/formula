import * as React from "react";

import type { PageMarginsInches, PageSetup, Scaling } from "./types";
import { t } from "../i18n/index.js";

type Props = {
  value: PageSetup;
  onChange: (next: PageSetup) => void;
  onClose: () => void;
};

function updateMargins(
  margins: PageMarginsInches,
  key: keyof PageMarginsInches,
  value: number,
): PageMarginsInches {
  return { ...margins, [key]: value };
}

export function PageSetupDialog({ value, onChange, onClose }: Props) {
  const setScaling = (scaling: Scaling) => onChange({ ...value, scaling });

  return (
    <div style={{ padding: 12, width: 420 }}>
      <h3 style={{ marginTop: 0 }}>{t("print.pageSetup.title")}</h3>

      <label style={{ display: "block", marginBottom: 8 }}>
        {t("print.pageSetup.orientation.label")}
        <select
          value={value.orientation}
          onChange={(e) =>
            onChange({ ...value, orientation: e.target.value as any })
          }
          style={{ marginInlineStart: 8 }}
        >
          <option value="portrait">{t("print.pageSetup.orientation.portrait")}</option>
          <option value="landscape">{t("print.pageSetup.orientation.landscape")}</option>
        </select>
      </label>

      <label style={{ display: "block", marginBottom: 8 }}>
        {t("print.pageSetup.paperSize.label")}
        <input
          type="number"
          value={value.paperSize}
          onChange={(e) =>
            onChange({ ...value, paperSize: Number(e.target.value) })
          }
          style={{ marginInlineStart: 8, width: 100 }}
        />
      </label>

      <fieldset style={{ border: "1px solid var(--dialog-border)", padding: 8 }}>
        <legend>{t("print.pageSetup.scaling.legend")}</legend>

        <label style={{ display: "block", marginBottom: 6 }}>
          <input
            type="radio"
            checked={value.scaling.kind === "percent"}
            onChange={() => setScaling({ kind: "percent", percent: 100 })}
          />{" "}
          {t("print.pageSetup.scaling.adjustTo")}
          <input
            type="number"
            min={10}
            max={400}
            value={value.scaling.kind === "percent" ? value.scaling.percent : 100}
            onChange={(e) =>
              setScaling({ kind: "percent", percent: Number(e.target.value) })
            }
            style={{ marginInlineStart: 8, width: 80 }}
            disabled={value.scaling.kind !== "percent"}
          />
          %
        </label>

        <label style={{ display: "block" }}>
          <input
            type="radio"
            checked={value.scaling.kind === "fitTo"}
            onChange={() => setScaling({ kind: "fitTo", widthPages: 1, heightPages: 0 })}
          />{" "}
          {t("print.pageSetup.scaling.fitTo")}
          <input
            type="number"
            min={0}
            value={value.scaling.kind === "fitTo" ? value.scaling.widthPages : 1}
            onChange={(e) =>
              setScaling({
                kind: "fitTo",
                widthPages: Number(e.target.value),
                heightPages: value.scaling.kind === "fitTo" ? value.scaling.heightPages : 0,
              })
            }
            style={{ marginInlineStart: 8, width: 60 }}
            disabled={value.scaling.kind !== "fitTo"}
          />{" "}
          {t("print.pageSetup.scaling.pagesWide")}{" "}
          <input
            type="number"
            min={0}
            value={value.scaling.kind === "fitTo" ? value.scaling.heightPages : 0}
            onChange={(e) =>
              setScaling({
                kind: "fitTo",
                widthPages: value.scaling.kind === "fitTo" ? value.scaling.widthPages : 1,
                heightPages: Number(e.target.value),
              })
            }
            style={{ width: 60 }}
            disabled={value.scaling.kind !== "fitTo"}
          />{" "}
          {t("print.pageSetup.scaling.tall")}
        </label>
      </fieldset>

      <fieldset style={{ border: "1px solid var(--dialog-border)", padding: 8, marginTop: 10 }}>
        <legend>{t("print.pageSetup.margins.legend")}</legend>

        {(["left", "right", "top", "bottom"] as const).map((k) => (
          <label key={k} style={{ display: "inline-block", marginInlineEnd: 10 }}>
            {t(`print.pageSetup.margins.${k}`)}
            <input
              type="number"
              step={0.01}
              value={value.margins[k]}
              onChange={(e) =>
                onChange({
                  ...value,
                  margins: updateMargins(value.margins, k, Number(e.target.value)),
                })
              }
              style={{ marginInlineStart: 6, width: 80 }}
            />
          </label>
        ))}
      </fieldset>

      <div style={{ marginTop: 12, display: "flex", justifyContent: "flex-end" }}>
        <button onClick={onClose}>{t("print.pageSetup.close")}</button>
      </div>
    </div>
  );
}
