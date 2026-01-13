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
    <div className="page-setup-dialog__body">
      <h3 className="page-setup-dialog__title">{t("print.pageSetup.title")}</h3>

      <label className="page-setup-dialog__row">
        <span>{t("print.pageSetup.orientation.label")}</span>
        <select
          className="page-setup-dialog__select"
          value={value.orientation}
          onChange={(e) =>
            onChange({ ...value, orientation: e.target.value as any })
          }
        >
          <option value="portrait">{t("print.pageSetup.orientation.portrait")}</option>
          <option value="landscape">{t("print.pageSetup.orientation.landscape")}</option>
        </select>
      </label>

      <label className="page-setup-dialog__row">
        <span>{t("print.pageSetup.paperSize.label")}</span>
        <input
          className="page-setup-dialog__input page-setup-dialog__input--paper-size"
          type="number"
          value={value.paperSize}
          onChange={(e) =>
            onChange({ ...value, paperSize: Number(e.target.value) })
          }
        />
      </label>

      <fieldset className="page-setup-dialog__fieldset">
        <legend className="page-setup-dialog__legend">{t("print.pageSetup.scaling.legend")}</legend>

        <label className="page-setup-dialog__radio-row">
          <input
            type="radio"
            checked={value.scaling.kind === "percent"}
            onChange={() => setScaling({ kind: "percent", percent: 100 })}
          />
          <span>{t("print.pageSetup.scaling.adjustTo")}</span>
          <input
            className="page-setup-dialog__input page-setup-dialog__input--percent"
            type="number"
            min={10}
            max={400}
            value={value.scaling.kind === "percent" ? value.scaling.percent : 100}
            onChange={(e) =>
              setScaling({ kind: "percent", percent: Number(e.target.value) })
            }
            disabled={value.scaling.kind !== "percent"}
          />
          <span>%</span>
        </label>

        <label className="page-setup-dialog__radio-row">
          <input
            type="radio"
            checked={value.scaling.kind === "fitTo"}
            onChange={() => setScaling({ kind: "fitTo", widthPages: 1, heightPages: 0 })}
          />
          <span>{t("print.pageSetup.scaling.fitTo")}</span>
          <input
            className="page-setup-dialog__input page-setup-dialog__input--pages"
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
            disabled={value.scaling.kind !== "fitTo"}
          />
          <span>{t("print.pageSetup.scaling.pagesWide")}</span>
          <input
            className="page-setup-dialog__input page-setup-dialog__input--pages"
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
            disabled={value.scaling.kind !== "fitTo"}
          />
          <span>{t("print.pageSetup.scaling.tall")}</span>
        </label>
      </fieldset>

      <fieldset className="page-setup-dialog__fieldset">
        <legend className="page-setup-dialog__legend">{t("print.pageSetup.margins.legend")}</legend>

        <div className="page-setup-dialog__margins-grid">
          {(["left", "right", "top", "bottom"] as const).map((k) => (
            <label key={k} className="page-setup-dialog__margin">
              <span>{t(`print.pageSetup.margins.${k}`)}</span>
              <input
                className="page-setup-dialog__input page-setup-dialog__input--margin"
                type="number"
                step={0.01}
                value={value.margins[k]}
                onChange={(e) =>
                  onChange({
                    ...value,
                    margins: updateMargins(value.margins, k, Number(e.target.value)),
                  })
                }
              />
            </label>
          ))}
        </div>
      </fieldset>

      <div className="dialog__controls page-setup-dialog__controls">
        <button type="button" onClick={onClose}>
          {t("print.pageSetup.close")}
        </button>
      </div>
    </div>
  );
}
