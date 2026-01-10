import type { CalcSettings } from "./types";

export type SettingField =
  | {
      id: string;
      label: string;
      kind: "toggle";
      value: boolean;
    }
  | {
      id: string;
      label: string;
      kind: "number";
      value: number;
      min?: number;
      step?: number;
      disabled?: boolean;
    }
  | {
      id: string;
      label: string;
      kind: "select";
      value: string;
      options: { value: string; label: string }[];
    };

export interface SettingsSection {
  id: string;
  title: string;
  fields: SettingField[];
}

/**
 * Declarative schema for a calculation settings panel.
 *
 * The desktop app can render this schema with its existing settings UI system.
 */
export function calculationSettingsSchema(settings: CalcSettings): SettingsSection[] {
  return [
    {
      id: "calculation.mode",
      title: "Calculation",
      fields: [
        {
          id: "calculationMode",
          label: "Workbook calculation",
          kind: "select",
          value: settings.calculationMode,
          options: [
            { value: "automatic", label: "Automatic" },
            { value: "manual", label: "Manual" },
          ],
        },
        {
          id: "calculateBeforeSave",
          label: "Calculate workbook before saving",
          kind: "toggle",
          value: settings.calculateBeforeSave,
        },
        {
          id: "fullPrecision",
          label: "Full precision (disable 'precision as displayed')",
          kind: "toggle",
          value: settings.fullPrecision,
        },
      ],
    },
    {
      id: "calculation.iterative",
      title: "Iterative Calculation (Circular References)",
      fields: [
        {
          id: "iterative.enabled",
          label: "Enable iterative calculation",
          kind: "toggle",
          value: settings.iterative.enabled,
        },
        {
          id: "iterative.maxIterations",
          label: "Maximum iterations",
          kind: "number",
          value: settings.iterative.maxIterations,
          min: 1,
          step: 1,
          disabled: !settings.iterative.enabled,
        },
        {
          id: "iterative.maxChange",
          label: "Maximum change",
          kind: "number",
          value: settings.iterative.maxChange,
          min: 0,
          step: 0.0001,
          disabled: !settings.iterative.enabled,
        },
      ],
    },
  ];
}

