import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { ToolExecutor } from "../../../../../packages/ai-tools/src/index.js";

import { ChartStore } from "../../charts/chartStore";
import { DocumentControllerSpreadsheetApi } from "./documentControllerSpreadsheetApi.js";

describe("create_chart desktop integration", () => {
  it("creates a chart record via ToolExecutor + DocumentControllerSpreadsheetApi", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", "Category");
    controller.setCellValue("Sheet1", "B1", "Value");
    controller.setCellValue("Sheet1", "A2", "A");
    controller.setCellValue("Sheet1", "B2", 10);
    controller.setCellValue("Sheet1", "A3", "B");
    controller.setCellValue("Sheet1", "B3", 20);

    const chartStore = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: (sheetId, row, col) => {
        const cell = controller.getCell(sheetId, { row, col }) as { value: unknown } | null;
        return cell?.value ?? null;
      }
    });

    const api = new DocumentControllerSpreadsheetApi(controller, {
      createChart: chartStore.createChart.bind(chartStore)
    });
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "create_chart",
      parameters: {
        chart_type: "bar",
        data_range: "A1:B3",
        title: "Sales"
      }
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("create_chart");
    if (!result.ok || result.tool !== "create_chart") throw new Error("Unexpected tool result");
    expect(result.data?.status).toBe("ok");
    expect(result.data?.chart_id).toBe("chart_1");

    const charts = chartStore.listCharts();
    expect(charts).toHaveLength(1);
    expect(charts[0]?.title).toBe("Sales");
    expect(charts[0]?.series[0]).toMatchObject({
      name: "Value",
      categories: "Sheet1!$A$2:$A$3",
      values: "Sheet1!$B$2:$B$3"
    });
  });
});

