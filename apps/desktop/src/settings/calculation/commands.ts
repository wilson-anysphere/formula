export const CALCULATE_NOW_COMMAND_ID = "workbook.calculateNow";

export interface Command {
  id: string;
  title: string;
}

export const calculateNowCommand: Command = {
  id: CALCULATE_NOW_COMMAND_ID,
  title: "Calculate Now",
};

