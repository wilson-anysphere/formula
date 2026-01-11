export type InferredValueType = "empty" | "number" | "boolean" | "datetime" | "string";

export type ParsedScalar = { value: string | number | boolean | null; type: InferredValueType; numberFormat?: string };

export function inferValueType(rawInput: string): InferredValueType;

export function parseScalar(rawInput: string): ParsedScalar;

export function parseScalarValue(rawInput: string): ParsedScalar["value"];

export function isNumberString(rawInput: string): boolean;

export function isBooleanString(rawInput: string): boolean;

export function isIsoLikeDateString(rawInput: string): boolean;

export function dateToExcelSerial(dateUtc: Date): number;

export function excelSerialToDate(serial: number): Date;

export function parseIsoLikeToUtcDate(input: string): Date | null;
