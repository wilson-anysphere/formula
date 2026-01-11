export type DataType = string;

export type QuerySource = { type: string; [key: string]: any };

export type QueryOperation = { type: string; [key: string]: any };

export type QueryStep = {
  id: string;
  name: string;
  operation: QueryOperation;
};

export type RefreshPolicy = { type: string; [key: string]: any };

export type Query = {
  id: string;
  name: string;
  source: QuerySource;
  steps: QueryStep[];
  refreshPolicy?: RefreshPolicy;
  destination?: unknown;
  [key: string]: any;
};
