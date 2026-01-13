import type { VectorSearchResult } from "../store/inMemoryVectorStore.js";

export function rerankWorkbookResults(params: {
  queryText: string;
  results: VectorSearchResult[];
}): VectorSearchResult[];

export function dedupeOverlappingResults(params: {
  results: VectorSearchResult[];
  overlapRatio?: number;
}): VectorSearchResult[];

