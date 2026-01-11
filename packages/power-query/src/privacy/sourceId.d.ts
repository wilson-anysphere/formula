export function normalizeFilePath(input: string): string;

export function getFileSourceId(path: string): string;

export function getHttpSourceId(url: string): string;

export function getSqlSourceId(connection: unknown): string;

export function getSourceIdForQuerySource(source: unknown): string | null;

export function getSourceIdForProvenance(provenance: unknown): string | null;
