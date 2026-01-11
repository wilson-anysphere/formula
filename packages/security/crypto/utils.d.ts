/// <reference types="node" />

export function assertBufferLength(buf: Buffer, expected: number, name: string): void;
export function toBase64(buf: Buffer): string;
export function fromBase64(value: string, name?: string): Buffer;
export function canonicalJson(value: unknown): string;
export function aadFromContext(context: unknown | null | undefined): Buffer | null;
export function randomId(bytes?: number): string;
