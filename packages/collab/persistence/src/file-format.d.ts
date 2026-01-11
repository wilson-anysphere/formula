import type { KeyRing } from "../../../security/crypto/keyring.js";

export const FILE_MAGIC: Buffer;
export const FILE_HEADER_BYTES: number;
export const FILE_FLAG_ENCRYPTED: number;

export function encodeLegacyRecord(update: Uint8Array): Buffer;
export function decodeLegacyRecords(data: Buffer, offset?: number): Uint8Array[];

export function hasFileHeader(data: Buffer): boolean;
export function parseFileHeader(data: Buffer): { flags: number };
export function encodeFileHeader(flags: number): Buffer;

export function encodeEncryptedRecord(
  update: Uint8Array,
  opts: { keyRing: KeyRing; aadContext: unknown }
): Buffer;
export function decodeEncryptedRecords(
  data: Buffer,
  opts: { keyRing: KeyRing; aadContext: unknown },
  offset?: number
): Uint8Array[];

export function atomicWriteFile(filePath: string, contents: Buffer): Promise<void>;

