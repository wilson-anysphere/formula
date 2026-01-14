/**
 * Maximum allowed image size when inserting pictures from disk (file picker / drag-drop).
 *
 * This limit exists to avoid unbounded memory usage when reading local files into the
 * webview process (both in browser builds and when using Tauri's invoke APIs).
 *
 * Note: This is intentionally higher than the clipboard image guard (5MiB) because:
 * - file insertion is explicit user intent
 * - clipboard payloads are often raw PNG bytes and can balloon quickly
 *
 * Must be <= the Tauri backend IPC read limits (`MAX_READ_FULL_BYTES`) and should remain
 * comfortably below it to avoid transport overhead.
 */
export const MAX_INSERT_IMAGE_BYTES = 10 * 1024 * 1024; // 10MiB
