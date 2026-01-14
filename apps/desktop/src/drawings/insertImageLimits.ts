/**
 * Maximum allowed image size when inserting pictures from disk.
 *
 * This limit exists to avoid unbounded memory usage when reading local files into
 * the webview process (both in browser builds and when using Tauri's invoke APIs).
 *
 * Must be <= the Tauri backend IPC read limits (`MAX_READ_FULL_BYTES`) but is
 * intentionally smaller than that backend cap.
 */
// Keep this in sync with the clipboard image size guards so we don't allow users to
// insert images that the desktop IPC boundary can't safely transport.
export const MAX_INSERT_IMAGE_BYTES = 5 * 1024 * 1024; // 5MiB
