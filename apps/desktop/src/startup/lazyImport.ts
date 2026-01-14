export type LazyImportErrorHandler = (error: unknown) => void;

export type CreateLazyImportOptions = {
  /**
   * Human-friendly label for log/toast messages.
   */
  label: string;
  /**
   * Optional error handler invoked when the import fails.
   */
  onError?: LazyImportErrorHandler;
};

/**
 * Wrap a dynamic import() in a small cache so repeated calls share the same in-flight
 * promise and loaded module.
 *
 * If the import rejects, the cache is cleared so callers can retry (useful for
 * transient failures like missing chunks during updates).
 */
export function createLazyImport<TModule>(
  importer: () => Promise<TModule>,
  options: CreateLazyImportOptions,
): () => Promise<TModule | null> {
  let promise: Promise<TModule> | null = null;

  return async () => {
    if (!promise) {
      promise = importer();
    }

    try {
      return await promise;
    } catch (err) {
      promise = null;
      try {
        options.onError?.(err);
      } catch {
        // ignore error handler failures
      }
      return null;
    }
  };
}

