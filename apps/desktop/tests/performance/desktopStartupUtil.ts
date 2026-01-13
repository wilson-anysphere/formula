// Thin compatibility wrapper around the shared desktop startup benchmark implementation.
//
// `desktopStartupRunnerShared.ts` contains the actual logic (bin discovery, Xvfb selection,
// startup log parsing, process-tree termination, and environment isolation). Keeping this
// re-export allows callers to depend on a stable "util" module name.
export * from './desktopStartupRunnerShared.ts';
