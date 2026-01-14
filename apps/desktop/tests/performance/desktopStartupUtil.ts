// Thin compatibility wrapper around the shared desktop startup benchmark implementation.
//
// `desktopStartupRunnerShared.ts` contains the core logic (bin discovery, Xvfb selection,
// startup log parsing, env isolation).
// `processTree.ts` provides a platform-aware process-tree termination helper.
//
// Keeping this stable `desktopStartupUtil.ts` entrypoint avoids copy/paste drift across
// standalone runners + integrated benchmarks.
export * from './desktopStartupRunnerShared.ts';
export * from './processTree.ts';
export * from './linuxProcUtil.ts';
