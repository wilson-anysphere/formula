// Thin compatibility wrapper around the shared desktop startup benchmark implementation.
//
// `desktopStartupRunnerShared.ts` contains the core logic (bin discovery, Xvfb selection,
// startup log parsing, env isolation, and env parsing helpers.
// `processTree.ts` provides a platform-aware process-tree termination helper.
// `linuxProcUtil.ts` provides Linux `/proc` helpers used for RSS sampling in perf benchmarks.
//
// Keeping this stable `desktopStartupUtil.ts` entrypoint avoids copy/paste drift across
// standalone runners + integrated benchmarks.
export * from './desktopStartupRunnerShared.ts';
export * from './processTree.ts';
export * from './linuxProcUtil.ts';
