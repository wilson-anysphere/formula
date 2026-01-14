// Thin compatibility wrapper around the shared desktop startup benchmark implementation.
//
// `desktopStartupRunnerShared.ts` contains the core logic (bin discovery, Xvfb selection,
// startup log parsing, env isolation, and env parsing helpers.
// `processTree.ts` provides a platform-aware process-tree termination helper.
// `linuxProcUtil.ts` provides Linux `/proc` helpers used for RSS sampling in perf benchmarks.
// `sleep.ts` provides a shared AbortSignal-aware sleep helper used across perf scripts.
// `desktopMemoryUtil.ts` provides shared env parsing for the idle memory benchmark.
//
// Keeping this stable `desktopStartupUtil.ts` entrypoint avoids copy/paste drift across
// standalone runners + integrated benchmarks.
export * from './desktopStartupRunnerShared.ts';
export * from './processTree.ts';
export * from './linuxProcUtil.ts';
export * from './sleep.ts';
export * from './desktopMemoryUtil.ts';
export * from './benchmark.ts';
