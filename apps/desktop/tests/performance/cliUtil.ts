let installedEpipeHandler = false;

export function installEpipeHandler(): void {
  if (installedEpipeHandler) return;
  installedEpipeHandler = true;
  // When piping output to tools like `head`, the consumer can close the pipe early. Node treats
  // subsequent writes to stdout as an EPIPE error, which can crash the process with a noisy stack
  // trace if unhandled.
  //
  // Gracefully exit on EPIPE so CLI runner scripts behave like typical Unix tools.
  process.stdout.on('error', (err) => {
    const code = (err as NodeJS.ErrnoException | null)?.code;
    if (code === 'EPIPE') process.exit(0);
  });
}
