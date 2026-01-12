import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const hasBash = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('bash', ['-lc', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

const hasCargo = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('cargo', ['--version'], { stdio: 'ignore' });
  if (probe.error) return false;
  return probe.status === 0;
})();

const hasCargoFmt = (() => {
  if (!hasCargo) return false;
  const probe = spawnSync('cargo', ['fmt', '--', '--version'], { stdio: 'ignore', cwd: repoRoot });
  if (probe.error) return false;
  return probe.status === 0;
})();

function runBash(command) {
  const proc = spawnSync('bash', ['-lc', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return { stdout: proc.stdout, stderr: proc.stderr };
}

test(
  'cargo_agent can run cargo fmt (subcommand does not accept -j)',
  { skip: !hasBash || !hasCargoFmt },
  () => {
  const { stdout, stderr } = runBash('bash scripts/cargo_agent.sh fmt -- --version');
  assert.ok(!stderr.includes("unexpected argument '-j'"), stderr);
  assert.match(stdout, /rustfmt/i);
  },
);

test(
  'cargo_agent can run cargo clean (subcommand does not accept -j)',
  { skip: !hasBash || !hasCargo },
  () => {
  const { stderr } = runBash('bash scripts/cargo_agent.sh clean -n');
  assert.ok(!stderr.includes("unexpected argument '-j'"), stderr);
  },
);

test(
  'cargo_agent uses CARGO_BUILD_JOBS when FORMULA_CARGO_JOBS is unset',
  { skip: !hasBash || !hasCargo },
  () => {
    const { stderr } = runBash(
      'unset FORMULA_CARGO_JOBS && export CARGO_BUILD_JOBS=7 && bash scripts/cargo_agent.sh check -h',
    );
    assert.ok(stderr.includes('jobs=7'), stderr);
  },
);
