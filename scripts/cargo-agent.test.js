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

test(
  'cargo_agent clamps very high CARGO_BUILD_JOBS unless FORMULA_CARGO_JOBS is set',
  { skip: !hasBash || !hasCargo },
  () => {
    {
      const { stderr } = runBash(
        'unset FORMULA_CARGO_JOBS && export CARGO_BUILD_JOBS=99 && bash scripts/cargo_agent.sh check -h',
      );
      assert.ok(stderr.includes('clamping to 8'), stderr);
      assert.ok(stderr.includes('jobs=8'), stderr);
    }

    {
      const { stderr } = runBash(
        'export FORMULA_CARGO_JOBS=99 && export CARGO_BUILD_JOBS=99 && bash scripts/cargo_agent.sh check -h',
      );
      assert.ok(!stderr.includes('clamping to 8'), stderr);
      assert.ok(stderr.includes('jobs=99'), stderr);
    }
  },
);

test(
  'cargo_agent preserves stdout (cargo metadata JSON not corrupted by stderr logs)',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        // Force cargo to emit logs on stderr while still writing JSON to stdout.
        'export CARGO_LOG=trace && bash scripts/cargo_agent.sh metadata --format-version 1 --no-deps',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.equal(proc.status, 0, proc.stderr);

    // Ensure cargo actually emitted something to stderr so this test would have caught a
    // regression to the old `2>&1 | tee ...` behavior.
    assert.ok(proc.stderr.includes('TRACE cargo:'), proc.stderr);

    const stdout = String(proc.stdout).trim();
    assert.ok(stdout.startsWith('{'), stdout.slice(0, 200));
    assert.doesNotThrow(() => JSON.parse(stdout));
  },
);
