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

test(
  'cargo_agent rejects invalid FORMULA_LLD_THREADS',
  { skip: !hasBash },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        // Should fail fast before invoking cargo.
        'export FORMULA_LLD_THREADS=not-a-number && bash scripts/cargo_agent.sh check -h',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_LLD_THREADS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_LLD_THREADS'),
      `expected stderr to mention invalid FORMULA_LLD_THREADS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_LLD_THREADS even when --target is set',
  { skip: !hasBash },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        // Use `-h` so cargo prints help and does not attempt a real build for the target.
        'export FORMULA_LLD_THREADS=not-a-number && bash scripts/cargo_agent.sh check -h --target wasm32-unknown-unknown',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_LLD_THREADS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_LLD_THREADS'),
      `expected stderr to mention invalid FORMULA_LLD_THREADS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent clamps default RUST_TEST_THREADS based on chosen test jobs',
  { skip: !hasBash || !hasCargo },
  () => {
    const { stderr } = runBash(
      'unset FORMULA_CARGO_JOBS FORMULA_CARGO_TEST_JOBS FORMULA_RUST_TEST_THREADS RUST_TEST_THREADS && bash scripts/cargo_agent.sh test -h',
    );
    const match = stderr.match(/test_threads=([0-9]+)/);
    assert.ok(match, stderr);
    const threads = Number(match[1]);
    // With no explicit configuration, cargo_agent defaults `cargo test` jobs to 1. The wrapper
    // should therefore clamp the default test thread pool to <= 4 (jobs * 4).
    assert.ok(stderr.includes('jobs=1'), stderr);
    assert.ok(
      threads <= 4,
      `expected test_threads <= 4 when jobs=1, got ${threads}\nstderr:\n${stderr}`,
    );
  },
);

test(
  'cargo_agent aligns default RAYON_NUM_THREADS with test jobs when it was tracking jobs',
  { skip: !hasBash || !hasCargo },
  () => {
    // Simulate a common environment setup where Rayon is preconfigured to match build jobs.
    // The wrapper should still be able to reduce Rayon threads when it reduces `cargo test` jobs.
    const { stderr } = runBash(
      'export RAYON_NUM_THREADS=4 && unset FORMULA_RAYON_NUM_THREADS FORMULA_CARGO_JOBS FORMULA_CARGO_TEST_JOBS RUST_TEST_THREADS FORMULA_RUST_TEST_THREADS && bash scripts/cargo_agent.sh test -h',
    );
    assert.ok(stderr.includes('jobs=1'), stderr);
    assert.ok(stderr.includes('rayon=1'), stderr);
  },
);

test(
  'cargo_agent rejects invalid FORMULA_RAYON_NUM_THREADS',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        'unset RAYON_NUM_THREADS && export FORMULA_RAYON_NUM_THREADS=not-a-number && bash scripts/cargo_agent.sh check -h',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_RAYON_NUM_THREADS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_RAYON_NUM_THREADS'),
      `expected stderr to mention invalid FORMULA_RAYON_NUM_THREADS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_RUST_TEST_THREADS',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        'unset RUST_TEST_THREADS FORMULA_RUST_TEST_THREADS && export FORMULA_RUST_TEST_THREADS=0 && bash scripts/cargo_agent.sh test -h',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_RUST_TEST_THREADS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_RUST_TEST_THREADS'),
      `expected stderr to mention invalid FORMULA_RUST_TEST_THREADS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_CARGO_TEST_JOBS',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        'export FORMULA_CARGO_TEST_JOBS=0 && unset FORMULA_CARGO_JOBS && bash scripts/cargo_agent.sh test -h',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_CARGO_TEST_JOBS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_CARGO_TEST_JOBS'),
      `expected stderr to mention invalid FORMULA_CARGO_TEST_JOBS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_CARGO_RETRY_ATTEMPTS',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      [
        '-lc',
        'export FORMULA_CARGO_RETRY_ATTEMPTS=0 && bash scripts/cargo_agent.sh check -h',
      ],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_CARGO_RETRY_ATTEMPTS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_CARGO_RETRY_ATTEMPTS'),
      `expected stderr to mention invalid FORMULA_CARGO_RETRY_ATTEMPTS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_CARGO_JOBS',
  { skip: !hasBash || !hasCargo },
  () => {
    const proc = spawnSync(
      'bash',
      ['-lc', 'export FORMULA_CARGO_JOBS=not-a-number && bash scripts/cargo_agent.sh check -h'],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_CARGO_JOBS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_CARGO_JOBS'),
      `expected stderr to mention invalid FORMULA_CARGO_JOBS, got:\n${proc.stderr}`,
    );
  },
);

test(
  'cargo_agent rejects invalid FORMULA_CARGO_LIMIT_AS',
  { skip: !hasBash },
  () => {
    const proc = spawnSync(
      'bash',
      ['-lc', 'export FORMULA_CARGO_LIMIT_AS=not-a-size && bash scripts/cargo_agent.sh check -h'],
      { encoding: 'utf8', cwd: repoRoot },
    );
    if (proc.error) throw proc.error;
    assert.notEqual(proc.status, 0, 'expected non-zero exit for invalid FORMULA_CARGO_LIMIT_AS');
    assert.ok(
      proc.stderr.includes('invalid FORMULA_CARGO_LIMIT_AS'),
      `expected stderr to mention invalid FORMULA_CARGO_LIMIT_AS, got:\n${proc.stderr}`,
    );
  },
);
