import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const hasBash = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('bash', ['-lc', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

const hasSh = (() => {
  if (process.platform === 'win32') return false;
  const probe = spawnSync('sh', ['-c', 'exit 0'], { stdio: 'ignore' });
  return probe.status === 0;
})();

function runBash(command) {
  const proc = spawnSync('bash', ['-lc', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return proc.stdout.trim();
}

function runSh(command) {
  const proc = spawnSync('sh', ['-c', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return { stdout: proc.stdout.trim(), stderr: proc.stderr.trim() };
}

test('agent-init warns when executed instead of sourced', { skip: !hasBash }, () => {
  const proc = spawnSync(
    'bash',
    [
      '-lc',
      [
        // Prevent agent-init from spawning Xvfb during this test.
        'export DISPLAY=:99',
        'bash scripts/agent-init.sh >/dev/null',
      ].join(' && '),
    ],
    { encoding: 'utf8', cwd: repoRoot },
  );
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  assert.match(
    proc.stderr,
    /warning: scripts\/agent-init\.sh is meant to be sourced/,
    `expected a warning on stderr; got: ${proc.stderr}`,
  );
});

test('agent-init defaults CARGO_HOME to a repo-local directory', { skip: !hasBash }, () => {
  const cargoHome = runBash(
    [
      'unset CARGO_HOME',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init treats $HOME/.cargo as default and overrides it', { skip: !hasBash }, () => {
  const cargoHome = runBash(
    [
      'unset CI',
      'export CARGO_HOME="$HOME/.cargo/"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init can preserve $HOME/.cargo when FORMULA_ALLOW_GLOBAL_CARGO_HOME=1', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CI',
      'export CARGO_HOME="$HOME/.cargo"',
      'export FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$HOME" "$CARGO_HOME"',
    ].join(' && '),
  );

  const [home, cargoHome] = out.split('\n');
  assert.equal(cargoHome, resolve(home, '.cargo'));
});

test('agent-init preserves $HOME/.cargo when running in CI', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'export CI=1',
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      'export CARGO_HOME="$HOME/.cargo"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$HOME" "$CARGO_HOME"',
    ].join(' && '),
  );

  const [home, cargoHome] = out.split('\n');
  assert.equal(cargoHome, resolve(home, '.cargo'));
});

test('agent-init preserves an existing CARGO_HOME override', { skip: !hasBash }, () => {
  const override = mkdtempSync(resolve(tmpdir(), 'formula-cargo-home-'));
  const cargoHome = runBash(
    [
      `export CARGO_HOME="${override}"`,
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, override);
});

test('agent-init prepends CARGO_HOME/bin to PATH', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CARGO_HOME',
      // Stabilize PATH so the assertion isn't affected by login shell defaults.
      'export PATH="/usr/bin:/bin"',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$CARGO_HOME" "$PATH"',
    ].join(' && '),
  );

  const [cargoHome, pathValue] = out.split('\n');
  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
  assert.ok(pathValue.split(':')[0] === resolve(repoRoot, 'target', 'cargo-home', 'bin'));
});

test('agent-init derives CARGO_BUILD_JOBS from FORMULA_CARGO_JOBS', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset CARGO_HOME',
      'unset CARGO_BUILD_JOBS',
      'unset MAKEFLAGS',
      'unset RAYON_NUM_THREADS',
      'export FORMULA_CARGO_JOBS=7',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s\\n%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS"',
    ].join(' && '),
  );

  const [jobs, makeflags, rayon] = out.split('\n');
  assert.equal(jobs, '7');
  assert.equal(makeflags, '-j7');
  assert.equal(rayon, '7');
});

test(
  'agent-init updates MAKEFLAGS and RAYON_NUM_THREADS when re-sourced after changing FORMULA_CARGO_JOBS',
  { skip: !hasBash },
  () => {
    const out = runBash(
      [
        'unset MAKEFLAGS',
        'unset RAYON_NUM_THREADS',
        'unset CARGO_BUILD_JOBS',
        'export DISPLAY=:99',
        'export FORMULA_CARGO_JOBS=7',
        'source scripts/agent-init.sh >/dev/null',
        'first=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'export FORMULA_CARGO_JOBS=3',
        'source scripts/agent-init.sh >/dev/null',
        'second=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'printf "%s\\n%s" "$first" "$second"',
      ].join(' && '),
    );

    const [first, second] = out.split('\n');
    assert.equal(first, '7,-j7,7');
    assert.equal(second, '3,-j3,3');
  },
);

test('agent-init exports FORMULA_CARGO_JOBS when set without export', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset FORMULA_CARGO_JOBS',
      // Set but do not export (common interactive usage).
      'FORMULA_CARGO_JOBS=7',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_CARGO_JOBS="',
    ].join(' && '),
  );
  assert.equal(out, 'FORMULA_CARGO_JOBS=7');
});

test('agent-init exports FORMULA_ALLOW_GLOBAL_CARGO_HOME when set without export', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      // Set but do not export (common interactive usage).
      'FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_ALLOW_GLOBAL_CARGO_HOME="',
    ].join(' && '),
  );
  assert.equal(out, 'FORMULA_ALLOW_GLOBAL_CARGO_HOME=1');
});

test('agent-init can be sourced with nounset enabled and DISPLAY unset (bash)', { skip: !hasBash }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const out = runBash(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      'source scripts/agent-init.sh >/dev/null',
      'printf "ok"',
    ].join(' && '),
  );

  assert.equal(out, 'ok');
});

test('agent-init does not enable errexit in bash when it was previously disabled', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'set +e',
      'export DISPLAY=:99',
      'before=$-',
      'source scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "%s\\n%s" "$before" "$after"',
    ].join(' && '),
  );

  const [before, after] = out.split('\n');
  assert.ok(!before.includes('e'), `expected errexit disabled before sourcing; got $-=${before}`);
  assert.ok(!after.includes('e'), `expected errexit disabled after sourcing; got $-=${after}`);
});

test('agent-init preserves errexit in bash when it was previously enabled', { skip: !hasBash }, () => {
  const out = runBash(
    [
      'set -e',
      'export DISPLAY=:99',
      'before=$-',
      'source scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "%s\\n%s" "$before" "$after"',
    ].join(' && '),
  );

  const [before, after] = out.split('\n');
  assert.ok(before.includes('e'), `expected errexit enabled before sourcing; got $-=${before}`);
  assert.ok(after.includes('e'), `expected errexit enabled after sourcing; got $-=${after}`);
});

test(
  'agent-init treats CARGO_HOME=$HOME/.cargo as unset in local runs (defaults to repo-local cargo-home)',
  { skip: !hasBash },
  () => {
    const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
    const cargoHome = runBash(
      [
        `export HOME="${fakeHome}"`,
        'export CARGO_HOME="$HOME/.cargo/"',
        'unset CI',
        'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
        'export DISPLAY=:99',
        'source scripts/agent-init.sh >/dev/null',
        'printf "%s" "$CARGO_HOME"',
      ].join(' && '),
    );

    assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
  },
);

test(
  'agent-init preserves CARGO_HOME=$HOME/.cargo when FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
  { skip: !hasBash },
  () => {
    const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
    const cargoHome = runBash(
      [
        `export HOME="${fakeHome}"`,
        'export CARGO_HOME="$HOME/.cargo"',
        'unset CI',
        'export FORMULA_ALLOW_GLOBAL_CARGO_HOME=1',
        'export DISPLAY=:99',
        'source scripts/agent-init.sh >/dev/null',
        'printf "%s" "$CARGO_HOME"',
      ].join(' && '),
    );

    assert.equal(cargoHome, resolve(fakeHome, '.cargo'));
  },
);

test('agent-init preserves CARGO_HOME=$HOME/.cargo when CI is set', { skip: !hasBash }, () => {
  const fakeHome = mkdtempSync(resolve(tmpdir(), 'formula-home-'));
  const cargoHome = runBash(
    [
      `export HOME="${fakeHome}"`,
      'export CARGO_HOME="$HOME/.cargo"',
      'export CI=1',
      'unset FORMULA_ALLOW_GLOBAL_CARGO_HOME',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(cargoHome, resolve(fakeHome, '.cargo'));
});

test('agent-init can be sourced under /bin/sh (no bash-only syntax)', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset CARGO_HOME',
      // Prevent agent-init from spawning Xvfb during this test.
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'printf "%s" "$CARGO_HOME"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, resolve(repoRoot, 'target', 'cargo-home'));
});

test('agent-init exports FORMULA_CARGO_JOBS when set without export under /bin/sh', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'unset FORMULA_CARGO_JOBS',
      'FORMULA_CARGO_JOBS=7',
      'export DISPLAY=:99',
      '. scripts/agent-init.sh >/dev/null',
      'env | grep "^FORMULA_CARGO_JOBS="',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'FORMULA_CARGO_JOBS=7');
});

test(
  'agent-init updates MAKEFLAGS and RAYON_NUM_THREADS when re-sourced after changing FORMULA_CARGO_JOBS under /bin/sh',
  { skip: !hasSh },
  () => {
    const { stdout, stderr } = runSh(
      [
        'unset MAKEFLAGS',
        'unset RAYON_NUM_THREADS',
        'unset CARGO_BUILD_JOBS',
        'export DISPLAY=:99',
        'FORMULA_CARGO_JOBS=7',
        '. scripts/agent-init.sh >/dev/null',
        'first=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'FORMULA_CARGO_JOBS=3',
        '. scripts/agent-init.sh >/dev/null',
        'second=$(printf "%s,%s,%s" "$CARGO_BUILD_JOBS" "$MAKEFLAGS" "$RAYON_NUM_THREADS")',
        'printf "%s\\n%s" "$first" "$second"',
      ].join(' && '),
    );

    assert.equal(stderr, '');
    const [first, second] = stdout.split('\n');
    assert.equal(first, '7,-j7,7');
    assert.equal(second, '3,-j3,3');
  },
);

test('agent-init can be sourced with nounset enabled and DISPLAY unset under /bin/sh', { skip: !hasSh }, () => {
  const stubDir = mkdtempSync(resolve(tmpdir(), 'formula-xvfb-stub-'));
  const stubPath = resolve(stubDir, 'Xvfb');
  writeFileSync(stubPath, '#!/bin/sh\nexit 0\n', { mode: 0o755 });

  const { stdout, stderr } = runSh(
    [
      `export PATH="${stubDir}:$PATH"`,
      'set -u',
      'unset DISPLAY',
      '. scripts/agent-init.sh >/dev/null',
      'printf "ok"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  assert.equal(stdout, 'ok');
});

test('agent-init does not enable errexit in /bin/sh when it was previously disabled', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'set +e',
      'export DISPLAY=:99',
      'before=$-',
      '. scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "before=%s\\nafter=%s" "$before" "$after"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  const [beforeLine, afterLine] = stdout.split('\n');
  const before = beforeLine.replace(/^before=/, '');
  const after = afterLine.replace(/^after=/, '');
  assert.ok(!before.includes('e'), `expected errexit disabled before sourcing; got $-=${before}`);
  assert.ok(!after.includes('e'), `expected errexit disabled after sourcing; got $-=${after}`);
});

test('agent-init preserves errexit in /bin/sh when it was previously enabled', { skip: !hasSh }, () => {
  const { stdout, stderr } = runSh(
    [
      'set -e',
      'export DISPLAY=:99',
      'before=$-',
      '. scripts/agent-init.sh >/dev/null',
      'after=$-',
      'printf "before=%s\\nafter=%s" "$before" "$after"',
    ].join(' && '),
  );

  assert.equal(stderr, '');
  const [beforeLine, afterLine] = stdout.split('\n');
  const before = beforeLine.replace(/^before=/, '');
  const after = afterLine.replace(/^after=/, '');
  assert.ok(before.includes('e'), `expected errexit enabled before sourcing; got $-=${before}`);
  assert.ok(after.includes('e'), `expected errexit enabled after sourcing; got $-=${after}`);
});
