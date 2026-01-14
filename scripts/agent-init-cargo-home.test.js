import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync } from 'node:fs';
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
