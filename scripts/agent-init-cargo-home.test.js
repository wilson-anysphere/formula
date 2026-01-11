import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function runBash(command) {
  const proc = spawnSync('bash', ['-lc', command], {
    encoding: 'utf8',
    cwd: repoRoot,
  });
  if (proc.error) throw proc.error;
  assert.equal(proc.status, 0, proc.stderr);
  return proc.stdout.trim();
}

test('agent-init defaults CARGO_HOME to a repo-local directory', () => {
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

test('agent-init preserves an existing CARGO_HOME override', () => {
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

test('agent-init prepends CARGO_HOME/bin to PATH', () => {
  const out = runBash(
    [
      'unset CARGO_HOME',
      'export DISPLAY=:99',
      'source scripts/agent-init.sh >/dev/null',
      'printf "%s\\n%s" "$CARGO_HOME" "$PATH"',
    ].join(' && '),
  );

  const [cargoHome, pathValue] = out.split('\n');
  assert.equal(cargoHome, resolve(repoRoot, 'target', 'cargo-home'));
  assert.ok(pathValue.split(':')[0] === resolve(repoRoot, 'target', 'cargo-home', 'bin'));
});
