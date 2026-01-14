#!/usr/bin/env node
import process from "node:process";

const EXPECTED_CI_WORKFLOW_PATH = ".github/workflows/ci.yml";
const DEFAULT_WORKFLOW_NAME = "CI";
const DEFAULT_PER_PAGE = 100;
const GITHUB_API_BASE_URL = (process.env.GITHUB_API_URL ?? "https://api.github.com").replace(/\/$/, "");
const GITHUB_SERVER_BASE_URL = (process.env.GITHUB_SERVER_URL ?? "https://github.com").replace(/\/$/, "");

function usage() {
  console.log(`Usage: node scripts/check-tag-ci-status.mjs [options]

Options:
  --sha <commit>        Commit SHA to check (default: GITHUB_SHA)
  --repo <owner/name>   GitHub repository (default: GITHUB_REPOSITORY)
  --workflow <name>     Workflow name (default: "${DEFAULT_WORKFLOW_NAME}")
  -h, --help            Show help

Environment:
  GITHUB_TOKEN          GitHub token with actions:read permission (preferred)
  GH_TOKEN              Alternative token env var (optional)
  GITHUB_API_URL        GitHub API base URL (default: https://api.github.com)
  GITHUB_SERVER_URL     GitHub server base URL for links (default: https://github.com)
`);
}

/**
 * @param {string} message
 */
function fatal(message) {
  console.error(message);
  process.exit(1);
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {{ sha?: string; repo?: string; workflow?: string; help?: boolean }} */
  const out = {};

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];

    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }

    const eqIndex = arg.indexOf("=");
    const key = eqIndex === -1 ? arg : arg.slice(0, eqIndex);
    const hasInlineValue = eqIndex !== -1;
    const value = hasInlineValue ? arg.slice(eqIndex + 1) : argv[i + 1];

    switch (key) {
      case "--sha":
        if (!value) fatal(`Missing value for ${key}.`);
        out.sha = value;
        if (!hasInlineValue) i++;
        break;
      case "--repo":
        if (!value) fatal(`Missing value for ${key}.`);
        out.repo = value;
        if (!hasInlineValue) i++;
        break;
      case "--workflow":
        if (!value) fatal(`Missing value for ${key}.`);
        out.workflow = value;
        if (!hasInlineValue) i++;
        break;
      default:
        fatal(`Unknown argument: ${arg}\n\nRun with --help for usage.`);
    }
  }

  return out;
}

/**
 * @param {string} repo
 */
function parseRepo(repo) {
  const parts = repo.split("/");
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    fatal(
      `Invalid --repo value "${repo}". Expected the form "owner/name" (example: "acme/widgets").`,
    );
  }
  return { owner: parts[0], repo: parts[1] };
}

/**
 * @param {string} message
 * @param {unknown} details
 */
function formatGitHubError(message, details) {
  if (typeof details === "string") return `${message}\n\nResponse:\n${details}`;
  if (details && typeof details === "object") return `${message}\n\nResponse:\n${JSON.stringify(details, null, 2)}`;
  return message;
}

/**
 * @param {string} url
 * @param {{ token: string }} opts
 */
async function githubGetJson(url, opts) {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${opts.token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "formula-release-preflight",
    },
  });

  const text = await res.text();
  /** @type {any} */
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = text;
  }

  if (!res.ok) {
    const extraHelp =
      res.status === 403
        ? `\n\nHint: if this is running in GitHub Actions, ensure the job grants GITHUB_TOKEN \"actions: read\" permission.`
        : "";
    fatal(
      formatGitHubError(
        `GitHub API request failed (${res.status} ${res.statusText}) for ${url}.${extraHelp}`,
        json,
      ),
    );
  }

  return json;
}

/**
 * Resolve a workflow by name while deliberately restricting to the project's main CI workflow file.
 *
 * @param {{ owner: string; repo: string; workflowName: string; token: string }}
 */
async function resolveCiWorkflow({ owner, repo, workflowName, token }) {
  /** @type {Array<{ id: number; name: string; path: string }>} */
  const workflows = [];

  for (let page = 1; page <= 20; page++) {
    const url = `${GITHUB_API_BASE_URL}/repos/${owner}/${repo}/actions/workflows?per_page=${DEFAULT_PER_PAGE}&page=${page}`;
    /** @type {{ workflows?: Array<{ id: number; name: string; path: string }> }} */
    const data = await githubGetJson(url, { token });
    const batch = Array.isArray(data?.workflows) ? data.workflows : [];
    workflows.push(...batch);
    if (batch.length < DEFAULT_PER_PAGE) break;
  }

  const expected = workflows.filter(
    (w) => w?.name === workflowName && w?.path === EXPECTED_CI_WORKFLOW_PATH,
  );
  if (expected.length === 1) return expected[0];

  const sameName = workflows.filter((w) => w?.name === workflowName);
  const ciPath = workflows.find((w) => w?.path === EXPECTED_CI_WORKFLOW_PATH);

  if (expected.length > 1) {
    fatal(
      `Release preflight failed: multiple workflows matched name "${workflowName}" at ${EXPECTED_CI_WORKFLOW_PATH}. This is unexpected.`,
    );
  }

  if (sameName.length > 0) {
    fatal(
      [
        `Release preflight failed: found workflow(s) named "${workflowName}", but none are the main CI workflow file (${EXPECTED_CI_WORKFLOW_PATH}).`,
        ``,
        `Matches:`,
        ...sameName.map((w) => `  - ${w.name} (${w.path})`),
        ``,
        `This check is intentionally restricted to ${EXPECTED_CI_WORKFLOW_PATH} to avoid blocking releases on optional workflows.`,
      ].join("\n"),
    );
  }

  if (ciPath) {
    fatal(
      [
        `Release preflight failed: expected the main CI workflow (${EXPECTED_CI_WORKFLOW_PATH}) to be named "${workflowName}",`,
        `but found "${ciPath.name}".`,
        ``,
        `Fix: update the release workflow to pass --workflow ${JSON.stringify(ciPath.name)} (or rename the workflow back to "${workflowName}").`,
      ].join("\n"),
    );
  }

  fatal(
    [
      `Release preflight failed: could not find the main CI workflow at ${EXPECTED_CI_WORKFLOW_PATH}.`,
      `Available workflows:`,
      ...workflows.map((w) => `  - ${w.name} (${w.path})`),
    ].join("\n"),
  );
}

/**
 * @param {{
 *   owner: string;
 *   repo: string;
 *   workflowId: number;
 *   sha: string;
 *   token: string;
 * }}
 */
async function findNewestSuccessfulRun({ owner, repo, workflowId, sha, token }) {
  /** @type {any | null} */
  let newestRun = null;
  /** @type {any | null} */
  let newestCompletedRun = null;

  for (let page = 1; page <= 20; page++) {
    const url =
      `${GITHUB_API_BASE_URL}/repos/${owner}/${repo}/actions/workflows/${workflowId}/runs` +
      `?head_sha=${encodeURIComponent(sha)}` +
      `&per_page=${DEFAULT_PER_PAGE}` +
      `&page=${page}`;

    /** @type {{ workflow_runs?: Array<any> }} */
    const data = await githubGetJson(url, { token });
    const runs = Array.isArray(data?.workflow_runs) ? data.workflow_runs : [];

    for (const run of runs) {
      // Defensive: even though we query with head_sha, validate it to avoid passing
      // the release preflight if GitHub ignores the query parameter (or changes API behavior).
      if (run?.head_sha !== sha) continue;

      if (!newestRun) newestRun = run;
      if (!newestCompletedRun && run?.status === "completed") newestCompletedRun = run;

      if (run?.status === "completed" && run?.conclusion === "success") {
        return { run, newestRun, newestCompletedRun };
      }
    }

    if (runs.length < DEFAULT_PER_PAGE) break;
  }

  return { run: null, newestRun, newestCompletedRun };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    usage();
    return;
  }

  const sha = args.sha ?? process.env.GITHUB_SHA;
  const repo = args.repo ?? process.env.GITHUB_REPOSITORY;
  const workflowName = args.workflow ?? DEFAULT_WORKFLOW_NAME;

  if (!sha) {
    fatal(`Missing commit SHA. Pass --sha <commit> or set GITHUB_SHA.`);
  }
  if (!repo) {
    fatal(`Missing repository. Pass --repo <owner/name> or set GITHUB_REPOSITORY.`);
  }

  const token = process.env.GITHUB_TOKEN ?? process.env.GH_TOKEN;
  if (!token) {
    fatal(
      [
        `Missing GitHub token.`,
        `Set the GITHUB_TOKEN (preferred) or GH_TOKEN environment variable to a GitHub token with "actions:read" permission.`,
        `In GitHub Actions, you can typically use: env: { GITHUB_TOKEN: \${{ secrets.GITHUB_TOKEN }} }`,
      ].join("\n"),
    );
  }

  const { owner, repo: repoName } = parseRepo(repo);

  const workflow = await resolveCiWorkflow({ owner, repo: repoName, workflowName, token });

  const { run: successfulRun, newestRun, newestCompletedRun } = await findNewestSuccessfulRun({
    owner,
    repo: repoName,
    workflowId: workflow.id,
    sha,
    token,
  });

  if (!successfulRun) {
    const header = `Release preflight failed: no successful "${workflowName}" workflow run found for commit ${sha}.`;
    const workflowUrl = `${GITHUB_SERVER_BASE_URL}/${owner}/${repoName}/actions/workflows/${workflow.path}`;

    if (!newestRun) {
      fatal(
        [
          header,
          ``,
          `GitHub Actions shows no workflow runs for this commit (workflow: ${workflow.path}).`,
          `Make sure CI has finished successfully for the commit you tagged, then re-run the release.`,
          `Workflow: ${workflowUrl}`,
        ].join("\n"),
      );
    }

    const runUrl = newestRun?.html_url;
    const status = newestRun?.status ?? "unknown";

    if (status !== "completed") {
      const previousConclusion =
        newestCompletedRun && newestCompletedRun !== newestRun
          ? ` (previous completed run: conclusion=${newestCompletedRun?.conclusion ?? "unknown"}${
              newestCompletedRun?.html_url ? ` (${newestCompletedRun.html_url})` : ""
            })`
          : "";

      fatal(
        [
          header,
          ``,
          `Newest run: status=${status}${runUrl ? ` (${runUrl})` : ""}${previousConclusion}`,
          `Fix: wait for CI to finish and pass on this commit, then re-tag / re-run the release workflow.`,
          `Workflow: ${workflowUrl}`,
        ].join("\n"),
      );
    }

    const conclusion = newestRun?.conclusion ?? "unknown";
    fatal(
      [
        header,
        ``,
        `Newest completed run: conclusion=${conclusion}${runUrl ? ` (${runUrl})` : ""}`,
        `Fix: wait for CI to pass on this commit, then re-tag / re-run the release workflow.`,
        `Workflow: ${workflowUrl}`,
      ].join("\n"),
    );
  }

  console.log(
    `CI status check passed: ${workflowName} succeeded for ${sha} (${successfulRun.html_url ?? "run URL unavailable"}).`,
  );
}

main().catch((err) => {
  const msg = err instanceof Error ? err.stack ?? err.message : String(err);
  fatal(`Unexpected error while checking CI status.\n${msg}`);
});
