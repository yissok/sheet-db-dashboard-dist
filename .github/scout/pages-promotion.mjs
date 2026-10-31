import { scoutRepo } from './config.mjs';
import { appendFile, readFile } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { pathToFileURL } from 'node:url';

export function githubApi(token, fetcher = fetch) {
  return async (path, { method = 'GET', body } = {}) => {
    const response = await fetcher(`https://api.github.com/${path}`, {
      method,
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json', 'X-GitHub-Api-Version': '2022-11-28' },
      ...(body ? { body: JSON.stringify(body) } : {}),
      signal: AbortSignal.timeout(30000),
    });
    if (response.status === 404 && method === 'GET') return null;
    if (!response.ok) throw new Error(`GitHub ${method} ${path}: HTTP ${response.status}`);
    return response.status === 204 ? null : response.json();
  };
}

export async function repositoryJson(api, repository, path, ref) {
  const file = await api(`repos/${repository}/contents/${path}?ref=${encodeURIComponent(ref)}`);
  if (!file) return null;
  return JSON.parse(Buffer.from(file.content, 'base64').toString());
}

// All deployments in one distribution repository share the scout-pages lock.
export async function promotePages({ repository, sha, manifest, api, dispatch, verify }) {
  if (!/^[\w.-]+\/[\w.-]+-dist$/.test(repository) || !/^[a-f0-9]{40}$/.test(sha)
      || manifest.version !== 4 || `${manifest.app}-dist` !== repository.split('/')[1]) {
    throw new Error('Invalid Pages promotion context.');
  }
  const root = () => repositoryJson(api, repository, 'manifest.json', 'main');
  if ((await root())?.assets !== manifest.assets) return 'superseded';
  const ready = await api(`repos/${repository}/git/ref/heads/scout-ready`);
  const previous = ready && await repositoryJson(api, repository, 'manifest.json', ready.object.sha);
  const live = await repositoryJson(api, repository, '.scout-live.json', 'main');
  // Cleanup deploys the same release again. Do not form a Pages -> Scout loop.
  if (previous?.assets === manifest.assets && live?.current === manifest.assets.split('/').at(-1)) return 'already promoted';
  await verify();
  if ((await root())?.assets !== manifest.assets) return 'superseded';
  if (ready) {
    await api(`repos/${repository}/git/refs/heads/scout-ready`, { method: 'PATCH', body: { sha, force: true } });
  } else {
    await api(`repos/${repository}/git/refs`, { method: 'POST', body: { ref: 'refs/heads/scout-ready', sha } });
  }
  // Retry dispatch on reruns, even if updating scout-ready succeeded previously.
  await dispatch(`repos/${scoutRepo}/dispatches`, {
    method: 'POST', body: { event_type: 'app_source_updated', client_payload: { app: manifest.app, release: manifest.assets, distribution_sha: sha } },
  });
  return 'notified Scout';
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  if (!process.env.GH_TOKEN || !process.env.SCOUT_REPO_PAT) throw new Error('Missing Pages promotion credentials.');
  const result = await promotePages({
    repository: process.env.GITHUB_REPOSITORY,
    sha: process.env.GITHUB_SHA,
    manifest: JSON.parse(await readFile('manifest.json', 'utf8')),
    api: githubApi(process.env.GH_TOKEN),
    dispatch: githubApi(process.env.SCOUT_REPO_PAT),
    verify: async () => {
      const { stdout } = await promisify(execFile)(process.execPath, [new URL('./release-files.mjs', import.meta.url).pathname, 'verify', 'manifest.json'], {
        env: { ...process.env, RELEASE_WAIT_MS: process.env.RELEASE_WAIT_MS ?? '60000' },
      });
      process.stdout.write(stdout);
    },
  });
  console.log(result);
  if (process.env.GITHUB_OUTPUT) {
    await appendFile(process.env.GITHUB_OUTPUT, `cleanup=${result !== 'superseded'}\n`);
  }
}
