import { pagesOrigin } from './config.mjs';
import { createHash } from 'node:crypto';
import { readFile, readdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

export const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');

export async function inventory(directory, prefix = '') {
  const files = {};
  for (const entry of await readdir(join(directory, prefix), { withFileTypes: true })) {
    if (['.gitkeep', '.nojekyll', '.DS_Store'].includes(entry.name)) continue;
    const path = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isDirectory()) Object.assign(files, await inventory(directory, path));
    else if (entry.isFile() && path !== 'release.json') files[path] = digest(await readFile(join(directory, path)));
    else if (!entry.isFile()) throw new Error(`Unsupported release entry: ${path}`);
  }
  return files;
}

export async function verifyRelease(manifest, { fetcher = fetch, concurrency = 8, verified = new Set(), scope = 'all' } = {}) {
  if (!['all', 'main'].includes(scope)) throw new Error('Release verification scope must be all or main.');
  if (!Number.isInteger(concurrency) || concurrency < 1 || concurrency > 64) {
    throw new Error('Release verification concurrency must be an integer from 1 to 64.');
  }
  if (manifest.version !== 4) throw new Error('Expected a version 4 release manifest.');
  const expected = `${pagesOrigin}/${manifest.app}-dist/releases/`;
  if (typeof manifest.assets !== 'string' || !manifest.assets.startsWith(expected)
      || !/^[a-f0-9]{40}-[0-9]+-[0-9]+$/.test(manifest.assets.slice(expected.length))) {
    throw new Error('Invalid immutable release URL.');
  }
  const base = `${manifest.assets}/`;
  async function bytes(path) {
    const response = await fetcher(new URL(path, base), { cache: 'no-store', signal: AbortSignal.timeout(30000) });
    if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`);
    return Buffer.from(await response.arrayBuffer());
  }
  const index = JSON.parse((await bytes('release.json')).toString());
  let entries = Object.entries(index.files ?? {});
  if (index.version !== 1 || !entries.length || !index.files['manifest.json'] || !index.files['index.html']
      || !index.files[manifest.start] || !index.files[manifest.appEntry]) throw new Error('Incomplete release inventory.');
  for (const [path, hash] of entries) {
    if (!/^[A-Za-z0-9_.][A-Za-z0-9._/-]*$/.test(path) || path.split('/').some((part) => part === '..' || part === '.')
        || path.includes('//') || !/^[a-f0-9]{64}$/.test(hash)) throw new Error(`Invalid inventory entry: ${path}`);
  }
  if (scope === 'main') {
    const required = new Set(['manifest.json', 'index.html', manifest.start, manifest.appEntry]);
    for (const resource of [...(manifest.styles ?? []), ...(manifest.scripts ?? []), ...(manifest.preloads ?? [])]) {
      const url = new URL(resource, base);
      // Third-party resources are outside this release's integrity inventory.
      if (url.href.startsWith(base)) required.add(url.pathname.slice(new URL(base).pathname.length));
    }
    for (const path of required) {
      if (!Object.hasOwn(index.files, path)) throw new Error(`Missing main-page inventory entry: ${path}`);
    }
    entries = entries.filter(([path]) => required.has(path));
  }
  let next = 0;
  const results = await Promise.allSettled(Array.from({ length: Math.min(concurrency, entries.length) }, async () => {
    while (next < entries.length) {
      const [path, hash] = entries[next++];
      // Only reuse checks within this process, for the same immutable URL and hash.
      // Always compare the published manifest against the selected manifest again.
      const key = JSON.stringify([base, path, hash]);
      if (path !== 'manifest.json' && verified.has(key)) continue;
      const content = await bytes(path);
      if (digest(content) !== hash) throw new Error(`Release content mismatch: ${path}`);
      if (path === 'manifest.json' && JSON.stringify(JSON.parse(content)) !== JSON.stringify(manifest)) {
        throw new Error('Published manifest does not match the selected release.');
      }
      verified.add(key);
    }
  }));
  const failure = results.find((result) => result.status === 'rejected');
  if (failure) throw failure.reason;
  return entries.length;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [mode, path] = process.argv.slice(2);
  if (mode === 'inventory') {
    await writeFile(join(path, 'release.json'), JSON.stringify({ version: 1, files: await inventory(path) }) + '\n');
  } else if (mode === 'verify') {
    const manifest = JSON.parse(await readFile(path, 'utf8'));
    const concurrency = Number(process.env.RELEASE_VERIFY_CONCURRENCY ?? 8);
    const scope = process.env.RELEASE_VERIFY_SCOPE ?? 'all';
    const verified = new Set();
    const started = Date.now();
    const deadline = Date.now() + Number(process.env.RELEASE_WAIT_MS ?? 900000);
    for (;;) {
      try {
        console.log(`Verified ${await verifyRelease(manifest, { concurrency, verified, scope })} files for ${manifest.app}: ${manifest.assets} (${((Date.now() - started) / 1000).toFixed(1)}s, scope ${scope}, concurrency ${concurrency})`);
        break;
      } catch (error) {
        if (Date.now() >= deadline) throw error;
        console.log(`Release not ready: ${error.message}. Retrying in 15 seconds.`);
        await new Promise((resolve) => setTimeout(resolve, 15000));
      }
    }
  } else throw new Error('Usage: release-files.mjs inventory <directory> | verify <manifest>');
}
