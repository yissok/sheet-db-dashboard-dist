import { pagesOrigin } from './config.mjs';
import { cp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const releasePattern = /^[a-f0-9]{40}-[0-9]+-[0-9]+$/;
const stateFile = '.scout-live.json';

export async function pruneReleases(distribution, promoted) {
  const root = JSON.parse(await readFile(join(distribution, 'manifest.json'), 'utf8'));
  // A publisher may have uploaded the next candidate while Pages deployed.
  // Never prune that candidate, overwrite its root, or promote stale state.
  if (root.version !== 4 || root.app !== promoted.app) return { skipped: true };
  const newerCandidate = root.assets !== promoted.assets;
  const prefix = `${pagesOrigin}/${promoted.app}-dist/releases/`;
  const current = promoted.assets?.startsWith(prefix) ? promoted.assets.slice(prefix.length) : '';
  if (!releasePattern.test(current)) throw new Error('Invalid promoted release.');
  let state;
  try { state = JSON.parse(await readFile(join(distribution, stateFile), 'utf8')); }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  if (state && (!releasePattern.test(state.current) || (state.previous !== null && !releasePattern.test(state.previous)))) {
    throw new Error('Invalid previous promotion state.');
  }
  const previous = state ? (state.current === current ? state.previous : state.current) : null;
  const keep = [current, previous].filter(Boolean);
  // Validate all retained releases before deleting anything.
  for (const release of keep) {
    const manifest = JSON.parse(await readFile(join(distribution, 'releases', release, 'manifest.json'), 'utf8'));
    if (manifest.assets !== `${prefix}${release}` || manifest.app !== promoted.app) throw new Error('Retained release manifest mismatch.');
  }
  if (newerCandidate) {
    // Record the successful promotion without touching the next candidate.
    await writeFile(join(distribution, stateFile), JSON.stringify({ current, previous }, null, 2) + '\n');
    return { skipped: true };
  }
  const removed = [];
  for (const entry of await readdir(join(distribution, 'releases'), { withFileTypes: true })) {
    if (state && entry.isDirectory() && releasePattern.test(entry.name) && !keep.includes(entry.name)) {
      await rm(join(distribution, 'releases', entry.name), { recursive: true });
      removed.push(entry.name);
    }
  }
  // Preserve legacy root assets for the first successful versioned promotion.
  // On later promotions, rebuild the root copy too, so stale bundles do not
  // accumulate there after their release folders are removed.
  if (previous) {
    const protectedNames = new Set(['.git', '.github', '.gitignore', '.gitattributes', '.nojekyll', 'CNAME', 'releases', stateFile]);
    for (const entry of await readdir(distribution)) {
      if (!protectedNames.has(entry)) await rm(join(distribution, entry), { recursive: true, force: true });
    }
    await cp(join(distribution, 'releases', current), distribution, { recursive: true });
  }
  await writeFile(join(distribution, stateFile), JSON.stringify({ current, previous }, null, 2) + '\n');
  return { kept: keep, removed, migration: !state };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const [distribution, manifestPath] = process.argv.slice(2);
  console.log(JSON.stringify(await pruneReleases(distribution, JSON.parse(await readFile(manifestPath, 'utf8')))));
}
