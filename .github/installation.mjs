// Public installation settings. Never put tokens or service-account keys here.
/** @type {Readonly<Record<string, string>>} */
export const installation = Object.freeze({
  githubOwner: 'yissok',
  scoutRepository: 'scout',
  factoryRepository: 'project-factory',
  firebaseProject: 'scout-38aa2',
  hostingSite: 'scout-38aa2',
  publicUrl: 'https://scout-38aa2.web.app',
  rootApp: 'cell',
  contentEditor: 'enabled', // Set to 'disabled' unless you use the portfolio content schema.
  dnsProvider: 'one.com', // 'generic' for other DNS providers.
  contentRepository: 'personal',
  mediaBaseUrl: 'https://yissok.github.io/HQ_videos',
  previewUrl: 'https://personal-39136.web.app/editor-preview',
  iosBundleId: 'com.example.factory',
  iosAppName: 'factory',
  iosDevelopmentTeam: 'BR2XN9F953', // Your Apple team ID, or blank to choose in Xcode.
});
if (!['enabled', 'disabled'].includes(installation.contentEditor) || !['generic', 'one.com'].includes(installation.dnsProvider)) throw Error('Invalid optional integration setting.');
const name = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
if (!/^[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*$/.test(installation.githubOwner)) throw Error('Invalid GitHub owner.');
for (const key of ['scoutRepository', 'factoryRepository', 'contentRepository', 'rootApp']) {
  if (!name.test(installation[key])) throw Error(`Invalid installation setting: ${key}`);
}
for (const key of ['firebaseProject']) {
  if (!/^[a-z][a-z0-9-]{4,28}[a-z0-9]$/.test(installation[key])) throw Error(`Invalid installation setting: ${key}`);
}
if (!/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/.test(installation.hostingSite)) throw Error('Invalid Hosting site ID.');
if (!/^(?:[A-Z0-9]{10})?$/.test(installation.iosDevelopmentTeam)) throw Error('Invalid Apple development team.');
if (!/^[A-Za-z0-9.-]+$/.test(installation.iosBundleId)) throw Error('Invalid iOS bundle ID.');
for (const key of ['publicUrl', 'mediaBaseUrl', 'previewUrl']) {
  const url = new URL(installation[key]);
  if (url.protocol !== 'https:' || url.username || url.password || url.port || url.search || url.hash || installation[key].endsWith('/')) throw Error(`Invalid HTTPS installation setting: ${key}`);
}
if (new URL(installation.publicUrl).pathname !== '/') throw Error('publicUrl must be an origin without a trailing slash.');
export const owner = installation.githubOwner;
export const scoutRepo = `${owner}/${installation.scoutRepository}`;
export const factoryRepo = `${owner}/${installation.factoryRepository}`;
export const contentRepo = `${owner}/${installation.contentRepository}`;
export const pagesOrigin = `https://${owner.toLowerCase()}.github.io`;
export const publicUrl = installation.publicUrl;
/** @param {string} value */
export function repositoryName(value) {
  let name = value.trim();
  for (const prefix of [`https://github.com/${owner}/`, `${owner}/`]) {
    if (name.startsWith(prefix)) { name = name.slice(prefix.length); break; }
  }
  return name.replace(/\/$/, '');
}
