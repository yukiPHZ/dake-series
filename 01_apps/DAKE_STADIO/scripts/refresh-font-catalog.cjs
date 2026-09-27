// Explicit developer operation; the packaged application does not call this.
// Use --from-cache to regenerate from the recorded official responses offline.
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const root = path.resolve(__dirname, '..');
const sourceDir = path.join(root, 'evidence', 'font-catalog-sources');
const metadataUrl = 'https://fonts.google.com/metadata/fonts';
const repository = 'https://github.com/google/fonts';
const apiRoot = 'https://api.github.com/repos/google/fonts';
const headers = { 'User-Agent': 'DAKE-STADIO-build-catalog', Accept: 'application/vnd.github+json' };
const sha256 = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const normalized = value => value.toLowerCase().replace(/[^a-z0-9]/g, '');
async function fetchJson(url) {
  const response = await fetch(url, { headers, signal: AbortSignal.timeout(60000) });
  if (!response.ok) throw new Error(`Catalog source HTTP ${response.status}: ${url}`);
  const raw = await response.text();
  if (raw.length > 16 * 1024 * 1024) throw new Error('Catalog source exceeds 16 MiB');
  return JSON.parse(raw.replace(/^\)\]\}'[^\n]*\n/, ''));
}
async function readJson(name) { return JSON.parse(await fs.readFile(path.join(sourceDir, name), 'utf8')); }
async function writeJson(filename, value, pretty = false) {
  const temporary = `${filename}.tmp`;
  await fs.writeFile(temporary, JSON.stringify(value, null, pretty ? 2 : undefined) + '\n', 'utf8');
  await fs.rename(temporary, filename);
}
async function main() {
  await fs.mkdir(sourceDir, { recursive: true });
  let metadata, tree, commitInfo;
  if (process.argv.includes('--from-cache')) {
    [metadata, tree, commitInfo] = await Promise.all(['metadata.json', 'tree.json', 'commit.json'].map(readJson));
  } else {
    // Requests contain no artwork, text samples, project names or user data.
    metadata = await fetchJson(metadataUrl);
    const head = await fetchJson(`${apiRoot}/commits/main`);
    if (!/^[a-f0-9]{40}$/.test(head.sha)) throw new Error('Invalid source commit');
    tree = await fetchJson(`${apiRoot}/git/trees/${head.sha}?recursive=1`);
    commitInfo = { commit: head.sha, treeSha: tree.sha, metadataUrl, retrievedAt: new Date().toISOString() };
  }
  if (!/^[a-f0-9]{40}$/.test(commitInfo.commit) || tree.truncated || !Array.isArray(tree.tree)) throw new Error('Incomplete pinned source tree');
  if (!Array.isArray(metadata.familyMetadataList) || metadata.familyMetadataList.length < 1000) throw new Error('Unexpected family metadata schema');
  const directories = new Map();
  for (const item of tree.tree) {
    const match = /^(ofl|apache|ufl)\/([^/]+)\/([^/]+)$/.exec(item.path);
    if (!match || item.type !== 'blob') continue;
    const directory = `${match[1]}/${match[2]}`;
    if (!directories.has(match[2])) directories.set(match[2], new Map());
    const choices = directories.get(match[2]);
    if (!choices.has(directory)) choices.set(directory, []);
    choices.get(directory).push(item);
  }
  const skipped = [];
  const families = [];
  const weightNames = { extrablack: 950, ultrablack: 950, extrabold: 800, ultrabold: 800, semibold: 600, demibold: 600, extralight: 200, ultralight: 200, hairline: 100, thin: 100, light: 300, regular: 400, normal: 400, book: 400, medium: 500, bold: 700, black: 900, heavy: 900 };
  for (const family of metadata.familyMetadataList) {
    if (typeof family.family !== 'string' || !Array.isArray(family.subsets)) throw new Error('Invalid family metadata');
    const choices = directories.get(normalized(family.family));
    if (!choices || choices.size !== 1) { skipped.push({ family: family.family, reason: 'No unique exact normalized directory match' }); continue; }
    const [directory, items] = [...choices][0];
    const license = items.find(item => /\/(?:OFL|LICENSE|LICENCE|UFL)\.txt$/i.test(item.path));
    const fontItems = items.filter(item => /\.(?:ttf|otf)$/i.test(item.path));
    if (!license || !fontItems.length) { skipped.push({ family: family.family, directory, reason: 'No root font or license file' }); continue; }
    const axes = (family.axes || []).filter(axis => typeof axis.tag === 'string').map(axis => ({ tag: axis.tag, min: axis.min, max: axis.max, defaultValue: axis.defaultValue }));
    const styles = Object.keys(family.fonts || {});
    const files = fontItems.map(item => {
      const filename = item.path.split('/').pop();
      const axisTags = /\[([^\]]+)\]/.exec(filename)?.[1].split(',') || [];
      const weightAxis = axisTags.includes('wght') && axes.find(axis => axis.tag === 'wght');
      const suffix = filename.replace(/\.[^.]+$/, '').split('-').slice(1).join('').toLowerCase();
      const style = /italic|oblique/i.test(filename) ? 'italic' : 'normal';
      const namedWeight = Object.keys(weightNames).find(name => suffix.includes(name));
      const onlyStyle = styles.filter(value => value.endsWith('i') === (style === 'italic'));
      const fallbackWeight = onlyStyle.length === 1 ? Number.parseInt(onlyStyle[0], 10) : 400;
      const weight = weightAxis ? weightAxis.defaultValue : namedWeight ? weightNames[namedWeight] : fallbackWeight;
      return { path: item.path, filename, size: item.size, blobSha: item.sha, style, weight, weightMin: weightAxis ? weightAxis.min : weight, weightMax: weightAxis ? weightAxis.max : weight, axisTags };
    }).sort((a, b) => a.style.localeCompare(b.style) || a.weight - b.weight || a.path.localeCompare(b.path));
    const preferred = [...files].sort((a, b) => Number(a.style !== 'normal') - Number(b.style !== 'normal') || Math.abs(a.weight - 400) - Math.abs(b.weight - 400) || Number(b.axisTags.includes('wght')) - Number(a.axisTags.includes('wght')))[0];
    families.push({ family: family.family, displayName: family.displayName || family.family, category: family.category, subsets: family.subsets, primaryScript: family.primaryScript || '', axes, variants: styles, lastModified: family.lastModified, directory, licenseType: { ofl: 'OFL-1.1', apache: 'Apache-2.0', ufl: 'UFL-1.0' }[directory.split('/')[0]], licensePath: license.path, licenseBlobSha: license.sha, defaultFile: preferred.path, files });
  }
  families.sort((a, b) => a.family.localeCompare(b.family, 'en'));
  const japaneseCount = families.filter(family => family.subsets.includes('japanese')).length;
  if (families.length < 1000 || japaneseCount < 50 || new Set(families.map(f => f.family)).size !== families.length) throw new Error('Catalog coverage or uniqueness check failed');
  if (!families.some(f => f.family === 'Noto Sans JP') || !families.some(f => f.family === 'BIZ UDPGothic')) throw new Error('Required Japanese families missing');
  const catalog = { schemaVersion: 1, generatedAt: commitInfo.retrievedAt, sourceCommit: commitInfo.commit, sourceTreeSha: tree.sha, sources: { metadata: metadataUrl, repository, tree: `${apiRoot}/git/trees/${commitInfo.commit}?recursive=1` }, familyCount: families.length, japaneseFamilyCount: japaneseCount, families };
  const output = path.join(root, 'assets', 'google-fonts-catalog.json');
  await writeJson(output, catalog);
  if (!process.argv.includes('--from-cache')) {
    await writeJson(path.join(sourceDir, 'metadata.json'), metadata);
    await writeJson(path.join(sourceDir, 'tree.json'), tree);
    await writeJson(path.join(sourceDir, 'commit.json'), commitInfo, true);
  }
  const outputBytes = await fs.readFile(output);
  const results = { generatedAt: new Date().toISOString(), sourceCommit: commitInfo.commit, metadataFamilies: metadata.familyMetadataList.length, pinnedTreeEntries: tree.tree.length, treeTruncated: tree.truncated, families: families.length, japaneseFamilies: japaneseCount, fontFiles: families.reduce((sum, f) => sum + f.files.length, 0), licenseCounts: Object.fromEntries(['OFL-1.1', 'Apache-2.0', 'UFL-1.0'].map(type => [type, families.filter(f => f.licenseType === type).length])), skipped, bytes: outputBytes.length, sha256: sha256(outputBytes), checks: ['Pinned complete official source tree', 'Each family has root font and license', 'All download paths come from tree blobs', 'Unique family names', 'At least 1000 families and 50 Japanese families', 'Noto Sans JP and BIZ UDPGothic included'], notes: ['Metadata endpoint is not a documented public API; a validated snapshot is bundled and runtime search is offline.', 'Filename-inferred file style and static weight are hints; actual downloaded font metadata must be parsed for rendering.', 'Catalog generation does not download or bundle font binaries.'] };
  await writeJson(path.join(root, 'evidence', 'font-catalog-results.json'), results, true);
  process.stdout.write(JSON.stringify(results, null, 2) + '\n');
}
main().catch(error => { console.error(error); process.exitCode = 1; });