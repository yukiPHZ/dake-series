'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');

// A child document is canonical. Group objects are a regenerated rendering cache.
function validateComponents(data, validateLeaf) {
  let documents = 0, objects = 0;
  const visitDocument = (doc, ancestors = [], depth = 0) => {
    if (!doc || typeof doc !== 'object' || ![1,2,3].includes(doc.version) || doc.format !== 'dake-stadio' ||
        !Array.isArray(doc.canvas?.objects) || depth > 8 || ++documents > 64) throw new Error('COMPONENT_INVALID');
    const id = doc.documentId;
    if (id !== undefined && (typeof id !== 'string' || id.length > 160 || !id.length)) throw new Error('COMPONENT_INVALID');
    if (id && ancestors.includes(id)) throw new Error('COMPONENT_CYCLE');
    const next = id ? [...ancestors, id] : ancestors;
    if (validateLeaf && depth) validateLeaf(doc);
    const visitObject = object => {
      if (++objects > 12000) throw new Error('COMPONENT_LIMIT');
      const component = object?.dakeComponent;
      if (component) {
        if (String(object.type).toLowerCase() !== 'group' || doc.version !== 3 ||
            typeof component.id !== 'string' || component.id.length > 160 ||
            !['embedded','linked'].includes(component.mode)) throw new Error('COMPONENT_INVALID');
        if (component.originPath !== undefined && (typeof component.originPath !== 'string' || component.originPath.length > 8192)) throw new Error('COMPONENT_INVALID');
        if (component.mode === 'linked') {
          const link = component.link;
          if (!link || typeof link.path !== 'string' || link.path.length > 8192 || !/^[a-f0-9]{64}$/i.test(link.hash || '') ||
              (link.token !== undefined && (typeof link.token !== 'string' || link.token.length > 160))) throw new Error('COMPONENT_INVALID');
          if (link.path.includes('\0') || !/\.dake$/i.test(link.path)) throw new Error('COMPONENT_INVALID');
        }
        visitDocument(component.document, next, depth + 1);
        // Check resource envelopes in cache as well; logical dependency cycles live in document.
        for (const child of object.objects || []) if (++objects > 12000) throw new Error('COMPONENT_LIMIT');
      } else for (const child of object?.objects || []) visitObject(child);
    };
    for (const object of doc.canvas.objects) visitObject(object);
  };
  visitDocument(data);
  return data;
}
function createComponentService({ readProject, rememberSource = async () => {} }) {
  const grants = new Map();
  const read = async (filePath, mode, token) => {
    if (!['embedded','linked'].includes(mode) || typeof filePath !== 'string' || !/\.dake$/i.test(filePath)) throw new Error('COMPONENT_INVALID');
    const canonical = await fs.realpath(filePath);
    const data = await readProject(canonical);
    validateComponents(data);
    const hash = crypto.createHash('sha256').update(await fs.readFile(canonical)).digest('hex');
    data.documentId ||= 'file-' + hash;
    await rememberSource(canonical);
    const capability = token || crypto.randomUUID();
    grants.set(capability, canonical);
    return { data, originPath:canonical, name:path.basename(canonical), mode, link:mode==='linked'?{path:canonical,hash,token:capability}:undefined };
  };
  return {
    importFile: (filePath, mode = 'embedded') => read(filePath, mode),
    refresh: token => {
      if (typeof token !== 'string' || !grants.has(token)) throw new Error('COMPONENT_RELINK');
      return read(grants.get(token), 'linked', token);
    }
  };
}
module.exports = { validateComponents, createComponentService };
