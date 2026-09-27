'use strict';
require('./local-test-env.cjs')();
// Integration QA only: real Electron/Canvas/IPC/files, DOM controls and CDP input.
// Native file dialogs return isolated fixture paths; the renderer bridge is untouched.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { app, BrowserWindow, dialog, session, net, nativeImage } = require('electron');
const root = path.resolve(__dirname, '..');
const output = path.resolve(process.env.STADIO_PRODUCTION_OUTPUT || path.join(root, 'test-output', `production-${Date.now()}`));
const phase = process.argv[2] || 'prepare';
if (!output.startsWith(path.join(root, 'test-output') + path.sep) || !['prepare', 'resume', 'safety'].includes(phase)) throw new Error('Expected isolated production output and phase');
fs.mkdirSync(output, { recursive: true });
process.env.STADIO_USER_DATA = path.join(output, `profile-${phase}`);
let saveTarget, openTargets = [];
const allowed = file => path.resolve(file).startsWith(output + path.sep);
dialog.showSaveDialog = async () => { if (!saveTarget || !allowed(saveTarget)) throw new Error('Unexpected QA save destination'); return { canceled: false, filePath: saveTarget }; };
dialog.showOpenDialog = async () => { if (!openTargets.length || openTargets.some(file => !allowed(file))) throw new Error('Unexpected QA open destination'); return { canceled: false, filePaths: openTargets }; };
dialog.showErrorBox = (title, message) => { throw new Error(`${title}: ${message}`); };
const report = { phase, processId: process.pid, started: new Date().toISOString(), output, checks: [], steps: [], errors: [], network: [], works: {} };
const external = value => /^https?:/i.test(typeof value === 'string' ? value : value?.url || String(value));
const recordRequest = (transport, input) => { const url = typeof input === 'string' ? input : input?.url || String(input); if (external(url)) { report.network.push({ transport, url, blocked: phase !== 'prepare' }); if (phase !== 'prepare') throw new Error('QA_OFFLINE: external network disabled'); } };
const https = require('node:https');
const originalHttpsGet = https.get;
https.get = function(input, ...args) { recordRequest('node-https-get', input); return originalHttpsGet.call(this, input, ...args); };
const originalFetch = globalThis.fetch;
globalThis.fetch = function(input, init) { recordRequest('node-fetch', input); return originalFetch.call(this, input, init); };
const originalNetFetch = net.fetch;
if (originalNetFetch) net.fetch = function(input, init) { recordRequest('electron-net-fetch', input); return originalNetFetch.call(this, input, init); };
require('../desktop/main.cjs');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const hash = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const json = value => JSON.stringify(value);
let win, activeWork = '';
const js = expression => win.webContents.executeJavaScript(expression, true);
const command = (method, params = {}) => win.webContents.debugger.sendCommand(method, params);
function check(name, value, detail) { assert.ok(value, `${activeWork ? activeWork + ': ' : ''}${name}`); report.checks.push({ work: activeWork || null, name, ...(detail === undefined ? {} : { detail }) }); }
async function waitUntil(expression, timeout = 15000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) { if (await js(expression)) return; await delay(35); }
  const context = await js(`({status:document.getElementById('status')?.textContent,dialog:document.querySelector('.dialog')?.textContent?.slice(0,500),busy:__studio?.engine.busy})`).catch(() => ({}));
  throw new Error(`Timed out: ${expression}\n${json(context)}`);
}
async function idle() { await waitUntil('!!window.__studio && !__studio.engine.busy', 60000); await delay(55); }
async function click(selector) {
  await js(`(()=>{const el=document.querySelector(${json(selector)});if(!el)throw new Error('Missing control '+${json(selector)});if(el.disabled)throw new Error('Disabled control '+${json(selector)});for(let p=el.parentElement;p;p=p.parentElement)if(p.tagName==='DETAILS')p.open=true;el.scrollIntoView({block:'nearest'});el.click();})()`);
  await idle();
}
async function change(id, value) {
  await js(`(()=>{const el=document.getElementById(${json(id)});if(!el)throw new Error('Missing field '+${json(id)});if(el.disabled)throw new Error('Disabled field '+${json(id)});if(el.type==='checkbox')el.checked=${json(value)};else el.value=${json(value)};el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await idle();
}
async function step(name, fn) { const start = Date.now(); await fn(); report.steps.push({ work: activeWork || null, name, ms: Date.now() - start }); console.log(`[production ${phase}] ${activeWork} ${name}`); }
async function selectedId() { return js('__studio.engine.selected?.id'); }
async function select(ids) {
  for (let i = 0; i < ids.length; i++) {
    await js(`(()=>{const row=document.querySelector('[data-layer="'+${json(ids[i])}+'"]');if(!row)throw new Error('Layer missing');row.dispatchEvent(new MouseEvent('click',{bubbles:true,ctrlKey:${i > 0}}));})()`); await idle();
  }
  check('layer panel selection matches requested layers', await js(`__studio.engine.canvas.getActiveObjects().length===${ids.length}`));
}
async function deselect() { await js(`__studio.engine.canvas.discardActiveObject();__studio.engine.canvas.requestRenderAll();__studio.engine.onSelection(null);`); await idle(); }
async function position(values) {
  for (const [key,value] of Object.entries(values)) {
    const conversion=await js(`(()=>{const el=document.getElementById('prop-'+${json(key)});return /\\(mm\\)/.test(el?.closest('label')?.querySelector('span')?.textContent||'')?25.4/__studio.engine.meta.dpi:1;})()`);
    await change('prop-'+key,value*conversion);
  }
}
async function color(value) { await change('style-fill', value); await change('style-width', 0); }
async function rename(id, name) {
  await js(`document.querySelector('[data-layer="'+${json(id)}+'"]').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}))`);
  await waitUntil(`!!document.getElementById('rename-value')`); await change('rename-value', name); await click('#dialog-apply');
  check('layer rename persists through panel dialog', await js(`__studio.engine.layers.find(o=>o.id===${json(id)})?.name===${json(name)}`));
}
async function newWork(kind, title, dimensions = {}) {
  activeWork = kind;
  await js(`document.activeElement?.blur();document.body.focus();`);
  win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'N', modifiers: ['control'] });
  win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'N', modifiers: ['control'] });
  await waitUntil(`!!document.getElementById('new-preset')`);
  await change('new-preset', kind); await change('new-name', title);
  for (const [key, value] of Object.entries(dimensions)) await change('new-' + key, value);
  await click('#dialog-create'); await waitUntil(`__studio.engine.name===${json(title)} && !document.querySelector('.dialog')`);
  report.works[kind] = { title, width: await js('__studio.engine.width'), height: await js('__studio.engine.height'), meta: await js('structuredClone(__studio.engine.meta)') };
  check('preset and dimension fields produce new editable artwork', await js('__studio.engine.layers.length===0 && __studio.dirty'));
}
async function scenePoint(x, y) { return js(`(()=>{const r=document.querySelector('.canvas-host').getBoundingClientRect(),v=__studio.engine.canvas.viewportTransform;return{x:r.left+v[4]+${x}*v[0],y:r.top+v[5]+${y}*v[3]};})()`); }
async function mouse(type, point, extra = {}) { await command('Input.dispatchMouseEvent', { type, ...point, ...(type === 'mouseMoved' ? {} : { button: 'left', buttons: type === 'mousePressed' ? 1 : 0, clickCount: 1 }), ...extra }); }
async function drop(file, x, y) {
  const point = await scenePoint(x, y), before = await js('__studio.engine.layers.length');
  const data = { items: [], files: [file], dragOperationsMask: 1 };
  for (const type of ['dragEnter', 'dragOver', 'drop']) await command('Input.dispatchDragEvent', { type, ...point, data });
  await waitUntil(`__studio.engine.layers.length>${before}`, 20000); await idle();
  check('real file drop passes through Chromium File and native path boundary', true, path.basename(file));
  return selectedId();
}
async function shape(kind, name, box, fill) { await click(`[data-tool="${kind}"]`); await position(box); await color(fill); const id = await selectedId(); await rename(id, name); return id; }
async function chooseGoogleFont(family = 'Noto Sans JP', weight = 400) {
  await click('#choose-font'); await click('#font-google'); await waitUntil(`document.querySelectorAll('.font-row').length>0`, 20000);
  await change('font-search', family); await change('font-download-weight', weight); await change('font-download-style', 'normal');
  await js(`(()=>{const row=[...document.querySelectorAll('.font-row')].find(row=>row.querySelector('strong')?.textContent===${json(family)});if(!row)throw new Error('Font missing from catalog');row.querySelector('button').click();})()`);
  await waitUntil(`__studio.engine.fonts?.some(f=>f.family===${json(family)}&&f.data&&f.license)&&__studio.engine.selected?.fontFamily===${json(family)}`, 120000);
  check('Google Fonts UI uses actual downloaded font with embedded license', await js(`__studio.engine.fonts.some(f=>f.family===${json(family)}&&f.source==='google'&&f.data.length>100000&&f.license.length>1000)`));
  await click('#dialog-close');
}
async function text(name, content, box, fontSize, fill, google = false) {
  await click('[data-tool="text"]'); await change('prop-text', content); await change('prop-fontsize', fontSize);
  if (google) await chooseGoogleFont(); else await change('prop-font', 'Noto Sans JP');
  await change('prop-tracking', .025); await change('prop-leading', 1.35); await change('prop-weight', 500);
  await color(fill); await position(box);
  const id = await selectedId(); await rename(id, name);
  check('text tracking leading and font are edited through DOM fields', await js(`__studio.engine.selected.fontFamily==='Noto Sans JP'&&__studio.engine.selected.charSpacing===25&&__studio.engine.selected.lineHeight===1.35`));
  return id;
}
async function guides(vertical, horizontal) {
  await deselect(); await change('guide-axis', 'x'); await change('guide-pos', vertical); await click('#guide-add');
  await deselect(); await change('guide-axis', 'y'); await change('guide-pos', horizontal); await click('#guide-add');
  check('guide panel stores horizontal and vertical guides', await js('__studio.engine.meta.guides.vertical.length>0&&__studio.engine.meta.guides.horizontal.length>0'));
}
async function screenshot(name) { const shot = await command('Page.captureScreenshot', { format: 'png' }); fs.writeFileSync(path.join(output, name), Buffer.from(shot.data, 'base64')); }
async function saveProject(kind, suffix = '') {
  saveTarget = path.join(output, `${kind}${suffix}.dake`);
  if (suffix) { await js(`document.activeElement?.blur();`); win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'S', modifiers: ['control', 'shift'] }); win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'S', modifiers: ['control', 'shift'] }); }
  else await click('#save');
  await waitUntil(`!__studio.dirty&&__studio.currentPath===${json(saveTarget)}`, 30000);
  const document = JSON.parse(fs.readFileSync(saveTarget, 'utf8'));
  check('native save writes versioned editable document and embedded font', document.version === 2 && document.canvas.objects.length > 2 && document.fonts?.some(f => f.data && f.license));
  return document;
}
async function baseline(kind) {
  await deselect(); await saveProject(kind);
  const png = await js(`__studio.engine.exportRaster('png',1)`);
  fs.writeFileSync(path.join(output, `${kind}-before.png`), Buffer.from(png.split(',')[1], 'base64'));
  const extent=await js(`(()=>{let visible=0;for(const object of __studio.engine.layers){if(!object.visible)continue;const b=object.getBoundingRect();if(b.left+b.width>0&&b.top+b.height>0&&b.left<__studio.engine.width&&b.top<__studio.engine.height)visible++;}return visible;})()`);
  check('at least three visible composition elements overlap the artboard',extent>=3,extent);
  const bitmap=nativeImage.createFromBuffer(fs.readFileSync(path.join(output,`${kind}-before.png`))).toBitmap();let ink=0;
  for(let i=0;i<bitmap.length;i+=4)if(bitmap[i+3]>20&&Math.min(bitmap[i],bitmap[i+1],bitmap[i+2])<230)ink++;
  check('rendered artwork is not an empty or white canvas',ink/(bitmap.length/4)>.005,{inkPixels:ink,ratio:ink/(bitmap.length/4)});
  report.works[kind].baselinePngSha256 = hash(fs.readFileSync(path.join(output, `${kind}-before.png`)));
  report.works[kind].documentSha256 = hash(fs.readFileSync(path.join(output, `${kind}.dake`)));
  report.works[kind].meta = await js('structuredClone(__studio.engine.meta)');
  report.works[kind].textState = await js(`__studio.engine.layers.filter(o=>o.type.toLowerCase()==='textbox').map(o=>({id:o.id,text:o.text,fontFamily:o.fontFamily,fontWeight:o.fontWeight,charSpacing:o.charSpacing,lineHeight:o.lineHeight,left:o.left,top:o.top}))`);
  await screenshot(`${kind}-screen.png`);
}
async function prepareSources() {
  const data = await js(`(()=>{const c=document.createElement('canvas');c.width=900;c.height=720;const x=c.getContext('2d'),g=x.createLinearGradient(0,0,900,720);g.addColorStop(0,'#426a60');g.addColorStop(1,'#dce8c4');x.fillStyle=g;x.fillRect(0,0,900,720);for(let i=0;i<14;i++){x.fillStyle=i%2?'#f5df8b80':'#183f3640';x.beginPath();x.ellipse(510+i*11,80+i*39,250-i*8,150+i*7,-.35,0,Math.PI*2);x.fill();}return c.toDataURL('image/png');})()`);
  fs.writeFileSync(path.join(output, 'source-photo.png'), Buffer.from(data.split(',')[1], 'base64'));
  fs.writeFileSync(path.join(output, 'source-mark.svg'), '<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300" viewBox="0 0 300 300"><path d="M35 190 C40 45 230 15 265 140 C220 95 150 290 35 190Z" fill="#416b5d"/><circle cx="212" cy="74" r="29" fill="#e9bd54"/></svg>');
  report.sources = Object.fromEntries(['source-photo.png','source-mark.svg'].map(name => [name, hash(fs.readFileSync(path.join(output, name)))]));
}
async function makeBanner() {
  await newWork('banner', '暮らしを、軽やかに。', { width: 1920, height: 1080 });
  await step('drop, raster painting and nondestructive mask', async () => {
    const image = await drop(path.join(output, 'source-photo.png'), 1030, 155);
    await position({ width: 770, height: 750, x: 1030, y: 155 }); await change('adjust-brightness', 9); await rename(image, '背景の写真');
    const before = await js('__studio.engine.selected.toObject().src');
    await click('[data-tool="brush"]'); await change('brush-size', 24);
    const a = await scenePoint(1330, 640), b = await scenePoint(1460, 690);
    await mouse('mouseMoved', a); await mouse('mousePressed', a); await mouse('mouseMoved', b, { buttons: 1 }); await mouse('mouseReleased', b); await idle();
    check('native pointer brush changes pixels in the imported image', before !== await js('__studio.engine.layers.find(o=>o.id===' + json(image) + ').toObject().src'));
    await click('[data-tool="select"]');
    const mask = await shape('ellipse', '写真のマスク', { x: 1030, y: 155, width: 770, height: 750 }, '#ffffff');
    await select([image, mask]); await click('#mask-create');
    await waitUntil('__studio.engine.selected?.dakeMask?.enabled');
    const masked = await selectedId(); await click('#mask-toggle'); check('mask can be disabled without losing source', await js('__studio.engine.selected.dakeMask.enabled===false')); await click('#mask-toggle');
    await click('#mask-release'); await waitUntil(`__studio.engine.layers.some(o=>o.id===${json(image)})&&__studio.engine.layers.some(o=>o.id===${json(mask)})`);
    check('mask release restores image and vector source layers', true); await select([image, mask]); await click('#mask-create'); await rename(await selectedId(), '写真 / 編集可能なマスク');
  });
  await step('Japanese typography, layout and layer grouping', async () => {
    const title = await text('見出し', '暮らしを、\n軽やかに。', { x: 110, y: 255 }, 122, '#214639', true);
    await text('説明', '暮らしの道具 展\n2026.10.24 – 25  /  10:00–18:00', { x: 120, y: 635 }, 39, '#466859');
    const pill = await shape('rect', 'ラベル', { x: 122, y: 165, width: 240, height: 52 }, '#d9e6bd');
    const line = await shape('rect', '区切り線', { x: 120, y: 816, width: 535, height: 7 }, '#e1bb67');
    await select([pill, line]); await click('#align-left');
    check('alignment UI aligns selected layers', await js(`(()=>{const b=__studio.engine.canvas.getActiveObjects().map(o=>o.getBoundingRect().left);return Math.abs(b[0]-b[1])<.01;})()`));
    await click('#group'); check('group UI creates editable group', await js(`__studio.engine.selected.type.toLowerCase()==='group'`)); await click('#ungroup');
    const dragBefore=await js('__studio.engine.layers.map(o=>o.id)');
    await js(`(()=>{const from=document.querySelector('[data-layer="'+${json(dragBefore[0])}+'"]'),to=document.querySelector('[data-layer="'+${json(dragBefore.at(-1))}+'"]'),data=new DataTransfer();from.dispatchEvent(new DragEvent('dragstart',{bubbles:true,dataTransfer:data}));to.dispatchEvent(new DragEvent('dragover',{bubbles:true,cancelable:true,dataTransfer:data}));to.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true,dataTransfer:data}));from.dispatchEvent(new DragEvent('dragend',{bubbles:true,dataTransfer:data}));})()`);await idle();
    check('layer DOM drag-and-drop changes stacking order',await js(`__studio.engine.layers.at(-1).id===${json(dragBefore[0])}`));
    await guides(120, 165); await select([title]); await change('prop-tracking', .04);
  });
  await baseline('banner');
}
async function makeFlyer() {
  await newWork('flyer', '暮らしの道具 展 / チラシ');
  await drop(path.join(output, 'source-photo.png'), 330, 1380); await position({ x: 300, y: 1380, width: 1950, height: 1180 }); await change('adjust-saturation', -12); await rename(await selectedId(), '写真');
  await text('チラシ見出し', '暮らしの道具 展', { x: 250, y: 390 }, 181, '#234a3b', true);
  await text('チラシ説明', '触れて、選ぶ。\n明日からの暮らしに、新しいひとつを。', { x: 263, y: 800 }, 69, '#466658');
  await text('日程', '2026年10月24日（土）・25日（日）\n10:00–18:00　/　入場無料', { x: 264, y: 2800 }, 66, '#254837');
  await text('会場', '緑のギャラリー\n架空市みどり町1-2-3', { x: 265, y: 3190 }, 50, '#536e62');
  await guides(20, 30);
  await click('#layout-settings'); await change('snap-grid', true); await change('grid-size', 5); await click('#dialog-apply');
  check('layout dialog persists print unit dpi and snap settings', await js(`__studio.engine.meta.unit==='mm'&&__studio.engine.meta.dpi===300&&__studio.engine.meta.snap.grid`));
  await baseline('flyer');
}
async function penPath() {
  await click('[data-tool="pen"]');
  const points = [[180,690],[390,600],[585,690],[745,600]];
  for (let i=0;i<points.length;i++) { const p=await scenePoint(...points[i]); await mouse('mouseMoved',p); await mouse('mousePressed',p); if(i===1||i===2)await mouse('mouseMoved',{x:p.x+30,y:p.y-24},{buttons:1}); await mouse('mouseReleased',i===1||i===2?{x:p.x+30,y:p.y-24}:p); }
  await click('#close-path'); await idle(); await click('[data-tool="select"]');
  check('native pointer pen creates a closed cubic path', await js(`__studio.engine.layers.some(o=>o.path?.some(c=>c[0]==='C')&&o.path.some(c=>c[0]==='Z'))`));
}
async function makeLogo() {
  await newWork('logo', 'MIDORI / ロゴ');
  await step('boolean shapes and editable bezier nodes', async () => {
    const a=await shape('ellipse','輪郭',{x:185,y:145,width:385,height:385},'#315d4e');
    const b=await shape('rect','角の伸び',{x:430,y:330,width:185,height:200},'#315d4e');
    await select([a,b]);await click('#vector-union');await waitUntil(`__studio.engine.selected?.type.toLowerCase()==='path'`);
    const union=await selectedId();const hole=await shape('ellipse','抜き形',{x:282,y:240,width:185,height:185},'#ffffff');
    await select([union,hole]);await click('#vector-subtract');await waitUntil(`__studio.engine.selected?.type.toLowerCase()==='path'`);
    check('union and subtract produce a vector with independent contours',await js(`__studio.engine.selected.path.filter(c=>c[0]==='M').length>=2`));
    const count=await js('__studio.engine.selected.path.length');await change('node-index',0);await click('#node-insert');check('node insert modifies editable curve topology',await js(`__studio.engine.selected.path.length>${count}`));
    await click('#undo');await idle();await click('#redo');await idle();
    await penPath();await change('style-fill','#dfa947');await rename(await selectedId(),'自由な曲線');
  });
  await step('font outline and pixel conversion',async()=>{
    const original=await text('編集できるロゴ文字','MIDORI',{x:174,y:745},100,'#315d4e',true);
    await click('#duplicate');const copy=await selectedId();await position({x:174,y:745});
    await select([original]);await js(`document.querySelector('[data-layer="'+${json(original)}+'"]').querySelector('button').click()`);await select([copy]);
    const before=await js(`(()=>{const r=__studio.engine.selected.getBoundingRect();return{left:r.left,top:r.top,width:r.width,height:r.height};})()`);
    await click('#outline-text');await waitUntil('!!__studio.engine.selected?.dakeOutline',60000);
    const after=await js(`(()=>{const r=__studio.engine.selected.getBoundingRect();return{left:r.left,top:r.top,width:r.width,height:r.height};})()`);
    check('outline produces editable paths retaining source text metadata',await js(`__studio.engine.selected.type.toLowerCase()==='path'&&__studio.engine.selected.dakeOutline.originalText==='MIDORI'`),{before,after});
    const badge=await shape('ellipse','画像化したアクセント',{x:725,y:130,width:105,height:105},'#dfa947');
    await click('#rasterize');await waitUntil(`__studio.engine.selected?.type.toLowerCase()==='image'`);check('rasterize UI generates actual image layer',await js(`__studio.engine.selected.toObject().src.startsWith('data:image/')`));
  });
  await guides(100,100);await baseline('logo');
}
async function makeCard() {
  await newWork('card', '緑野 花 / 名刺');
  const mark=await drop(path.join(output,'source-mark.svg'),115,150);await position({x:115,y:150,width:175,height:175});await rename(mark,'読み込んだSVGロゴ');
  await text('名前','緑野 花',{x:385,y:175},68,'#254e3d',true);
  await text('肩書き','デザインと暮らしの道具',{x:389,y:305},24,'#5b7464');
  await text('連絡先','midori@example.invalid\n架空市みどり町 1-2-3',{x:135,y:510},26,'#385b49');
  const bar=await shape('rect','下部ライン',{x:120,y:640,width:890,height:7},'#d9b55f');
  await guides(10,10);await baseline('card');
  check('business card full canvas includes 3mm bleed and inner safe region',await js(`Math.abs(__studio.engine.width/300*25.4-97)<.06&&Math.abs(__studio.engine.height/300*25.4-61)<.06&&Math.abs(__studio.engine.meta.bleed*25.4/300-3)<.001&&Math.abs(__studio.engine.meta.safe*25.4/300-4)<.001`));
}
async function exportFormat(kind,format){
  saveTarget=path.join(output,`${kind}.${format==='jpeg'?'jpg':format}`);
  const start=Date.now();await click('#export');await change('export-format',format);await click('#dialog-export');
  const end=Date.now()+60000;while(!fs.existsSync(saveTarget)&&Date.now()<end)await delay(50);
  check(`${format} native export creates a real nonempty file`,fs.existsSync(saveTarget)&&fs.statSync(saveTarget).size>100);
  const bytes=fs.readFileSync(saveTarget);const result={path:saveTarget,bytes:bytes.length,sha256:hash(bytes),ms:Date.now()-start};
  if(format==='pdf'){
    check('PDF output has PDF header and page MediaBox',bytes.subarray(0,5).toString()==='%PDF-'&&bytes.includes(Buffer.from('/MediaBox')));
  }else if(format==='svg'){
    const svg=bytes.toString('utf8');check('SVG contains vector paths and no external resource reference',svg.includes('<svg')&&/<(?:path|text|rect|ellipse)\b/.test(svg)&&!/(?:href|xlink:href)=["']https?:/i.test(svg));
    result.dimensions=await js(`(async()=>{const image=new Image();image.src='data:image/svg+xml;base64,'+${json(bytes.toString('base64'))};await image.decode();return{width:image.naturalWidth,height:image.naturalHeight};})()`);
  }else{
    result.dimensions=await js(`(async()=>{const image=new Image();image.src='data:image/${format};base64,'+${json(bytes.toString('base64'))};await image.decode();return{width:image.naturalWidth,height:image.naturalHeight};})()`);
    check(`${format} decoder confirms document pixel dimensions`,result.dimensions.width===report.works[kind].width&&result.dimensions.height===report.works[kind].height);
  }
  return result;
}
function canonical(value) {
  if (typeof value === 'number') return Math.round(value * 1e8) / 1e8;
  if (typeof value === 'string' && value.length > 2000) return { sha256: hash(Buffer.from(value)) };
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === 'object') return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])]));
  return value;
}
function pixelComparison(before, after) {
  const a = nativeImage.createFromBuffer(before), b = nativeImage.createFromBuffer(after);
  assert.deepEqual(a.getSize(), b.getSize());
  const left = a.toBitmap(), right = b.toBitmap();
  let maximum = 0, sum = 0, changed = 0; const bounds = [Infinity, Infinity, -1, -1], width = a.getSize().width;
  for (let i = 0; i < left.length; i += 4) {
    let different = false;
    for (let channel = 0; channel < 4; channel++) { const value = Math.abs(left[i + channel] - right[i + channel]); maximum = Math.max(maximum, value); if (channel < 3) sum += value; different ||= value !== 0; }
    if (different) { changed++; const p=i/4,x=p%width,y=Math.floor(p/width);bounds[0]=Math.min(bounds[0],x);bounds[1]=Math.min(bounds[1],y);bounds[2]=Math.max(bounds[2],x);bounds[3]=Math.max(bounds[3],y); }
  }
  const pixels = left.length / 4;
  return { byteIdentical: before.equals(after), differenceBounds: changed ? bounds : null, maximumChannelDifference: maximum, changedPixels: changed, totalPixels: pixels, changedRatio: changed / pixels, meanRGBDifference: sum / (pixels * 3), threshold: { maxChannel: 3, changedRatio: .02, meanRGB: .02 } };
}
async function resumeWorks(){
  const prior=JSON.parse(fs.readFileSync(path.join(output,'prepare.json'),'utf8'));report.sources=prior.sources;
  check('resume uses a different Electron process',prior.processId!==process.pid);
  check('resume starts with fresh user profile and no font cache',!fs.existsSync(path.join(process.env.STADIO_USER_DATA,'fonts')));
  for(const kind of ['banner','flyer','logo','card']){
    activeWork=kind;report.works[kind]={...prior.works[kind],exports:{}};
    await step('open embedded-font document with native network disabled',async()=>{
      openTargets=[path.join(output,kind+'.dake')];await js(`document.activeElement?.blur()`);win.webContents.sendInputEvent({type:'keyDown',keyCode:'O',modifiers:['control']});win.webContents.sendInputEvent({type:'keyUp',keyCode:'O',modifiers:['control']});
      await waitUntil(`__studio.currentPath===${json(openTargets[0])}&&!__studio.dirty`,60000);await documentFontsReady();
      check('document metadata survives process restart',json(await js('structuredClone(__studio.engine.meta)'))===json(prior.works[kind].meta));
      const raster=await js(`__studio.engine.exportRaster('png',1)`);const bytes=Buffer.from(raster.split(',')[1],'base64');fs.writeFileSync(path.join(output,kind+'-reopened.png'),bytes);
      const pixels=pixelComparison(fs.readFileSync(path.join(output,kind+'-before.png')),bytes);report.works[kind].reopenPixels=pixels;
      if(kind==='card'){
        pixels.threshold={maxChannel:8,changedRatio:.001,meanRGB:.001};pixels.allowedReason='Only imported SVG circle edge antialiasing; geometry and text separately compared';
        const b=pixels.differenceBounds;check('card pixel differences are confined to imported SVG circle edge',!b||(b[0]>=225&&b[1]>=148&&b[2]<=275&&b[3]<=212),b);
      }
      check('offline reopen preserves rendered pixels within recorded interpolation tolerance',pixels.maximumChannelDifference<=pixels.threshold.maxChannel&&pixels.changedRatio<pixels.threshold.changedRatio&&pixels.meanRGBDifference<pixels.threshold.meanRGB,pixels);
      const saved=JSON.parse(fs.readFileSync(openTargets[0],'utf8')),reopened=await js('__studio.engine.serialize()');
      check('all serialized text font metadata and geometry survive restart (1e-8 numeric precision)',json(canonical({canvas:saved.canvas,fonts:saved.fonts,meta:saved.meta}))===json(canonical({canvas:reopened.canvas,fonts:reopened.fonts,meta:reopened.meta})));
      check('embedded Japanese font is loaded without network',await js(`document.fonts.check('16px "Noto Sans JP"')&&__studio.engine.fonts.some(f=>f.family==='Noto Sans JP'&&f.license)`));
    });
    await step('re-edit text and vector after restart then save and export',async()=>{
      const textId=await js(`__studio.engine.layers.find(o=>o.type.toLowerCase()==='textbox'&&o.visible)?.id`);
      if(textId){await select([textId]);const existing=await js('__studio.engine.selected.text');await change('prop-text',existing+'\n');await change('prop-tracking',.035);await change('prop-leading',1.4);}
      else {const id=await js('__studio.engine.layers.find(o=>o.visible&&!o.locked)?.id');await select([id]);const x=await js('__studio.engine.selected.left');await change('prop-x',x+8);}
      check('post-restart property edits mark document dirty',await js('__studio.dirty'));
      await click('#undo');await idle();await click('#redo');await idle();await saveProject(kind,'-edited');await deselect();
      for(const format of ['png','jpeg','webp','svg','pdf'])report.works[kind].exports[format]=await exportFormat(kind,format);
      await screenshot(kind+'-edited-screen.png');
    });
  }
  check('all imported source files remain byte-identical',Object.entries(report.sources).every(([file,before])=>hash(fs.readFileSync(path.join(output,file)))===before));
  check('offline project reopening performs no external request',report.network.length===0,report.network);
}
async function prepareAbsentFontSmoke() {
  const installed=await js('window.stadio.listLocalFonts()');
  const known=new Set(installed.flatMap(f=>String(f.family).split(',').map(x=>x.trim().replace(/^["']|["']$/g,''))));
  const catalog=JSON.parse(fs.readFileSync(path.join(root,'assets/google-fonts-catalog.json'),'utf8')).families;
  const candidates=['Zen Kaku Gothic New','Shippori Mincho','Zen Maru Gothic','Klee One'];
  const family=candidates.find(name=>!known.has(name)&&catalog.some(f=>f.family===name&&f.subsets.includes('japanese')));
  check('uninstalled Japanese family chosen against actual Windows font inventory',!!family,{family,installedCount:installed.length});
  await newWork('custom','未導入フォントの再開確認',{width:1000,height:600,unit:'px',dpi:96,bleed:0,safe:32});
  await click('[data-tool="text"]');await change('prop-text','はじめての書体。\n手元に残る、日本語。');await change('prop-fontsize',54);
  await chooseGoogleFont(family,700);const smokeTextId=await selectedId();
  await click('#choose-font');await click('#font-local');await change('font-search','Noto Sans JP');
  await js(`(()=>{const row=[...document.querySelectorAll('.font-row')].find(row=>row.querySelector('strong')?.textContent==='Noto Sans JP');if(!row)throw new Error('Expected installed Noto family');row.querySelector('button').click();})()`);
  await waitUntil(`__studio.engine.selected?.fontFamily==='Noto Sans JP'`);await click('#dialog-close');
  check('switch to local font prunes unused embedded descriptor and dynamic face',await js(`!__studio.engine.fonts.some(f=>f.family===${json(family)})&&![...document.fonts].some(f=>f.family.replace(/"/g,'')===${json(family)})`));
  await click('#undo');await waitUntil(`__studio.engine.layers.find(o=>o.id===${json(smokeTextId)})?.fontFamily===${json(family)}`);await select([smokeTextId]);await documentFontsReady();
  report.fontUndoDebug=await js(`({text:__studio.engine.selected.fontFamily,fonts:__studio.engine.fonts.map(f=>({family:f.family,sha256:f.sha256,weight:f.weight,style:f.style})),faces:[...document.fonts].map(f=>({family:f.family.replace(/"/g,''),status:f.status,weight:f.weight,style:f.style}))})`);
  check('font undo restores embedded descriptor and loaded dynamic face',report.fontUndoDebug.fonts.some(f=>f.family===family)&&report.fontUndoDebug.faces.some(f=>f.family===family&&f.status==='loaded'),report.fontUndoDebug);
  await click('#redo');await waitUntil(`__studio.engine.layers.find(o=>o.id===${json(smokeTextId)})?.fontFamily==='Noto Sans JP'`);await select([smokeTextId]);await click('#undo');await waitUntil(`__studio.engine.layers.find(o=>o.id===${json(smokeTextId)})?.fontFamily===${json(family)}`);await select([smokeTextId]);await documentFontsReady();
  check('font redo and second undo remain reversible',await js(`__studio.engine.fonts.some(f=>f.family===${json(family)})`));
  await change('prop-tracking',.03);await change('prop-leading',1.4);await change('prop-style','italic');
  check('text style selector changes selected font style',await js("__studio.engine.selected.fontStyle==='italic'"));await change('prop-style','normal');
  await position({x:80,y:170});await color('#254e3d');
  await shape('rect','上部ライン',{x:80,y:95,width:780,height:6},'#d9b55f');await shape('ellipse','アクセント',{x:835,y:445,width:65,height:65},'#83ad87');await deselect();
  const document=await saveProject('font-smoke');await documentFontsReady();const png=await js("__studio.engine.exportRaster('png',1)");
  fs.writeFileSync(path.join(output,'font-smoke-before.png'),Buffer.from(png.split(',')[1],'base64'));
  report.absentFont={family,installedCount:installed.length,installed:false,sha256:document.fonts.find(f=>f.family===family)?.sha256,faces:await js(`[...document.fonts].map(f=>({family:f.family.replace(/"/g,''),status:f.status,weight:f.weight,style:f.style}))`)};
  check('previously absent font is a loaded actual FontFace',report.absentFont.faces.some(f=>f.family===family&&f.status==='loaded'));
  check('new document prunes previous Noto dynamic faces',!report.absentFont.faces.some(f=>f.family==='Noto Sans JP'));
  delete report.works.custom;await screenshot('font-smoke-screen.png');
}
async function resumeAbsentFontSmoke() {
  const prior=JSON.parse(fs.readFileSync(path.join(output,'prepare.json'),'utf8'));const {family,sha256}=prior.absentFont;activeWork='absent-font';
  const installed=await js('window.stadio.listLocalFonts()');check('font remains absent from Windows after earlier download',!installed.some(f=>String(f.family).split(',').map(x=>x.trim()).includes(family)));
  openTargets=[path.join(output,'font-smoke.dake')];await js('window.__studio.open()');await waitUntil(`__studio.currentPath===${json(openTargets[0])}&&!__studio.dirty`,60000);await documentFontsReady();
  const descriptor=await js(`__studio.engine.fonts.find(f=>f.family===${json(family)})`);check('fresh offline process restores identical embedded font and license',descriptor?.sha256===sha256&&descriptor.license.length>1000);
  const faces=await js(`[...document.fonts].map(f=>({family:f.family.replace(/"/g,''),status:f.status}))`);check('offline embedded uninstalled family loads actual FontFace',faces.some(f=>f.family===family&&f.status==='loaded'));
  check('loading another document prunes previous dynamic faces',!faces.some(f=>f.family==='Noto Sans JP'));
  const png=Buffer.from((await js("__studio.engine.exportRaster('png',1)")).split(',')[1],'base64');fs.writeFileSync(path.join(output,'font-smoke-reopened.png'),png);
  const pixels=pixelComparison(fs.readFileSync(path.join(output,'font-smoke-before.png')),png);report.absentFont={family,sha256,faces,pixels};check('uninstalled embedded font reproduces pixels offline',pixels.maximumChannelDifference<=3&&pixels.changedRatio<=.02&&pixels.meanRGBDifference<=.02,pixels);
  const textId=await js("__studio.engine.layers.find(o=>o.type.toLowerCase()==='textbox').id");await select([textId]);await change('prop-text','再開後にも、編集できる。');await saveProject('font-smoke','-edited');check('offline uninstalled font remains editable',JSON.parse(fs.readFileSync(saveTarget,'utf8')).canvas.objects.some(o=>o.text==='再開後にも、編集できる。'));
  check('offline font smoke performs no external request',report.network.length===0,report.network);
}
async function additionalWorkflowSafety() {
  activeWork='safety';
  const before=await js('JSON.stringify(__studio.engine.serialize())');
  for (const [field,value] of [['new-dpi',0],['new-safe',999999]]) {
    await js('document.activeElement?.blur()');win.webContents.sendInputEvent({type:'keyDown',keyCode:'N',modifiers:['control']});win.webContents.sendInputEvent({type:'keyUp',keyCode:'N',modifiers:['control']});
    await waitUntil(`!!document.getElementById('new-name')`);await change(field,value);await click('#dialog-create');
    await waitUntil(`!!document.getElementById('dialog-close')`);
    check('invalid new-document '+field+' preserves complete current document',await js('JSON.stringify(__studio.engine.serialize())')===before);
    await click('#dialog-close');
  }
  const dirty=await js('__studio.dirty');
  await js('document.activeElement?.blur()');win.webContents.sendInputEvent({type:'keyDown',keyCode:'Tab'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Tab'});await delay(200);
  check('Tab hides inspector without dirtying artwork',await js(`document.getElementById('app').classList.contains('panels-hidden')&&__studio.dirty===${dirty}`));
  win.webContents.sendInputEvent({type:'keyDown',keyCode:'Tab'});win.webContents.sendInputEvent({type:'keyUp',keyCode:'Tab'});await delay(200);
  await js(`window.__qaPointer=[];for(const type of ['pointerdown','pointermove','pointerup','mousedown','mousemove','mouseup'])document.addEventListener(type,e=>window.__qaPointer.push({type,target:e.target.id,x:e.clientX,y:e.clientY,buttons:e.buttons,pointerId:e.pointerId}),{capture:true});`);
  const handle=await js(`(()=>{const r=document.getElementById('panel-resizer').getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+100};})()`);
  await mouse('mouseMoved',handle);await delay(50);await mouse('mousePressed',handle);await delay(50);await mouse('mouseMoved',{x:handle.x-40,y:handle.y},{buttons:1});await delay(50);await mouse('mouseReleased',{x:handle.x-40,y:handle.y});await delay(250);
  report.resizerDebug={handle,events:await js("window.__qaPointer"),after:await js(`({dirty:__studio.dirty,inlineWidth:document.documentElement.style.getPropertyValue('--panel-width'),hidden:document.getElementById('app').classList.contains('panels-hidden'),hit:document.elementFromPoint(${handle.x},${handle.y})?.outerHTML?.slice(0,300)})`)};
  check('inspector resizing leaves document clean',await js(`__studio.dirty===${dirty}&&!!document.documentElement.style.getPropertyValue('--panel-width')`));
  win.webContents.send('stadio:menu','toggleRulers');await delay(100);check('native view command toggles rulers without dirtying artwork',await js(`document.querySelector('.workarea').classList.contains('no-rulers')&&__studio.dirty===${dirty}`));
  win.webContents.send('stadio:menu','resetWorkspace');await delay(250);check('workspace reset changes view only',await js(`!document.querySelector('.workarea').classList.contains('no-rulers')&&!document.getElementById('app').classList.contains('panels-hidden')&&__studio.dirty===${dirty}`));
  const target=path.join(output,'banner.dake'),point=await scenePoint(200,200),data={items:[],files:[target],dragOperationsMask:1};
  for(const type of ['dragEnter','dragOver','drop'])await command('Input.dispatchDragEvent',{type,...point,data});
  await waitUntil(`__studio.currentPath===${json(target)}&&!__studio.dirty`,60000);
  check('actual CDP .dake drop restores current path and clean state',await js(`__studio.engine.name==='暮らしを、軽やかに。'&&!__studio.dirty`));
}
async function documentFontsReady(){await js('document.fonts.ready.then(()=>true)');await delay(150);}
async function run(){
  await app.whenReady();win=BrowserWindow.getAllWindows()[0];win.setContentSize(1520,980);
  win.webContents.on('console-message',event=>{if(event.level==='error')report.errors.push(event.message);});
  if(win.webContents.isLoading())await new Promise(resolve=>win.webContents.once('did-finish-load',resolve));
  await waitUntil('!!window.__studio?.production && !!window.stadio');
  win.webContents.debugger.attach('1.3');await command('Page.enable');await command('Runtime.enable');
  if(phase!=='prepare')session.defaultSession.enableNetworkEmulation({offline:true});
  if(phase==='prepare'){
    await prepareSources();
    if(process.env.STADIO_FONT_ONLY!=='1'){await makeBanner();await makeFlyer();await makeLogo();await makeCard();}await prepareAbsentFontSmoke();
    check('Google font binary and license were downloaded via actual native request',report.network.some(item=>/raw\.githubusercontent\.com.*\.(?:ttf|otf)(?:\?|$)/i.test(decodeURIComponent(item.url)))&&report.network.some(item=>/raw\.githubusercontent\.com.*(?:OFL|LICENSE|LICENCE|UFL)\.txt/i.test(item.url)),report.network);
    check('font requests never contain artwork text or CSS text subset query',report.network.every(item=>!/[?&]text=/i.test(item.url)&&!item.url.includes(encodeURIComponent('暮らし'))));
  }else if(phase==='resume') { await resumeWorks(); await resumeAbsentFontSmoke(); await additionalWorkflowSafety(); }
  else {openTargets=[path.join(output,'card.dake')];await js('window.__studio.open()');await waitUntil(`__studio.currentPath===${json(openTargets[0])}&&!__studio.dirty`,60000);await additionalWorkflowSafety();}
  activeWork='';report.expectedErrors=report.errors.filter(message=>/INVALID_LAYOUT|INVALID_DIMENSIONS|INVALID_DOCUMENT/.test(message));report.unexpectedErrors=report.errors.filter(message=>!report.expectedErrors.includes(message));check('renderer reports no unexpected errors',report.unexpectedErrors.length===0,report.unexpectedErrors);
  report.passed=true;report.finished=new Date().toISOString();fs.writeFileSync(path.join(output,phase+'.json'),JSON.stringify(report,null,2));
  console.log(JSON.stringify({phase,passed:true,checks:report.checks.length,processId:process.pid,output}));
  await js('window.stadio.discardRecovery()');win.close();
}
run().catch(error=>{report.passed=false;report.failure=error.stack;try{fs.writeFileSync(path.join(output,phase+'.json'),JSON.stringify(report,null,2));}catch{}console.error(error);app.exit(1);});
