import { cache } from 'fabric';
import UI_TEXT from './ui-text.json';
const FONT_TEXT=UI_TEXT.fontReliability;
import { missingFontGlyphs } from './font-coverage.js';
import { validateDocument } from './engine.js';
import { normalizeMeta } from './engine-foundation.js';

export function createProductionUI(ctx) {
  const { engine, T, bridge, run, modal, btn, field, esc, fit, status, setTool, renderProps, renderLayers, ensureSaved, clearPath } = ctx;
  const $=q=>document.querySelector(q), $$=q=>[...document.querySelectorAll(q)];
  const presets={banner:{width:1600,height:600,unit:'px',dpi:96,bleed:0,safe:32},flyer:{width:216,height:303,unit:'mm',dpi:300,bleed:3,safe:5},logo:{width:1000,height:1000,unit:'px',dpi:96,bleed:0,safe:40,transparent:true},card:{width:97,height:61,unit:'mm',dpi:300,bleed:3,safe:4}};
  const view={panels:true,rulers:true,guides:true,grid:false};
  let cachedFonts=[],localFonts=[],googleFonts=[],loadedFonts=new Map(),dragLayer=null,overlayFrame=0;
  const toPx=(n,unit=engine.meta?.unit,dpi=engine.meta?.dpi||96)=>unit==='mm'?Number(n)*dpi/25.4:Number(n);
  const fromPx=(n,unit=engine.meta?.unit,dpi=engine.meta?.dpi||96)=>unit==='mm'?Number(n)*25.4/dpi:Number(n);
  const rounded=n=>Math.round(n*100)/100;
  const familyNames=value=>String(value||'').split(',').map(f=>f.trim().replace(/^["']|["']$/g,''));
  const isText=o=>['textbox','text','i-text','itext'].includes(o?.type?.toLowerCase());
  const walk=(objects,fn)=>objects.forEach(o=>{fn(o);if(o.getObjects)walk(o.getObjects(),fn);});
  function refreshFonts(){cache.clearFontCache();walk(engine.layers,o=>{if(isText(o)){o.initDimensions();o.setCoords();o.dirty=true;}});engine.canvas.requestRenderAll();}
  const fontKey=d=>[d.family,String(d.weight||400),d.style||'normal',d.sha256||d.postscriptName||d.id].join('|');
  const cleanDescriptor=d=>{const copy={...d};if(copy.source==='local'){delete copy.data;delete copy.fontBytes;}delete copy.cached;delete copy.cacheKey;return copy;};
  const missingGlyphs=(d,text)=>missingFontGlyphs(d,text)||[];
  let fontSetup=Promise.resolve(),fontApplySequence=0;
  async function resolveFace(descriptor){
    if(descriptor.source==='local'){
      if(!localFonts.length)localFonts=await bridge.listLocalFonts();
      const found=localFonts.find(f=>f.id===descriptor.id||f.postscriptName===descriptor.postscriptName)||localFonts.find(f=>f.family===descriptor.family&&Number(f.weight||400)===Number(descriptor.weight||400)&&(f.style||'normal')===(descriptor.style||'normal'));
      if(!found)throw new Error('fontUnavailable');
      const key=fontKey({...found,sha256:descriptor.sha256});const registered=loadedFonts.get(key);
      if(registered?.descriptor?.data)return registered.descriptor;
      return bridge.getLocalFont(found.id);
    }
    if(descriptor.data)return descriptor;
    if(descriptor.cacheKey)return bridge.getCachedFont(descriptor.cacheKey);
    cachedFonts=await bridge.listCachedFonts();const found=cachedFonts.find(f=>f.id===descriptor.id)||cachedFonts.find(f=>f.family===descriptor.family&&String(f.weight)===String(descriptor.weight||400)&&(f.style||'normal')===(descriptor.style||'normal'))||cachedFonts.find(f=>{const range=String(f.weightRange||'').split(' ').map(Number);return f.family===descriptor.family&&(f.style||'normal')===(descriptor.style||'normal')&&range.length===2&&Number(descriptor.weight)>=range[0]&&Number(descriptor.weight)<=range[1];});
    if(!found)throw new Error('fontUnavailable');const cached=await bridge.getCachedFont(found.cacheKey);const range=String(cached.weightRange||'').split(' ').map(Number);return range.length===2&&Number(descriptor.weight)>=range[0]&&Number(descriptor.weight)<=range[1]?{...cached,weight:Number(descriptor.weight)}:cached;
  }
  function activateFace(entry){
    const range=d=>{const values=String(d.weightRange||d.weight||400).split(' ').map(Number);return [values[0],values[1]||values[0]];};
    const a=entry.descriptor,ar=range(a);let changed=false;
    for(const other of loadedFonts.values()){if(other===entry)continue;const b=other.descriptor,br=range(b);if(a.family===b.family&&(a.style||'normal')===(b.style||'normal')&&a.sha256!==b.sha256&&ar[0]<=br[1]&&br[0]<=ar[1])changed=document.fonts.delete(other.face)||changed;}
    if(!document.fonts.has(entry.face)){document.fonts.add(entry.face);changed=true;}return changed;
  }
  async function installFace(input,activate=true){
    const descriptor=await resolveFace(input);if(!descriptor?.data||!descriptor.family)throw new Error('fontUnavailable');
    const key=fontKey(descriptor);let entry=loadedFonts.get(key);
    if(!entry){const face=new FontFace(descriptor.family,'url('+descriptor.data+')',{weight:String(descriptor.weightRange||descriptor.weight||400),style:descriptor.style||'normal'});await face.load();if(face.status!=='loaded')throw new Error('fontUnavailable');entry={face,descriptor,key};loadedFonts.set(key,entry);}
    if(activate)activateFace(entry);return entry.descriptor;
  }
  async function restoreFonts(data){
    validateDocument(data);const unavailable=[];
    for(const font of data.fonts||[])try{await installFace(font);}catch(error){if(font.source==='google'&&font.data)throw error;unavailable.push(font.family);}
    await document.fonts.ready;
    if(!localFonts.length&&bridge?.listLocalFonts)try{localFonts=await bridge.listLocalFonts();}catch{}
    const known=new Set([...localFonts.map(f=>f.family),...[...loadedFonts.values()].map(x=>x.descriptor.family)]);
    const inspect=objects=>objects.forEach(o=>{if(isText(o)&&!familyNames(o.fontFamily).some(f=>known.has(f)||['sans-serif','serif','monospace'].includes(f)))unavailable.push(o.fontFamily);if(o.objects)inspect(o.objects);});inspect(data.canvas.objects);
    if(unavailable.length)queueMicrotask(()=>{status('fontMissing');ctx.toast(T.fontMissing+': '+[...new Set(unavailable)].join(', '));});
  }
  async function loadDocument(data){
    const before=new Map(loadedFonts);
    try{await engine.loadDocument(data,{prepare:()=>restoreFonts(data)});refreshFonts();hideWelcome();}
    catch(error){for(const [key,value]of loadedFonts)if(value!==before.get(key))document.fonts.delete(value.face);loadedFonts=before;syncFontSources();refreshFonts();throw error;}
  }
  function syncFontSources(){
    // The library belongs to the app session, not to a single document. Keep loaded
    // faces available for new layers and for another document; .dake stores only dependencies.
    let changed=false;for(const descriptor of engine.fonts||[]){const entry=loadedFonts.get(fontKey(descriptor));if(entry)changed=activateFace(entry)||changed;}if(changed)refreshFonts();
  }
  async function prepareNewText(){await fontSetup;return engine.fontSelection||null;}
  function rememberDefault(descriptor){
    engine.fontSelection=cleanDescriptor(descriptor);
    const metadata={...engine.fontSelection};delete metadata.data;delete metadata.license;delete metadata.coverage;
    try{localStorage.setItem('stadio.defaultFont.v1',JSON.stringify(metadata));}catch{}
  }
  async function useFont(input){
    const sequence=++fontApplySequence,selectionAtStart=engine.selected,revisionAtStart=engine._revision;
    const unchanged=()=>sequence===fontApplySequence&&engine.selected===selectionAtStart&&engine._revision===revisionAtStart;
    const descriptor=await installFace(input,false);if(!unchanged())throw new Error('DOCUMENT_CHANGED');
    const target=isText(engine.selected)&&!engine.selected.locked?engine.selected:null;
    let next=[...(engine.fonts||[])].filter(f=>!(f.family===descriptor.family&&String(f.weight||400)===String(descriptor.weight||400)&&(f.style||'normal')===(descriptor.style||'normal')));
    next.push(cleanDescriptor(descriptor));
    const used=new Set([descriptor.family.toLocaleLowerCase()]);walk(engine.layers,o=>{if(isText(o))for(const family of familyNames(o===target?descriptor.family:o.fontFamily))used.add(family.toLocaleLowerCase());});next=next.filter(f=>used.has(f.family.toLocaleLowerCase()));
    try{validateDocument({...engine.serialize(),fonts:next});}catch{throw new Error('fontCapacity');}
    if(!unchanged())throw new Error('DOCUMENT_CHANGED');activateFace(loadedFonts.get(fontKey(descriptor)));engine.fonts=next;rememberDefault(descriptor);cache.clearFontCache(descriptor.family);
    if(target)engine.updateSelected({fontFamily:descriptor.family,fontWeight:descriptor.weight||400,fontStyle:descriptor.style||'normal'});else engine.commit('font');
    refreshFonts();renderProps();const missing=missingGlyphs(descriptor,target?.text||T.fontPreviewText);
    status(missing.length?FONT_TEXT.fontAppliedFallback:target?T.fontApplied:FONT_TEXT.defaultReady);
    if(missing.length)ctx.toast(FONT_TEXT.glyphFallback+missing.slice(0,16).join(''));
    return {descriptor,missing};
  }
  async function useVariant(partial){
    const target=engine.selected,revision=engine._revision;if(!isText(target))return;
    const family=target.fontFamily,weight=Number(partial.weight??(target.fontWeight==='bold'?700:target.fontWeight==='normal'?400:target.fontWeight||400)),style=partial.style||target.fontStyle||'normal';
    const stored=(engine.fonts||[]).find(f=>f.family===family);
    let descriptor;
    if(stored?.source==='google'){
      const range=String(stored.weightRange||stored.weight||400).split(' ').map(Number);
      if(stored.style===style&&range.length===2&&weight>=range[0]&&weight<=range[1])descriptor={...stored,weight,id:'google:'+stored.family+':'+style+':'+weight};
      else {cachedFonts=await bridge.listCachedFonts();const cached=cachedFonts.find(f=>f.family===family&&Number(f.weight)===weight&&(f.style||'normal')===style);descriptor=cached?await bridge.getCachedFont(cached.cacheKey):await bridge.downloadGoogleFont({family,weight,style,requestId:'variant-'+crypto.randomUUID()});}
    }else {if(!localFonts.length)localFonts=await bridge.listLocalFonts();descriptor=localFonts.find(f=>f.family===family&&Number(f.weight)===weight&&(f.style||'normal')===style);}
    if(!descriptor||Number(descriptor.weight)!==weight||(descriptor.style||'normal')!==style)throw new Error('fontUnavailable');
    if(engine.selected!==target||engine._revision!==revision)throw new Error('DOCUMENT_CHANGED');await useFont(descriptor);
  }
  async function fontLibrary(){
    const editedText=isText(engine.selected)?engine.selected.text:T.fontPreviewText;
    const done=modal('<h1>'+T.fontLibrary+'</h1><p>'+T.fontEmbeddedNote+'</p><div class="font-tabs">'+btn('font-local',T.fontLocal,'secondary')+btn('font-google',T.fontGoogle,'secondary')+btn('font-cached',T.fontCached,'secondary')+'</div>'+field('font-search',T.fontSearch,'','search','autocomplete="off"')+field('font-preview-text',T.fontPreview,editedText,'text')+'<label class="check"><input id="font-japanese" type="checkbox">'+T.fontJapanese+'</label><div id="download-variants" class="fieldgrid" hidden><label class="field"><span>'+T.fontWeight+'</span><select id="font-download-weight">'+[100,200,300,400,500,600,700,800,900].map(w=>'<option value="'+w+'" '+(w===400?'selected':'')+'>'+w+'</option>').join('')+'</select></label><label class="field"><span>'+T.fontStyle+'</span><select id="font-download-style"><option value="normal">'+T.fontNormal+'</option><option value="italic">'+T.fontItalic+'</option></select></label></div><p id="font-progress" role="status">'+T.fontLoading+'</p><div id="font-results" class="font-results"></div><p class="dialog-note">'+T.fontNetworkNote+'</p>',[{id:'close',text:T.close}],true);
    let mode='local',closed=false,downloadId=null,drawSequence=0;const previewFaces=[],rowDescriptors=new Map();
    const reportText=text=>{if(!closed&&$('#font-progress'))$('#font-progress').textContent=text;};
    done.then(()=>{closed=true;drawSequence++;for(const face of previewFaces)document.fonts.delete(face);if(downloadId)bridge?.cancelFontDownload?.(downloadId);});
    function updatePreview(row,d){const sample=$('#font-preview-text')?.value||'';row.querySelector('.font-preview').textContent=sample;const missing=missingGlyphs(d,sample);row.querySelector('.font-preview-status').textContent=missing.length?FONT_TEXT.glyphFallback+missing.slice(0,16).join(''):FONT_TEXT.previewReady;}
    const draw=()=>{
      if(closed)return;const ownSequence=++drawSequence;rowDescriptors.clear();$('#font-japanese').closest('label').hidden=mode!=='google';$('#download-variants').hidden=mode!=='google';for(const key of ['local','google','cached'])$('#font-'+key).classList.toggle('active',mode===key);
      const query=$('#font-search').value.toLocaleLowerCase(),onlyJapanese=$('#font-japanese').checked;const list=(mode==='local'?localFonts:mode==='cached'?cachedFonts:googleFonts).filter(f=>f.family.toLocaleLowerCase().includes(query)&&(!onlyJapanese||mode!=='google'||f.subsets?.includes('japanese')));
      $('#font-results').replaceChildren();reportText(list.length?list.length+' '+T.font:T.fontNoResults);
      for(const f of list.slice(0,60)){
        const row=document.createElement('div');row.className='font-row';row.innerHTML='<div><strong>'+esc(f.family)+'</strong><small>'+esc(f.nativeStyle||((f.style==='italic'?T.fontItalic:T.fontNormal)+' '+(f.weight||'')))+'</small><div class="font-preview"></div><small class="font-preview-status">'+FONT_TEXT.previewUnverified+'</small></div><div><button class="secondary font-preview-button">'+(mode==='google'?FONT_TEXT.downloadPreview:FONT_TEXT.previewLoad)+'</button><button class="secondary font-use-button">'+T.fontUse+'</button></div>';
        const preview=row.querySelector('.font-preview'),previewButton=row.querySelector('.font-preview-button'),useButton=row.querySelector('.font-use-button');preview.textContent='—';let prepared=null,loading=null;
        const prepare=()=>loading||(loading=(async()=>{
          previewButton.disabled=true;useButton.disabled=true;let ownDownload=null;reportText(FONT_TEXT.stageFetch);
          try{let d;if(mode==='google'){
            if(downloadId)throw Error('BUSY');const id='preview-'+crypto.randomUUID();downloadId=id;ownDownload=id;
            d=await bridge.downloadGoogleFont({family:f.family,id:f.id||f.family,requestId:id,weight:Number($('#font-download-weight').value),style:$('#font-download-style').value});cachedFonts=await bridge.listCachedFonts();
          }else d=await resolveFace(f);
          if(closed||ownSequence!==drawSequence)return null;reportText(FONT_TEXT.stageRegister);
          const alias='stadio-preview-'+crypto.randomUUID(),face=new FontFace(alias,'url('+d.data+')',{weight:String(d.weightRange||d.weight||400),style:d.style||'normal'});await face.load();
          if(closed||ownSequence!==drawSequence)return null;document.fonts.add(face);previewFaces.push(face);preview.style.fontFamily='"'+alias+'"';preview.style.fontWeight=String(d.weight||400);preview.style.fontStyle=d.style||'normal';prepared=d;rowDescriptors.set(row,d);row.dataset.preview='ready';updatePreview(row,d);reportText(FONT_TEXT.previewReady);return d;
          }catch(error){if(!closed&&ownSequence===drawSequence){row.dataset.preview='failed';row.querySelector('.font-preview-status').textContent=FONT_TEXT.previewFailed;reportText(ctx.errorText(error)===T.genericError?T.fontUnavailable:ctx.errorText(error));}throw error;
          }finally{if(ownDownload&&downloadId===ownDownload)downloadId=null;loading=null;previewButton.disabled=false;useButton.disabled=false;}
        })());
        previewButton.onclick=()=>prepare().catch(()=>{});
        useButton.onclick=async()=>{const target=engine.selected,revision=engine._revision;try{const d=prepared||await prepare();if(!d||closed||ownSequence!==drawSequence)return;if(engine.selected!==target||engine._revision!==revision)throw new Error('DOCUMENT_CHANGED');const result=await useFont(d);reportText(result.missing.length?FONT_TEXT.fontAppliedFallback:FONT_TEXT.defaultReady);}catch(error){reportText(ctx.errorText(error)===T.genericError?T.fontUnavailable:ctx.errorText(error));}};
        if(mode!=='google')row.onpointerenter=()=>{if(!prepared&&!loading)prepare().catch(()=>{});};
        $('#font-results').append(row);
      }
    };
    const load=async(next)=>{if(downloadId)return;mode=next;reportText(T.fontLoading);try{if(mode==='local'&&!localFonts.length)localFonts=await bridge.listLocalFonts();if(mode==='cached')cachedFonts=await bridge.listCachedFonts();if(mode==='google'&&!googleFonts.length){const result=await bridge.listGoogleFonts();googleFonts=Array.isArray(result)?result:result.fonts||result.families||[];}draw();}catch{reportText(T.fontUnavailable);}};
    $('#font-preview-text').oninput=()=>{for(const [row,d]of rowDescriptors)updatePreview(row,d);};
    $('#font-search').oninput=()=>{if(!downloadId)draw();};$('#font-japanese').onchange=()=>{if(!downloadId)draw();};
    for(const key of ['local','google','cached'])$('#font-'+key).onclick=()=>load(key);
    for(const id of ['font-download-weight','font-download-style'])$('#'+id).onchange=()=>{if(!downloadId)draw();};
    const unsubscribe=bridge?.onFontProgress?.(p=>{if(!closed&&downloadId&&p.id===downloadId)reportText(T.fontDownloading+' '+(p.total?Math.round(p.loaded/p.total*100)+'%':Math.round(p.loaded/1024)+' KB'));});
    await load('local');await done;unsubscribe?.();
  }
  async function outlineText(){
    const object=engine.selected;if(!isText(object))return;const families=familyNames(object.fontFamily);
    let descriptor=(engine.fonts||[]).find(f=>families.includes(f.family)&&(String(f.weight)===String(object.fontWeight)||String(f.weight).includes(' ')));
    if(!descriptor)descriptor=(engine.fonts||[]).find(f=>families.includes(f.family));
    if(!descriptor){if(!localFonts.length)localFonts=await bridge.listLocalFonts();const candidates=localFonts.filter(f=>families.includes(f.family));const weight=object.fontWeight==='bold'?700:Number(object.fontWeight)||400;const target=candidates.filter(f=>f.style===(object.fontStyle||'normal')).sort((a,b)=>Math.abs(a.weight-weight)-Math.abs(b.weight-weight))[0]||candidates[0];if(!target)throw new Error('FONT_UNSUPPORTED');descriptor=await bridge.getLocalFont(target.id);}
    if(descriptor.source==='local'&&!descriptor.data)descriptor=await resolveFace(descriptor);
    const encoded=descriptor.data||descriptor.fontBytes;if(!encoded)throw new Error('FONT_UNSUPPORTED');const bytes=typeof encoded==='string'?Uint8Array.from(atob(encoded.includes(',')?encoded.split(',')[1]:encoded),c=>c.charCodeAt(0)):new Uint8Array(encoded);
    await engine.outlineText(bytes.buffer,descriptor);
  }
  async function newDocument(preset='banner'){
    if(!await ensureSaved())return;const p=presets[preset]||presets.banner;
    const pending=modal(`<h1>${T.newTitle}</h1><p>${T.newDescription}</p>${field('new-name',T.name,T.untitled,'text','maxlength="100"')}<label class="field"><span>${T.preset}</span><select id="new-preset">${Object.keys(presets).map(k=>`<option value="${k}" ${k===preset?'selected':''}>${T['preset'+k[0].toUpperCase()+k.slice(1)]}</option>`).join('')}<option value="custom">${T.presetCustom}</option></select></label><div class="fieldgrid">${field('new-width',T.width,p.width,'number','min="1" step="0.1"')}${field('new-height',T.height,p.height,'number','min="1" step="0.1"')}<label class="field"><span>${T.unit}</span><select id="new-unit"><option value="px" ${p.unit==='px'?'selected':''}>px</option><option value="mm" ${p.unit==='mm'?'selected':''}>mm</option></select></label>${field('new-dpi',T.dpi,p.dpi,'number','min="36" max="1200"')}${field('new-bleed',T.bleed,p.bleed,'number','min="0" step="0.1"')}${field('new-safe',T.safe,p.safe,'number','min="0" step="0.1"')}</div><div class="colorrow">${field('new-background',T.background,'#ffffff','color')}<label class="check"><input id="new-transparent" type="checkbox" ${p.transparent?'checked':''}>${T.transparent}</label></div><p>${T.fullCanvasNote}</p><p>${T.limits}</p>`,[{id:'cancel',text:T.cancel},{id:'create',text:T.create,cls:'primary'}]);
    $('#new-preset').onchange=()=>{const v=presets[$('#new-preset').value];if(v){for(const k of ['width','height','unit','dpi','bleed','safe'])$('#new-'+k).value=v[k];$('#new-transparent').checked=!!v.transparent;}};
    const result=await pending;if(result.id!=='create')return;const v=result.values,unit=v['new-unit'],dpi=Number(v['new-dpi']);
    const width=Math.round(toPx(v['new-width'],unit,dpi)),height=Math.round(toPx(v['new-height'],unit,dpi));
    if(!Number.isInteger(width)||!Number.isInteger(height)||width<1||height<1||width>8192||height>8192||width*height>32000000)throw new Error('invalidSize');
    const metadata=normalizeMeta({dpi,unit,bleed:toPx(v['new-bleed'],unit,dpi),safe:toPx(v['new-safe'],unit,dpi)},width,height);
    await engine.newDocument({width,height,name:v['new-name'].trim()||T.untitled,background:v['new-transparent']?'transparent':v['new-background'],meta:metadata});
    clearPath();setTool('select');hideWelcome();fit();
  }
  async function documentSettings(){
    const m=engine.meta,unit=m.unit;
    const pending=modal(`<h1>${T.documentSettings}</h1><p>${engine.width} × ${engine.height} px</p><div class="fieldgrid"><label class="field"><span>${T.unit}</span><select id="layout-unit"><option value="px" ${unit==='px'?'selected':''}>px</option><option value="mm" ${unit==='mm'?'selected':''}>mm</option></select></label>${field('layout-dpi',T.dpi,m.dpi,'number','min="36" max="1200"')}${field('layout-bleed',T.bleed+' ('+unit+')',rounded(fromPx(m.bleed)),'number','min="0" step="0.1"')}${field('layout-safe',T.safe+' ('+unit+')',rounded(fromPx(m.safe)),'number','min="0" step="0.1"')}</div><p>${T.fullCanvasNote}</p><div class="checkgrid">${[['enabled','snap'],['objects','snapObjects'],['artboard','snapArtboard'],['guides','snapGuides'],['grid','snapGrid']].map(([k,label])=>`<label class="check"><input id="snap-${k}" type="checkbox" ${m.snap[k]?'checked':''}>${T[label]}</label>`).join('')}</div>${field('grid-size',T.gridSize+' ('+unit+')',rounded(fromPx(m.snap.gridSize)),'number','min="0.1" step="0.1"')}<p>${T.guideHint}</p>`,[{id:'cancel',text:T.cancel},{id:'apply',text:T.apply,cls:'primary'}],true);
    const r=await pending;if(r.id!=='apply')return;const v=r.values,newUnit=v['layout-unit'],dpi=Number(v['layout-dpi']);
    engine.updateDocumentSettings({dpi,unit:newUnit,bleed:toPx(v['layout-bleed'],unit,m.dpi),safe:toPx(v['layout-safe'],unit,m.dpi),snap:Object.fromEntries([['enabled',v['snap-enabled']],['objects',v['snap-objects']],['artboard',v['snap-artboard']],['guides',v['snap-guides']],['grid',v['snap-grid']],['gridSize',toPx(v['grid-size'],unit,m.dpi)]])});renderProps();scheduleOverlay();
  }
  function hideWelcome(){if($('.welcome-shell'))$('.welcome-shell').hidden=true;}
  function togglePanels(){view.panels=!view.panels;$('#app').classList.toggle('panels-hidden',!view.panels);fit();}
  function toggleRulers(){view.rulers=!view.rulers;$('.workarea').classList.toggle('no-rulers',!view.rulers);scheduleOverlay();}
  function toggleGuides(){view.guides=!view.guides;engine.setView({guidesVisible:view.guides});scheduleOverlay();}
  function resetWorkspace(){view.panels=view.rulers=view.guides=true;view.grid=false;$('#app').classList.remove('panels-hidden');$('.workarea').classList.remove('no-rulers');document.documentElement.style.removeProperty('--panel-width');engine.setView({guidesVisible:true,gridVisible:false});fit();}
  function scheduleOverlay(){if(overlayFrame)return;overlayFrame=requestAnimationFrame(()=>{overlayFrame=0;drawOverlay();});}
  function drawOverlay(){
    const canvas=$('#layout-overlay');if(!canvas)return;const host=$('.canvas-host').getBoundingClientRect();canvas.width=Math.round(host.width);canvas.height=Math.round(host.height);const c=canvas.getContext('2d');const v=engine.canvas.viewportTransform,z=engine.zoom,m=engine.meta;if(!m)return;
    const line=(axis,pos,color,dash=[])=>{const p=(axis==='x'?v[4]:v[5])+pos*z;c.strokeStyle=color;c.setLineDash(dash);c.beginPath();if(axis==='x'){c.moveTo(p,0);c.lineTo(p,canvas.height);}else{c.moveTo(0,p);c.lineTo(canvas.width,p);}c.stroke();};
    if(engine.view?.gridVisible){const size=m.snap.gridSize;const startX=Math.max(0,Math.floor(-v[4]/z/size)*size),endX=Math.min(engine.width,(canvas.width-v[4])/z),startY=Math.max(0,Math.floor(-v[5]/z/size)*size),endY=Math.min(engine.height,(canvas.height-v[5])/z);if(size*z>=5){for(let x=startX;x<=endX;x+=size)line('x',x,'#a1b9b433',[2,3]);for(let y=startY;y<=endY;y+=size)line('y',y,'#a1b9b433',[2,3]);}}
    if(engine.view?.guidesVisible!==false){for(const p of m.guides.vertical)line('x',p,'#66c8d3bb');for(const p of m.guides.horizontal)line('y',p,'#66c8d3bb');for(const [inset,color]of [[m.bleed,'#ee9daf'],[m.bleed+m.safe,'#90ca98']])if(inset>0){c.strokeStyle=color;c.setLineDash([6,4]);c.strokeRect(v[4]+inset*z,v[5]+inset*z,(engine.width-2*inset)*z,(engine.height-2*inset)*z);}}
    for(const l of engine.view?.snapLines||[])line(l.axis,l.position,'#f7b7ef',[3,2]);c.setLineDash([]);
    if(view.rulers){c.fillStyle='#20282a';c.fillRect(0,0,canvas.width,22);c.fillRect(0,0,22,canvas.height);c.font='10px "Yu Gothic UI"';c.fillStyle='#b7c5c8';c.strokeStyle='#6b777b';const factor=toPx(1),rawStep=65/(z*factor),base=10**Math.floor(Math.log10(rawStep)),step=[1,2,5,10].map(n=>n*base).find(n=>n>=rawStep)*factor;
      for(const axis of ['x','y']){const offset=axis==='x'?v[4]:v[5],limit=axis==='x'?canvas.width:canvas.height;for(let p=Math.ceil(-offset/z/step)*step;p*z+offset<limit;p+=step){const pos=p*z+offset;if(pos<22)continue;c.beginPath();if(axis==='x'){c.moveTo(pos,16);c.lineTo(pos,22);c.fillText(String(rounded(fromPx(p))),pos+3,12);}else{c.moveTo(16,pos);c.lineTo(22,pos);c.save();c.translate(11,pos+3);c.rotate(-Math.PI/2);c.fillText(String(rounded(fromPx(p))),0,0);c.restore();}c.stroke();}}
      c.fillStyle='#29363a';c.fillRect(0,0,22,22);c.fillStyle='#d4e1e3';c.fillText(m.unit,3,15);
    }
  }
  function enhanceProps(){
    if(!engine.meta)return;const s=engine.selected,props=$('#props'),type=s?.type?.toLowerCase(),multi=engine.canvas.getActiveObjects().length>1;
    const section=(title,html,open=true)=>{const d=document.createElement('details');d.className='production-section';d.open=open;d.innerHTML=`<summary>${esc(title)}</summary>${html}`;props.append(d);return d;};
    const bind=(id,fn,event='click')=>{const el=$('#'+id);if(el)el.addEventListener(event,()=>run(()=>fn(el)));};
    if(isText(s)){
      const base=$('#prop-text')?.closest('.props-section');if(base)props.prepend(base);
      const old=$('#prop-font');if(old){const control=document.createElement('input');control.id='prop-font';control.type='text';control.value=s.fontFamily;control.setAttribute('list','available-fonts');control.setAttribute('aria-label',T.font);old.replaceWith(control);const list=document.createElement('datalist');list.id='available-fonts';const families=[...new Set([...localFonts.map(f=>f.family),...cachedFonts.map(f=>f.family),...(engine.fonts||[]).map(f=>f.family),s.fontFamily])];list.innerHTML=families.map(f=>`<option value="${esc(f)}"></option>`).join('');control.after(list);control.onchange=()=>run(async()=>{const family=control.value;const descriptor=(engine.fonts||[]).find(f=>f.family===family)||cachedFonts.find(f=>f.family===family)||localFonts.find(f=>f.family===family);if(!descriptor)throw new Error('fontUnavailable');await useFont(descriptor);});}
      base?.insertAdjacentHTML('beforeend',`${btn('choose-font',T.fontLibrary,'widebtn')}<div class="fieldgrid">${field('prop-tracking',T.letterSpacing,(s.charSpacing||0)/1000,'number','min="-0.2" max="2" step="0.01"')}${field('prop-leading',T.lineHeight,s.lineHeight||1.16,'number','min="0.5" max="4" step="0.05"')}<label class="field"><span>${T.fontWeight}</span><select id="prop-weight">${[100,200,300,400,500,600,700,800,900].map(w=>`<option ${String(s.fontWeight)===String(w)||(s.fontWeight==='normal'&&w===400)||(s.fontWeight==='bold'&&w===700)?'selected':''}>${w}</option>`).join('')}</select></label><label class="field"><span>${T.fontStyle}</span><select id="prop-style"><option value="normal">${T.fontNormal}</option><option value="italic" ${s.fontStyle==='italic'?'selected':''}>${T.fontItalic}</option></select></label></div><label class="check"><input type="checkbox" id="prop-underline" ${s.underline?'checked':''}>${T.underline}</label>${btn('outline-text',T.textOutline,'widebtn')}<p class="small-note">${T.outlineNote}</p>`);
      $('#prop-bold')?.closest('.check')?.remove();const justify=document.createElement('option');justify.value='justify';justify.textContent=T.textJustify;$('#prop-align')?.append(justify);if($('#prop-align'))$('#prop-align').value=s.textAlign;
      bind('choose-font',fontLibrary);bind('outline-text',outlineText);bind('prop-tracking',e=>engine.updateSelected({charSpacing:Number(e.value)*1000}),'change');bind('prop-leading',e=>engine.updateSelected({lineHeight:Number(e.value)}),'change');bind('prop-weight',e=>useVariant({weight:Number(e.value)}),'change');bind('prop-style',e=>useVariant({style:e.value}),'change');bind('prop-underline',e=>engine.updateSelected({underline:e.checked}),'change');
    }
    if(s){
      const transform=$('#prop-x')?.closest('.props-section');if(transform){const alignButtons=['align-left','align-center','align-top','align-middle'];for(const id of alignButtons)$('#'+id)?.parentElement?.remove();const collapse=document.createElement('details');collapse.className='production-section';collapse.open=!isText(s);collapse.innerHTML=`<summary>${T.position}</summary>`;transform.replaceWith(collapse);collapse.append(transform);}
      section(T.layout,`<p class="small-note">${T.alignSelectionNote}</p><div class="align-grid">${[['left','alignLeftShort'],['center','alignCenterShort'],['right','alignRight'],['top','alignTopShort'],['middle','alignMiddleShort'],['bottom','alignBottom']].map(([k,key])=>btn('align-'+k,T[key],'secondary')).join('')}</div><div class="minirow">${btn('distribute-h',T.distributeHorizontal)}${btn('distribute-v',T.distributeVertical)}</div>`,multi);
      for(const k of ['left','center','right','top','middle','bottom'])bind('align-'+k,()=>engine.alignSelected(k));bind('distribute-h',()=>engine.distributeSelected('horizontal'));bind('distribute-v',()=>engine.distributeSelected('vertical'));
      $('#distribute-h').disabled=$('#distribute-v').disabled=engine.canvas.getActiveObjects().length<3;
      if(type!=='image'&&!isText(s)){
        section(T.vector,`${btn('convert-path',T.convertPath,'widebtn')}<div class="align-grid">${['union','subtract','intersect','divide','compound'].map(k=>btn('vector-'+k,T[k],'secondary')).join('')}${btn('vector-split',T.splitCompound,'secondary')}</div><p class="small-note">${T.splitNote}</p>${type==='path'?`${field('node-index',T.nodeIndex,engine.activeNodeIndex||0,'number','min="0" step="1"')}<div class="align-grid">${['insert','remove','smooth','corner','close','open'].map(k=>btn('node-'+k,T['node'+k[0].toUpperCase()+k.slice(1)],'secondary')).join('')}</div>`:''}`,true);
        bind('convert-path',()=>engine.convertSelectedToPath());for(const k of ['union','subtract','intersect','divide','compound']){bind('vector-'+k,()=>engine.booleanSelected(k));$('#vector-'+k).disabled=!multi;}bind('vector-split',()=>engine.splitCompound());for(const k of ['insert','remove','smooth','corner','close','open'])bind('node-'+k,()=>engine.editNode(k,Number($('#node-index')?.value||0)));
      }
      if(multi||s.dakeMask||(s.clipPath&&!s.dakeComponent)){
      section(T.mask,`<p class="small-note">${T.maskNote}</p>${btn('mask-create',T.maskCreate,'widebtn')}<div class="minirow">${btn('mask-toggle',T.maskToggle)}${btn('mask-release',T.maskRelease)}</div>`,multi||!!s.clipPath||!!s.dakeMask);
      bind('mask-create',()=>engine.createMask());bind('mask-toggle',()=>engine.toggleMask());bind('mask-release',()=>engine.releaseMask());$('#mask-create').disabled=!multi;$('#mask-toggle').disabled=$('#mask-release').disabled=!s.clipPath&&!s.dakeMask;
      }
      const dpi=engine.getEffectiveDPI?.(s);if(dpi){const min=typeof dpi==='number'?dpi:dpi.min;section(T.effectiveDpi,`<p>${Math.round(min)} dpi${min<150?' · '+T.lowDpi:''}</p>`,true);}
    }else{
      const m=engine.meta;section(T.layout,`${btn('document-settings',T.documentSettings,'widebtn')}<div class="checkgrid"><label class="check"><input id="quick-snap" type="checkbox" ${m.snap.enabled?'checked':''}>${T.snap}</label><label class="check"><input id="quick-grid" type="checkbox" ${engine.view?.gridVisible?'checked':''}>${T.showGrid}</label></div>`);
      bind('document-settings',documentSettings);bind('quick-snap',e=>engine.updateDocumentSettings({snap:{...m.snap,enabled:e.checked}}),'change');bind('quick-grid',e=>engine.setView({gridVisible:e.checked}),'change');
      const guides=section(T.guides,`<p class="small-note">${T.guideHint}</p><div class="fieldgrid"><label class="field"><span>${T.guideAdd}</span><select id="guide-axis"><option value="y">${T.guideHorizontal}</option><option value="x">${T.guideVertical}</option></select></label>${field('guide-pos',T.guidePosition+' ('+m.unit+')',0,'number','step="0.1"')}</div>${btn('guide-add',T.guideAdd,'widebtn')}<div class="guide-list"></div>`);
      bind('guide-add',()=>engine.addGuide($('#guide-axis').value,toPx($('#guide-pos').value)));
      for(const [axis,values]of [['x',m.guides.vertical],['y',m.guides.horizontal]])values.forEach((value,index)=>{const row=document.createElement('div');row.className='guide-row';row.innerHTML=`<span>${axis==='x'?T.vertical:T.horizontal}</span><input aria-label="${T.guidePosition}" type="number" min="0" step="0.1" value="${rounded(fromPx(value))}"><span>${m.unit}</span><button title="${T.guideRemove}" aria-label="${T.guideRemove}">×</button>`;row.querySelector('input').onchange=e=>run(()=>{const key=axis==='x'?'vertical':'horizontal',next=[...engine.meta.guides[key]];next[index]=toPx(e.target.value);engine.updateDocumentSettings({guides:{[key]:next}});renderProps();});row.querySelector('button').onclick=()=>run(()=>engine.removeGuide(axis,index));guides.querySelector('.guide-list').append(row);});
    }
    if(s){
      const unit=engine.meta.unit;
      for(const [id,key,value]of [['x','left',s.left],['y','top',s.top],['width','scaleX',s.getScaledWidth()],['height','scaleY',s.getScaledHeight()]]){
        const control=$('#prop-'+id);if(!control)continue;control.value=rounded(fromPx(value));control.step=unit==='mm'?'0.1':'1';const label=control.closest('label')?.querySelector('span');if(label)label.textContent=(id==='x'?'X':id==='y'?'Y':T[id])+' ('+unit+')';
        control.onchange=()=>run(()=>{const n=toPx(control.value);if(!Number.isFinite(n))throw new Error('INVALID_LAYOUT');if(id==='width'||id==='height'){const extent=id==='width'?s.getScaledWidth():s.getScaledHeight();engine.updateSelected({[key]:s[key]*Math.max(1,Math.min(32768,n))/extent});}else engine.updateSelected({[key]:n});renderLayers();});
      }
      const fontSize=$('#prop-fontsize')?.closest('label')?.querySelector('span');if(fontSize)fontSize.textContent=T.fontSize+' (px)';
    }
    if(s?.locked)props.querySelectorAll('input,textarea,select,button').forEach(el=>el.disabled=true);
  }
  function enhanceLayers(){
    for(const row of $$('.layer-row')){
      const id=row.dataset.layer;row.draggable=true;row.setAttribute('role','option');row.setAttribute('aria-selected',row.classList.contains('active'));
      row.onclick=e=>run(()=>engine.select(id,{additive:e.shiftKey||e.ctrlKey}));
      row.ondblclick=e=>{e.preventDefault();const layer=engine.layers.find(o=>o.id===id);engine.select(id);if(isText(layer)&&!layer.locked){setTool('select');layer.enterEditing();layer.selectAll();}else if(layer?.type?.toLowerCase()==='path')setTool('node');else $('#prop-width')?.focus();};
      row.ondragstart=e=>{dragLayer=id;e.dataTransfer.setData('application/x-dake-layer',id);e.dataTransfer.effectAllowed='move';};
      row.ondragover=e=>{if(dragLayer){e.preventDefault();row.classList.add('drag-target');}};row.ondragleave=()=>row.classList.remove('drag-target');
      row.ondrop=e=>{e.preventDefault();e.stopPropagation();row.classList.remove('drag-target');if(dragLayer){const index=engine.layers.findIndex(o=>o.id===id);run(()=>engine.moveLayer(dragLayer,index));dragLayer=null;}};row.ondragend=()=>{dragLayer=null;};
    }
  }
  async function dropped(event){
    event.preventDefault();$('.workarea').classList.remove('drop-active');if(dragLayer||!event.dataTransfer.files.length)return;
    const files=[...event.dataTransfer.files];
    await run(async()=>{
      const data=await bridge.importDroppedFiles(files);
      if(data?.kind==='project'||data?.data?.format==='dake-stadio'){if(!await ensureSaved())return;if(ctx.openDocument)await ctx.openDocument(data.data,data.path);else{await loadDocument(data.data);ctx.documentOpened(data.path);}hideWelcome();fit();return;}
      const rect=$('.canvas-host').getBoundingClientRect(),v=engine.canvas.viewportTransform;const x=(event.clientX-rect.left-v[4])/engine.zoom,y=(event.clientY-rect.top-v[5])/engine.zoom;
      for(const [i,file]of (Array.isArray(data)?data:data.files||[]).entries()){
        if(file.kind==='project'){if(!await ensureSaved())return;if(ctx.openDocument)await ctx.openDocument(file.data,file.path);else{await loadDocument(file.data);ctx.documentOpened(file.path);}fit();continue;}
        if(file.kind==='svg')await engine.importSVG(file.data,file.name);else await engine.importImage(file.data,file.name);
        if(engine.selected){const b=engine.selected.getBoundingRect();engine.updateSelected({left:engine.selected.left+x-b.left+i*24,top:engine.selected.top+y-b.top+i*24});}
      }
      hideWelcome();setTool('select');status('importDone');
    });
  }
  function setup(){
    const area=$('.workarea');area.insertAdjacentHTML('beforeend','<canvas id="layout-overlay" aria-hidden="true"></canvas><div class="ruler-hit ruler-top"></div><div class="ruler-hit ruler-left"></div>');
    for(const [selector,axis]of [['.ruler-top','x'],['.ruler-left','y']])$(selector).onclick=e=>run(()=>{if(engine.meta.guideLocked){status('guideLock');return;}const r=$('.canvas-host').getBoundingClientRect(),v=engine.canvas.viewportTransform;engine.addGuide(axis,Math.round(((axis==='x'?e.clientX-r.left-v[4]:e.clientY-r.top-v[5])/engine.zoom)*100)/100);});
    area.addEventListener('dragover',e=>{if(e.dataTransfer.types.includes('Files')){e.preventDefault();area.classList.add('drop-active');}});area.addEventListener('dragleave',e=>{if(!area.contains(e.relatedTarget))area.classList.remove('drop-active');});area.addEventListener('drop',dropped);
    document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
    engine.canvas.on('after:render',scheduleOverlay);
    for(const b of $$('[data-preset]'))b.onclick=()=>run(()=>newDocument(b.dataset.preset));
    $('#close-start').onclick=hideWelcome;$('#font-library').onclick=()=>run(fontLibrary);$('#layout-settings').onclick=()=>run(documentSettings);$('#panels-toggle').onclick=togglePanels;
    const resizer=$('#panel-resizer');resizer.onpointerdown=e=>{
      e.preventDefault();e.stopPropagation();
      const move=ev=>{document.documentElement.style.setProperty('--panel-width',Math.max(280,Math.min(440,window.innerWidth-ev.clientX))+'px');};
      const finish=()=>{window.removeEventListener('pointermove',move,true);window.removeEventListener('pointerup',finish,true);window.removeEventListener('pointercancel',finish,true);window.removeEventListener('blur',finish,true);fit();};
      window.addEventListener('pointermove',move,true);window.addEventListener('pointerup',finish,true);window.addEventListener('pointercancel',finish,true);window.addEventListener('blur',finish,true);
    };
    fontSetup=(async()=>{try{if(bridge?.listLocalFonts)localFonts=await bridge.listLocalFonts();}catch{}try{if(bridge?.listCachedFonts)cachedFonts=await bridge.listCachedFonts();const saved=JSON.parse(localStorage.getItem('stadio.defaultFont.v1')||'null');if(saved)engine.fontSelection=cleanDescriptor(await installFace(saved));}catch{status(FONT_TEXT.defaultUnavailable);}if(!$('#props').contains(document.activeElement))renderProps();})();
    scheduleOverlay();
  }
  return {setup,prepareNewText,syncFontSources,newDocument,documentSettings,fontLibrary,outlineText,restoreFonts,loadDocument,useFont,enhanceProps,enhanceLayers,hideWelcome,togglePanels,toggleRulers,toggleGuides,resetWorkspace,scheduleOverlay,dropped,presets,toPx,fromPx,get view(){return view;}};
}


