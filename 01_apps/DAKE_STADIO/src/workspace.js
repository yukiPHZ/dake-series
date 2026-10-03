import { StudioEngine } from './engine.js';
import { componentFonts } from './engine-components.js';

const copy=value=>structuredClone(value);
const uid=()=>crypto.randomUUID();
export function createWorkspace(ctx) {
  const {engine,T,bridge,run,modal,btn,esc,production}=ctx;
  const sessions=[];let active=null,switching=false,previewSequence=0,previewTimer,previewRunning=false,previewAgain=false;
  const baseNew=engine.newDocument.bind(engine),baseLoad=engine.loadDocument.bind(engine);
  const meta=()=>({path:null,dirty:false,changeEpoch:0,savedEpoch:0,documentEpoch:0});
  function history(){return {items:[...engine._history],index:engine._historyIndex,images:new Map(engine._images),tokens:new Map(engine._imageTokens),next:engine._nextImageToken};}
  function capture(){
    if(!active||switching)return;
    active.data=engine.serialize();active.meta=ctx.captureMeta();active.history=history();
    active.viewport=engine.canvas.viewportTransform.slice();active.view=copy(engine.view);active.fontSelection=copy(engine.fontSelection||null);
  }
  function makeSession(data,options={}){
    return {id:uid(),data:copy(data),meta:{...meta(),...options},history:null,viewport:null,child:null};
  }
  function changed(event){
    if(switching)return true;
    if(!active)return false;
    if(!['busy','ready','viewport','view','node-selection','tool'].includes(event.reason)) {
      active.data=engine.serialize();
      if(active.child){previewSequence++;clearTimeout(previewTimer);previewTimer=setTimeout(()=>renderParentPreview(),140);}
    }
    renderTabs();return false;
  }
  function syncMeta(){if(active&&!switching){active.meta=ctx.captureMeta();renderTabs();}bridge?.setDirty?.(sessions.some(s=>s.meta.dirty));}
  async function transition(task){
    if(switching)throw new Error('BUSY');
    await ctx.finishEditing?.();capture();switching=true;ctx.beginSwitch?.();
    const revision=engine._revision;
    try{return await task();}
    finally{engine._revision=Math.max(revision,engine._revision)+1;switching=false;ctx.endSwitch?.();renderTabs();ctx.afterSwitch?.();syncMeta();}
  }
  async function activate(id){
    const next=sessions.find(s=>s.id===id);if(!next||next===active)return;
    await transition(async()=>{
      const old=active;
      try{await production.loadDocument(next.data);active=next;
        if(next.history){engine._history=[...next.history.items];engine._historyIndex=next.history.index;engine._images=new Map(next.history.images);engine._imageTokens=new Map(next.history.tokens);engine._nextImageToken=next.history.next;}
        if(next.viewport)engine.canvas.setViewportTransform(next.viewport);if(next.view)engine.view=copy(next.view);
        engine.fontSelection=copy(next.fontSelection||engine.fontSelection||null);ctx.restoreMeta(next.meta);
      }catch(error){active=old;throw error;}
    });
    renderChildBar();if(active.child)renderParentPreview();
  }
  async function openDocument(data,{path=null,dirty=false}={}){
    const existing=findByPath(path);if(existing){await activate(existing.id);return existing;}
    if(sessions.length>=16)throw new Error('WORKSPACE_LIMIT');
    await transition(async()=>{await production.loadDocument(data);const s=makeSession(engine.serialize(),{path,dirty,documentEpoch:Date.now()});s.history=history();sessions.push(s);active=s;ctx.restoreMeta(s.meta);});
    renderChildBar();return active;
  }
  async function newDocument(options){
    if(sessions.length>=16)throw new Error('WORKSPACE_LIMIT');
    await transition(async()=>{await baseNew(options);const s=makeSession(engine.serialize(),{dirty:true,documentEpoch:Date.now()});s.history=history();sessions.push(s);active=s;ctx.restoreMeta(s.meta);});
    renderChildBar();return active;
  }
  async function askClose(session){
    if(session.child)return finishChild(false);
    if(session.meta.dirty){
      const answer=await modal('<h1>'+T.workspaceCloseTitle+'</h1><p>'+esc(session.data.name||T.untitled)+'</p>',[{id:'cancel',text:T.cancel},{id:'discard',text:T.discard},{id:'save',text:T.save,cls:'primary'}]);
      if(answer.id==='cancel')return false;
      if(answer.id==='save'){await activate(session.id);if(!await ctx.save())return false;}
    }
    return true;
  }
  async function close(id){
    capture();const s=sessions.find(s=>s.id===id);if(!s)return false;
    for(const child of sessions.filter(c=>c.child?.parentId===id)){await activate(child.id);if(!await finishChild(false))return false;}
    if(s.child){await activate(s.id);return finishChild(false);}
    if(!await askClose(s))return false;
    const i=sessions.indexOf(s);
    if(active===s){const next=sessions.find(c=>c!==s);if(next)await activate(next.id);else{await newDocument({width:1200,height:800,name:T.untitled,background:'#ffffff'});ctx.restoreMeta({...meta(),documentEpoch:Date.now()});syncMeta();}}
    sessions.splice(i,1);renderTabs();return true;
  }
  const pathKey=value=>String(value||'').replace(/\\/g,'/').replace(/\/$/,'').toLowerCase();
  function findByPath(path){capture();if(!path)return null;const key=pathKey(path);return sessions.find(s=>!s.child&&pathKey(s.meta.path)===key)||null;}
  function referencesPath(data,target){
    let found=false;const walk=objects=>{for(const o of objects||[]){const c=o.dakeComponent;if(c){if(pathKey(c.originPath||c.link?.path)===target)found=true;walk(c.document?.canvas?.objects);}else walk(o.objects);}};walk(data.canvas.objects);return found;
  }
  function guardSource(source){
    let owner=active;const seen=new Set();
    while(owner&&!seen.has(owner.id)){seen.add(owner.id);const target=pathKey(owner.meta.path);if(target&&(target===pathKey(source.originPath||source.link?.path)||referencesPath(source.data,target)))throw new Error('COMPONENT_CYCLE');owner=owner.child?sessions.find(s=>s.id===owner.child.parentId):null;}
  }
  function recordSaved(id,path,data){
    if(active?.id===id)capture();const session=sessions.find(s=>s.id===id);if(!session)return;
    session.meta.path=path;const equal=JSON.stringify(session.data)===JSON.stringify(data);if(equal){session.meta.dirty=false;session.meta.savedEpoch=session.meta.changeEpoch;}renderTabs();return equal;
  }
  async function place(mode='embedded'){
    const source=await bridge?.importComponent?.(mode);if(!source)return false;guardSource(source);
    await production.restoreFonts({...source.data,fonts:componentFonts(source.data)});
    await engine.insertComponent(source.data,{mode,link:source.link,originPath:source.originPath,name:source.name});capture();return true;
  }
  async function editComponent(layer=engine.selected){
    if(!layer?.dakeComponent)throw new Error('COMPONENT_MISSING');
    const parent=active;capture();
    const existing=sessions.find(s=>s.child?.parentId===parent.id&&s.child.layerId===layer.id);
    if(existing){await activate(existing.id);return;}
    const component=copy(layer.dakeComponent);
    await openDocument(component.document,{dirty:false});
    active.child={parentId:parent.id,layerId:layer.id};renderChildBar();renderTabs();await renderParentPreview();
  }
  async function finishChild(apply){
    if(!active?.child)return true;
    await ctx.finishEditing?.();capture();const child=active,parent=sessions.find(s=>s.id===child.child.parentId);
    if(!parent||!parent.data.canvas.objects.some(o=>o.id===child.child.layerId&&o.dakeComponent)){child.child=null;child.meta.path=null;child.meta.dirty=true;ctx.restoreMeta(child.meta);renderChildBar();renderTabs();ctx.status('COMPONENT_MISSING');return true;}
    if(!apply&&child.meta.dirty){
      const result=await modal('<h1>'+T.componentDiscardTitle+'</h1><p>'+T.componentDraftNote+'</p>',[{id:'cancel',text:T.cancel},{id:'discard',text:T.discard},{id:'apply',text:T.componentApply,cls:'primary'}]);
      if(result.id==='cancel')return false;apply=result.id==='apply';
    }
    await activate(parent.id);
    if(apply){
      await production.restoreFonts({...engine.serialize(),fonts:[...engine.fonts,...componentFonts(child.data)].filter((f,i,a)=>a.findIndex(g=>[g.family,g.weight,g.style,g.sha256].join('|')===[f.family,f.weight,f.style,f.sha256].join('|'))===i)});
      await engine.replaceComponent(child.child.layerId,child.data,{commit:true});
      capture();
    }
    sessions.splice(sessions.indexOf(child),1);renderTabs();renderChildBar();syncMeta();return true;
  }
  async function resolveDrafts(sessionId=active?.id){
    for(const draft of sessions.filter(s=>s.child?.parentId===sessionId)){
      await activate(draft.id);const answer=await modal('<h1>'+T.componentPendingTitle+'</h1><p>'+T.componentDraftNote+'</p>',[{id:'cancel',text:T.cancel},{id:'discard',text:T.discard},{id:'apply',text:T.componentApply,cls:'primary'}]);
      if(answer.id==='cancel')return false;
      if(answer.id==='discard'){draft.meta.dirty=false;ctx.restoreMeta({...ctx.captureMeta(),dirty:false});}
      if(!await finishChild(answer.id==='apply'))return false;
    }
    return true;
  }
  async function refreshComponent(){
    const layer=engine.selected;if(layer?.dakeComponent?.mode!=='linked')throw new Error('COMPONENT_MISSING');
    let source;
    try{source=await bridge.refreshComponent(layer.dakeComponent.link.token);}
    catch(error){
      if(!/COMPONENT_RELINK|ENOENT|ENOTDIR/.test(String(error.message)))throw error;
      const answer=await modal('<h1>'+T.componentRelinkTitle+'</h1><p>'+esc(layer.dakeComponent.link.path)+'</p><p>'+T.componentRelinkNote+'</p>',[{id:'cancel',text:T.cancel},{id:'select',text:T.componentRelink,cls:'primary'}]);
      if(answer.id!=='select')return;source=await bridge.importComponent('linked');
    }
    if(!source)return;guardSource(source);
    await production.restoreFonts({...source.data,fonts:componentFonts(source.data)});
    await engine.replaceComponent(layer.id,source.data,{link:source.link});ctx.status('componentRefreshed');
  }
  async function embedComponent(){const s=engine.selected;if(!s?.dakeComponent)return;await engine.replaceComponent(s.id,s.dakeComponent.document,{mode:'embedded'});}
  async function renderParentPreview(){
    if(!active?.child)return;
    if(previewRunning){previewAgain=true;return;}
    const child=active,parent=sessions.find(s=>s.id===child.child.parentId);if(!parent)return;
    const serial=++previewSequence,doc=copy(parent.data);
    const layer=doc.canvas.objects.find(o=>o.id===child.child.layerId);if(!layer)return;
    layer.dakeComponent.document=engine.serialize();doc.fonts=componentFonts(doc);
    previewRunning=true;const canvas=document.createElement('canvas'),renderer=new StudioEngine(canvas);
    try{
      await renderer.loadDocument(doc);
      const scale=Math.min(320/doc.width,180/doc.height);
      const image=await renderer.exportRaster('png',scale);
      if(serial!==previewSequence||active!==child)return;
      const img=document.querySelector('#component-parent-preview');if(img)img.src=image;
    }catch(error){if(serial===previewSequence)ctx.status('componentPreviewFailed');console.error(error);}
    finally{await renderer.dispose();previewRunning=false;if(previewAgain){previewAgain=false;setTimeout(()=>renderParentPreview(),0);}}
  }
  function renderTabs(){
    const bar=document.querySelector('#document-tabs');if(!bar)return;
    const wanted=new Set(sessions.map(s=>s.id));
    for(const element of [...bar.children])if(!wanted.has(element.dataset.session))element.remove();
    for(const session of sessions){
      let row=[...bar.children].find(e=>e.dataset.session===session.id);
      if(!row){row=document.createElement('div');row.className='document-tab';row.dataset.session=session.id;row.innerHTML='<button class="document-tab-name" role="tab"></button><button class="document-tab-close">×</button>';bar.append(row);row.firstChild.onclick=()=>run(()=>activate(session.id));row.lastChild.onclick=()=>run(()=>close(session.id));}
      const name=(session.child?T.componentTabPrefix:'')+(session.data.name||T.untitled)+(session.meta.dirty?' •':'');
      row.firstChild.textContent=name;row.firstChild.setAttribute('aria-selected',String(active===session));row.classList.toggle('active',active===session);row.lastChild.title=T.workspaceCloseTab;row.lastChild.setAttribute('aria-label',T.workspaceCloseTab+' '+name);
    }
  }
  function renderChildBar(){
    document.querySelector('#component-child-bar')?.remove();document.querySelector('#component-preview-card')?.remove();
    if(!active?.child)return;
    const bar=document.createElement('div');bar.id='component-child-bar';bar.className='component-child-bar';
    bar.innerHTML='<span>'+T.componentDraftNote+'</span>'+btn('component-apply',T.componentApply,'primary')+btn('component-cancel',T.cancel);
    document.querySelector('.workarea').append(bar);
    bar.querySelector('#component-apply').onclick=()=>run(()=>finishChild(true));bar.querySelector('#component-cancel').onclick=()=>run(()=>finishChild(false));
    const preview=document.createElement('aside');preview.id='component-preview-card';preview.className='component-preview-card';preview.innerHTML='<strong>'+T.componentParentPreview+'</strong><img id="component-parent-preview" alt="'+T.componentParentPreview+'">';
    document.querySelector('.workarea').append(preview);
  }
  function enhanceProps(){
    const object=engine.selected;if(!object?.dakeComponent)return;
    const div=document.createElement('section');div.className='props-section component-settings';
    div.innerHTML='<div class="section-label">'+T.componentTitle+'</div><p>'+T[object.dakeComponent.mode==='linked'?'componentLinked':'componentEmbedded']+'</p>'+btn('component-edit',T.componentEdit,'widebtn')+(object.dakeComponent.mode==='linked'?btn('component-refresh',T.componentRefresh,'widebtn')+btn('component-embed',T.componentEmbed,'widebtn')+'<p class="small-note">'+esc(object.dakeComponent.link.path)+'</p>':'');
    document.querySelector('#props').prepend(div);div.querySelector('#component-edit').onclick=()=>run(()=>editComponent(object));
    div.querySelector('#component-refresh')?.addEventListener('click',()=>run(refreshComponent));div.querySelector('#component-embed')?.addEventListener('click',()=>run(embedComponent));
    const ungroup=document.querySelector('#ungroup');if(ungroup)ungroup.disabled=true;
  }
  function recovery(){
    capture();return {format:'dake-workspace',version:1,activeId:active?.id,documents:sessions.map(s=>({id:s.id,data:copy(s.data),meta:copy(s.meta),...(s.child?{child:copy(s.child)}:{})}))};
  }
  async function restoreRecovery(value){
    if(value?.format!=='dake-workspace'||value.version!==1||!Array.isArray(value.documents)||!value.documents.length||value.documents.length>16)throw new Error('INVALID_DOCUMENT');
    // Recovery can be chosen after starting new work. Never replace those sessions.
    if(sessions.length+value.documents.length>16)throw new Error('WORKSPACE_LIMIT');
    const remap=new Map();for(const item of value.documents){if(typeof item.id!=='string'||remap.has(item.id))throw new Error('INVALID_DOCUMENT');remap.set(item.id,uid());}
    const loaded=value.documents.map(item=>{
      if(item.child&&!remap.has(item.child.parentId))throw new Error('INVALID_DOCUMENT');
      return {...makeSession(item.data,{...item.meta,path:null,dirty:true}),id:remap.get(item.id),child:item.child?{...copy(item.child),parentId:remap.get(item.child.parentId)}:null};
    });
    await transition(async()=>{
      const first=loaded.find(s=>s.id===remap.get(value.activeId))||loaded[0];
      await production.loadDocument(first.data);sessions.push(...loaded);active=first;ctx.restoreMeta(first.meta);
    });renderChildBar();if(active.child)renderParentPreview();
  }
  async function prepareExit(){
    capture();
    for(const s of [...sessions].filter(s=>s.child)){await activate(s.id);if(!await finishChild(false))return false;}
    for(const s of sessions){if(s.meta.dirty){await activate(s.id);if(!await askClose(s))return false;}}
    return true;
  }
  function setup(){
    active=makeSession(engine.serialize(),ctx.captureMeta());active.history=history();sessions.push(active);
    const tabs=document.createElement('nav');tabs.id='document-tabs';tabs.className='document-tabs';tabs.setAttribute('role','tablist');tabs.setAttribute('aria-label',T.workspaceTabs);document.querySelector('.header').after(tabs);
    const importButton=document.createElement('button');importButton.id='place-component';importButton.className='compact';importButton.textContent=T.componentPlace;document.querySelector('#import').after(importButton);
    importButton.onclick=()=>run(async()=>{const result=await modal('<h1>'+T.componentPlace+'</h1><p>'+T.componentModeNote+'</p>',[{id:'cancel',text:T.cancel},{id:'linked',text:T.componentLinked},{id:'embedded',text:T.componentEmbedded,cls:'primary'}]);if(result.id!=='cancel')await place(result.id);});
    // Existing presets/drop/native-open paths call the same engine; capture them centrally.
    engine.newDocument=options=>switching?baseNew(options):newDocument(options);
    engine.loadDocument=async(data,options)=>{
      if(switching)return baseLoad(data,options);
      if(sessions.length>=16)throw new Error('WORKSPACE_LIMIT');
      let result;await transition(async()=>{result=await baseLoad(data,options);const s=makeSession(engine.serialize(),{dirty:false,documentEpoch:Date.now()});s.history=history();sessions.push(s);active=s;ctx.restoreMeta(s.meta);});
      renderChildBar();return result;
    };
    engine.canvas.on('mouse:dblclick',event=>{if(event.target?.dakeComponent)run(()=>editComponent(event.target));});
    renderTabs();
  }
  return {setup,changed,syncMeta,capture,recordSaved,findByPath,activate,newDocument,openDocument,close,place,editComponent,finishChild,resolveDrafts,refreshComponent,embedComponent,enhanceProps,recovery,restoreRecovery,prepareExit,get sessions(){capture();return sessions;},get active(){return active;},get switching(){return switching;},get anyDirty(){capture();return sessions.some(s=>s.meta.dirty);},get editingChild(){return !!active?.child;}};
}
