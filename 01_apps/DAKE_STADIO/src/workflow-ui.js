import {createColorPicker, rememberColor} from './color-picker.js';
import { Rect, Ellipse, Triangle, Line, Point } from 'fabric';

// UI gestures only change live geometry. One commit is published when the gesture ends.
export function createWorkflowUI(ctx) {
  const {engine,T,run,modal,btn,field,esc,renderProps,renderLayers,setTool,production,status,save}=ctx;
  const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
  let tab='edit',preview=null,gesture=null,shape=null,menu=null;
  const text=o=>['textbox','i-text','text'].includes(o?.type?.toLowerCase());
  const clamp=(n,a,b)=>Math.max(a,Math.min(b,Number(n)||a));
  function begin(target,keys){
    finishPreview(true);
    if(!target||target.locked||engine.busy)return false;
    preview={target,before:Object.fromEntries(keys.map(k=>[k,target[k]]))};return true;
  }
  function live(props){if(!preview)return;preview.target.set(props);if(text(preview.target))preview.target.initDimensions();preview.target.setCoords();preview.target.dirty=true;engine.canvas.requestRenderAll();}
  function finishPreview(apply=true){
    if(!preview)return;const p=preview;preview=null;
    if(!apply){p.target.set(p.before);if(text(p.target))p.target.initDimensions();p.target.setCoords();engine.canvas.requestRenderAll();}
    $('.color-popover')?.remove();if(apply){if(p.colorKey)rememberColor(p.target[p.colorKey]);engine.commit('preview');}
  }
  function selectTab(next){finishPreview(true);tab=next;refresh();}
  function refresh(){
    if(!$('#props'))return;
    for(const b of $$('[data-panel]')){b.classList.toggle('active',b.dataset.panel===tab);b.setAttribute('aria-selected',String(b.dataset.panel===tab));}
    for(const section of [...$('#props').children]){
      let category='edit';
      if(section.querySelector('#document-settings,#quick-snap,#guide-add,#align-right'))category='layout';
      if(section.dataset.panelCategory==='effects'||section.querySelector('#image-adjust-open,#convert-path,#mask-create,#group,#adjust-brightness,#outline-text'))category='effects';
      // Text editing remains together; outlining moves to the layer context menu as well.
      if(section.querySelector('#prop-text'))category='edit';
      section.hidden=category!==tab;
      if(section.tagName==='DETAILS'&&category===tab)section.open=true;
    }
    if(tab==='layout')drawGuidesPanel();
    addSliders();colorButtons();
    $('.font-applied-state')?.remove();const font=$('#prop-font');if(font){font.style.fontFamily=engine.selected?.fontFamily;const state=document.createElement('small');state.className='font-applied-state';state.textContent=T.fontState+': '+engine.selected.fontFamily;font.closest('label').append(state);}
  }
  function addSliders(){
    if(!engine.selected||engine.selected.locked)return;
    const s=engine.selected;
    for(const [id,key,min,max,factor] of [['prop-fontsize','fontSize',1,800,1],['prop-tracking','charSpacing',-.2,2,1000],['prop-leading','lineHeight',.5,4,1],['prop-opacity','opacity',0,100,.01]]){
      const input=$('#'+id);if(!input||input.parentElement.querySelector('input[type=range]'))continue;
      const range=document.createElement('input');range.type='range';range.id=id+'-slider';range.min=min;range.max=max;range.step=id==='prop-tracking'?.01:id==='prop-leading'?.05:1;range.value=input.value;range.setAttribute('aria-label',input.closest('label').querySelector('span').textContent);input.after(range);
      const start=()=>{if(!preview)begin(s,[key]);};
      const update=e=>{start();const value=clamp(e.target.value,min,max);input.value=value;range.value=value;live({[key]:value*factor});};
      range.oninput=update;range.onchange=e=>{update(e);finishPreview(true);};range.onpointercancel=()=>finishPreview(false);
      input.oninput=update;input.onchange=e=>{update(e);finishPreview(true);};input.onblur=()=>finishPreview(true);
    }
    for(const axis of ['width','height']){
      const input=$('#prop-'+axis);if(!input||input.parentElement.querySelector('input[type=range]'))continue;
      const key=axis==='width'?'scaleX':'scaleY',extent=axis==='width'?'getScaledWidth':'getScaledHeight';
      const range=document.createElement('input');range.type='range';range.id='prop-'+axis+'-slider';range.min=1;range.max=Math.max(2000,Number(input.value)*2);range.step=engine.meta.unit==='mm'?.1:1;range.value=input.value;range.setAttribute('aria-label',T[axis]);input.after(range);
      const update=e=>{if(!preview)begin(s,[key]);const val=clamp(e.target.value,1,32768);const px=production.toPx(val);live({[key]:clamp(s[key]*px/s[extent](),.001,1000)});input.value=val;range.value=val;};
      range.oninput=update;range.onchange=e=>{update(e);finishPreview(true);};range.onpointercancel=()=>finishPreview(false);input.oninput=update;input.onchange=e=>{update(e);finishPreview(true);};input.onblur=()=>finishPreview(true);
    }
  }
  function colorButtons(){
    for(const id of ['style-fill','style-stroke']){
      const input=$('#'+id);if(!input||input.dataset.polished)continue;input.dataset.polished='true';
      input.onclick=e=>{if(!engine.selected||engine.selected.locked)return;e.preventDefault();openColor(id==='style-fill'?'fill':'stroke',input.value);};
    }
  }
  function openColor(key,initial){
    if(!begin(engine.selected,[key]))return;preview.colorKey=key;
    createColorPicker({initial,onPreview:color=>live({[key]:color}),onFinish:apply=>{finishPreview(apply);renderProps();}});
    status('previewHint');
  }
  async function rename(layer){const r=await modal(`<h1>${T.renameLayer}</h1>${field('rename-value',T.layerName,layer.name,'text','maxlength="100"')}`,[{id:'cancel',text:T.cancel},{id:'apply',text:T.apply,cls:'primary'}]);if(r.id==='apply')engine.renameLayer(layer.id,r.values['rename-value'].trim()||layer.name);}
  function context(layer,x,y){
    closeMenu();engine.select(layer.id);menu=document.createElement('div');menu.className='context-menu';menu.setAttribute('role','menu');
    const entries=[[T.editObject,()=>{if(text(layer)){setTool('select');layer.enterEditing();}else if(layer.type==='path')setTool('node');else{$('#prop-width')?.focus();}},layer.locked],[T.renameLayer,()=>rename(layer),layer.locked],[T.rasterize,()=>engine.rasterizeSelected(),layer.locked||layer.type==='image'],[T.textOutline,()=>production.outlineText(),layer.locked||!text(layer)],[T.duplicate,()=>engine.duplicate(),layer.locked],[T.show,()=>engine.toggleVisible(layer.id),false],[T.lock,()=>engine.toggleLock(layer.id),false],[T.delete,()=>engine.deleteSelected(),layer.locked]];
    for(const [label,fn,disabled] of entries){const b=document.createElement('button');b.textContent=label;b.disabled=!!disabled;b.setAttribute('role','menuitem');b.onclick=()=>{closeMenu();run(fn);};menu.append(b);}
    document.body.append(menu);menu.style.left=Math.max(0,Math.min(x,innerWidth-menu.offsetWidth-8))+'px';menu.style.top=Math.max(0,Math.min(y,innerHeight-menu.offsetHeight-8))+'px';menu.querySelector('button:not(:disabled)')?.focus();
  }
  function closeMenu(){menu?.remove();menu=null;}
  function layers(){for(const row of $$('.layer-row')){const layer=engine.layers.find(o=>o.id===row.dataset.layer);row.oncontextmenu=e=>{e.preventDefault();context(layer,e.clientX,e.clientY);};if(!row.querySelector('.layer-more')){const b=document.createElement('button');b.className='layer-more';b.textContent='⋯';b.title=b.ariaLabel=T.layerMenu;b.onclick=e=>{e.stopPropagation();const r=b.getBoundingClientRect();context(layer,r.left,r.bottom);};row.append(b);}row.tabIndex=0;}}
  function drawGuidesPanel(){
    // Independent of selection: guide controls are always reachable in this tab.
    for(const section of [...$('#props').children])if(section.querySelector('#guide-add'))section.remove();
    $('#workflow-guides')?.remove();const panel=document.createElement('section');panel.id='workflow-guides';panel.className='props-section';
    panel.innerHTML=`<div class="section-label">${T.guides} · ${engine.meta.unit}</div><div class="checkgrid"><label class="check"><input id="guide-visible" type="checkbox" ${engine.view.guidesVisible?'checked':''}>${T.guideVisible}</label><label class="check"><input id="guide-lock" type="checkbox" ${engine.meta.guideLocked?'checked':''}>${T.guideLock}</label></div><p>${T.guideDragHint}</p><div class="fieldgrid"><label class="field"><span>${T.guideAdd}</span><select id="guide-axis"><option value="x">${T.guideVertical}</option><option value="y">${T.guideHorizontal}</option></select></label>${field('guide-pos',T.guidePosition,0,'number','min="0" step="0.1"')}</div>${btn('guide-add',T.guideAdd,'widebtn')}<div class="guide-list"></div>`;
    $('#props').prepend(panel);
    $('#guide-visible').onchange=e=>engine.setView({guidesVisible:e.target.checked});$('#guide-lock').onchange=e=>engine.updateDocumentSettings({guideLocked:e.target.checked});
    $('#guide-add').onclick=()=>run(()=>{engine.addGuide($('#guide-axis').value,production.toPx($('#guide-pos').value));renderProps();});
    for(const [axis,key]of [['x','vertical'],['y','horizontal']])engine.meta.guides[key].forEach((v,i)=>{
      const row=document.createElement('div');row.className='guide-row';row.innerHTML=`<span>${axis==='x'?'X':'Y'}</span><input type="number" aria-label="${T.guidePosition}" value="${Math.round(production.fromPx(v)*100)/100}" min="0" step="0.1"><span>${engine.meta.unit}</span><button aria-label="${T.guideRemove}" title="${T.guideRemove}">×</button>`;
      row.querySelector('input').onchange=e=>run(()=>{const next=[...engine.meta.guides[key]];next[i]=production.toPx(e.target.value);engine.updateDocumentSettings({guides:{[key]:next}});renderProps();});row.querySelector('button').onclick=()=>run(()=>{engine.removeGuide(axis,i);renderProps();});row.querySelectorAll('input,button').forEach(c=>c.disabled=!!engine.meta.guideLocked);panel.querySelector('.guide-list').append(row);
    });$('#guide-add').disabled=!!engine.meta.guideLocked;
  }
  function drawGuideHits(){
    const host=$('#guide-handles');if(!host||gesture?.kind==='guide')return;host.replaceChildren();if(!engine.view.guidesVisible||engine.meta.guideLocked)return;
    const v=engine.canvas.viewportTransform,z=engine.zoom;
    for(const [axis,key]of [['x','vertical'],['y','horizontal']])engine.meta.guides[key].forEach((pos,index)=>{
      const handle=document.createElement('div');handle.className='guide-handle '+axis;handle.title=`${axis.toUpperCase()}: ${Math.round(production.fromPx(pos)*100)/100} ${engine.meta.unit}`;
      handle.style[axis==='x'?'left':'top']=(v[axis==='x'?4:5]+pos*z)+'px';handle.onmousedown=e=>{if(engine.busy)return;e.preventDefault();e.stopPropagation();finishPreview(true);gesture={kind:'guide',axis,key,index,before:[...engine.meta.guides[key]],last:e};};host.append(handle);
    });
  }
  function point(e){return engine.canvas.getScenePoint(e);}
  function updateGesture(e){
    if(!gesture)return;gesture.last=e;const p=point(e),g=gesture;
    if(g.kind==='guide'){engine.meta.guides[g.key][g.index]=clamp(g.axis==='x'?p.x:p.y,0,g.axis==='x'?engine.width:engine.height);production.scheduleOverlay();status(`${g.axis.toUpperCase()}: ${production.fromPx(engine.meta.guides[g.key][g.index]).toFixed(2)} ${engine.meta.unit}`);return;}
    if(g.kind==='shape'){
      let dx=p.x-g.start.x,dy=p.y-g.start.y;if(e.ctrlKey){dx=Math.round(dx/10)*10;dy=Math.round(dy/10)*10;}
      let width=Math.max(1,Math.abs(dx)),height=Math.max(1,Math.abs(dy));
      if(e.shiftKey){width=Math.max(width,height);height=g.shape==='triangle'?width*Math.sqrt(3)/2:width;}
      const left=g.start.x-(e.altKey?width:dx<0?width:0),top=g.start.y-(e.altKey?height:dy<0?height:0);if(e.altKey){width*=2;height*=2;}
      if(g.shape==='ellipse')g.object.set({left,top,rx:width/2,ry:height/2});
      else if(g.shape==='line')g.object.set({x1:0,y1:0,x2:dx,y2:e.shiftKey?0:dy,left:Math.min(g.start.x,p.x),top:Math.min(g.start.y,p.y)});
      else g.object.set({left,top,width,height});g.object.setCoords();engine.canvas.requestRenderAll();return;
    }
    if(g.kind==='resize'){
      const dx=p.x-g.start.x,dy=p.y-g.start.y,b=g.before;
      let sx=clamp(b.scaleX*(1+dx/g.width),.01,100),sy=clamp(b.scaleY*(1+dy/g.height),.01,100);
      if(e.shiftKey){const f=Math.abs(dx)>Math.abs(dy)?sx/b.scaleX:sy/b.scaleY;sx=b.scaleX*f;sy=b.scaleY*f;}
      live({scaleX:sx,scaleY:sy,left:b.left,top:b.top});
      if(e.altKey)g.object.setPositionByOrigin(g.center,'center','center');g.object.setCoords();engine.canvas.requestRenderAll();
    }
  }
  function finishGesture(apply=true){
    if(!gesture)return;const g=gesture;gesture=null;
    if(g.kind==='guide'){if(!apply)engine.meta.guides[g.key]=g.before;else engine.updateDocumentSettings({guides:{[g.key]:engine.meta.guides[g.key]}});renderProps();production.scheduleOverlay();drawGuideHits();}
    else if(g.kind==='resize')finishPreview(apply);
    else{engine.canvas.remove(g.object);if(apply){g.object.excludeFromExport=false;g.object.selectable=true;g.object.evented=true;engine._add(g.object,g.shape);}shape=null;setTool('select');}
  }
  function chooseShape(kind){finishGesture(false);finishPreview(true);shape=kind;setTool('select');engine.canvas.selection=false;engine.canvas.skipTargetFind=true;engine.canvas.defaultCursor='crosshair';$$('[data-tool]').forEach(b=>b.classList.toggle('active',b.dataset.tool===kind));production.hideWelcome();status('shapeHelp');}
  function setup(){
    const head=$('.inspector-head');const tabs=document.createElement('div');tabs.className='panel-tabs';tabs.setAttribute('role','tablist');tabs.innerHTML=[['edit','propertiesTab'],['effects','effectsTab'],['layout','layoutTab']].map(([id,k])=>`<button role="tab" data-panel="${id}">${T[k]}</button>`).join('');head.after(tabs);for(const b of tabs.children)b.onclick=()=>selectTab(b.dataset.panel);
    $('#layout-settings').onclick=()=>selectTab('layout');
    const saveMore=document.createElement('button');saveMore.id='save-menu';saveMore.textContent='▾';saveMore.title=saveMore.ariaLabel=T.saveMenu;$('#save').after(saveMore);saveMore.onclick=()=>{closeMenu();menu=document.createElement('div');menu.className='context-menu';for(const [label,fn]of [[T.save,()=>save(false)],[T.saveAs,()=>save(true)],[T.saveVersion,()=>save(false,true)]]){const b=document.createElement('button');b.textContent=label;b.onclick=()=>{closeMenu();run(fn);};menu.append(b);}document.body.append(menu);const r=saveMore.getBoundingClientRect();menu.style.right=(innerWidth-r.right)+'px';menu.style.top=r.bottom+'px';};
    $('.canvas-host').insertAdjacentHTML('beforeend','<div id="guide-handles"></div>');engine.canvas.on('after:render',drawGuideHits);
    for(const b of $$('[data-tool]'))if(['rect','ellipse','triangle','line'].includes(b.dataset.tool))b.onclick=()=>chooseShape(b.dataset.tool);else b.addEventListener('click',()=>{shape=null;});
    $('.canvas-host').addEventListener('mousedown',e=>{
      if($('.color-popover')){e.preventDefault();e.stopImmediatePropagation();finishPreview(true);return;}
      if(engine.busy||e.button!==0||e.target.closest('.guide-handle'))return;
      if(shape){e.preventDefault();e.stopImmediatePropagation();const start=point(e),props={...engine.style,left:start.x,top:start.y,originX:'left',originY:'top',width:1,height:1,selectable:false,evented:false,excludeFromExport:true};const object=shape==='ellipse'?new Ellipse({...props,rx:1,ry:1}):shape==='triangle'?new Triangle(props):shape==='line'?new Line([0,0,1,1],{...props,stroke:engine.style.stroke||engine.style.fill,strokeWidth:engine.style.strokeWidth||3}):new Rect(props);engine.canvas.add(object);gesture={kind:'shape',shape,object,start,last:e};return;}
      const object=engine.selected;if(e.ctrlKey&&object&&!object.locked&&object.containsPoint(point(e))){e.preventDefault();e.stopImmediatePropagation();begin(object,['scaleX','scaleY','left','top']);gesture={kind:'resize',object,start:point(e),before:{...preview.before},center:object.getCenterPoint(),width:object.getScaledWidth(),height:object.getScaledHeight(),last:e};}
    },true);
    window.addEventListener('mousemove',e=>{if(gesture){e.preventDefault();e.stopImmediatePropagation();updateGesture(e);}},true);
    window.addEventListener('mouseup',e=>{if(gesture){e.preventDefault();e.stopImmediatePropagation();finishGesture(true);}},true);
    for(const event of ['keydown','keyup'])document.addEventListener(event,e=>{
      if(gesture&&['Shift','Alt','Control'].includes(e.key)){const last=gesture.last;updateGesture({clientX:last.clientX,clientY:last.clientY,shiftKey:e.shiftKey,altKey:e.altKey,ctrlKey:e.ctrlKey});e.preventDefault();e.stopImmediatePropagation();}
      if(event==='keydown'&&e.key==='Escape'&&(preview||gesture||shape||menu)){e.preventDefault();e.stopImmediatePropagation();finishGesture(false);finishPreview(false);shape=null;closeMenu();setTool('select');renderProps();}
      if(event==='keydown'&&e.key==='F2'&&engine.selected&&!engine.selected.locked&&!$('.dialog-backdrop')){e.preventDefault();run(()=>rename(engine.selected));}
    },true);
    window.addEventListener('blur',()=>{finishGesture(false);finishPreview(false);});
    document.addEventListener('mousedown',e=>{if($('.color-popover')&&e.target.closest('.canvas-host')){e.preventDefault();e.stopImmediatePropagation();finishPreview(true);return;}if(menu&&!menu.contains(e.target)&&e.target.id!=='save-menu')closeMenu();if(preview&&!gesture&&!e.target.closest('.color-popover,#props'))finishPreview(true);},true);
    // Keep the existing extensible renderer, but refresh only after its own DOM update.
    const oldProps=production.enhanceProps;production.enhanceProps=()=>{oldProps();refresh();};const oldLayers=production.enhanceLayers;production.enhanceLayers=()=>{oldLayers();layers();};
    engine.canvas.uniformScaling=false;engine.canvas.uniScaleKey='shiftKey';engine.canvas.centeredKey='altKey';
    refresh();layers();
  }
  return {setup,refresh,layers,chooseShape,selectTab,openColor,finishPreview,finishGesture,get previewing(){return !!preview||!!gesture;}};
}
