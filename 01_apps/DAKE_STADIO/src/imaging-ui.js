import { normalizeImageRecipe } from './engine-imaging.js';

export function createImagingUI(ctx) {
  const {engine,T,run,modal,btn,field,esc,renderProps,setTool}=ctx;
  const $=s=>document.querySelector(s);
  let open=false;
  function drawPreview(){const c=$('#im-adjust-preview');if(!c||!engine.selected)return;const source=engine.selected._element;if(!source)return;const width=source.naturalWidth||source.width,height=source.naturalHeight||source.height;if(!width||!height)return;const g=c.getContext('2d');g.clearRect(0,0,c.width,c.height);const scale=Math.min(c.width/width,c.height/height);g.drawImage(source,(c.width-width*scale)/2,(c.height-height*scale)/2,width*scale,height*scale);}
  const keyOf=kind=>kind==='adjustmentMask'?'dakeAdjustmentMask':'dakeLayerMask';
  const t=key=>T[key]||key;
  const isImage=o=>o?.type?.toLowerCase()==='image';
  function style(){
    if($('#imaging-styles'))return;
    const s=document.createElement('style');s.id='imaging-styles';s.textContent=
      '.imaging-actions{display:flex;gap:5px;flex-wrap:wrap}.imaging-actions button{flex:1;min-width:60px}.imaging-actions button.active{background:#b7e9d5;color:#112f27}.imaging-summary{font-size:11px;color:#9eaea8;margin:7px 0}.image-adjust-workbench{min-width:min(460px,70vw)}.image-adjust-tabs{display:flex;gap:6px;margin:12px 0}.image-adjust-tabs button{flex:1}.image-adjust-page{min-height:220px}.image-adjust-page label{margin-bottom:14px}.image-adjust-page canvas{width:100%;height:190px;touch-action:none;border:1px solid #465d53;border-radius:8px}.imaging-mask-thumb{height:28px;width:38px;object-fit:contain;background:#33423b;border:1px solid #73877d;vertical-align:middle}.image-adjust-preview{display:block;width:100%;height:145px;background:repeating-conic-gradient(#293c32 0% 25%,#203027 0% 50%) 0 / 16px 16px;border-radius:8px}.image-adjust-caption{font-size:12px;color:#aebdb7;line-height:1.55}.imaging-mask-settings .field{margin:8px 0}.imaging-mask-settings{margin-top:9px}.imaging-target-label{font-size:12px;margin:9px 0;color:#b7e9d5}';
    document.head.append(s);
  }
  async function openAdjustment() {
    if(!isImage(engine.selected))return;
    open=true;engine.beginImagePreview();let recipe=engine.getImageRecipe(),failed=null,curveDragging=null;
    try{
      const slider=(key,label,min,max,step=1,value=recipe[key])=>'<label class="field"><span>'+t(label)+' <output id="im-value-'+key+'">'+value+'</output></span><input id="im-adjust-'+key+'" type="range" min="'+min+'" max="'+max+'" step="'+step+'" value="'+value+'"></label>';
      const content='<div class="image-adjust-workbench"><h1>'+t('imageRecipeTitle')+'</h1><p class="image-adjust-caption">'+t('imageRecipeNote')+'</p>'+
        '<canvas id="im-adjust-preview" class="image-adjust-preview" width="600" height="170" aria-label="'+t('imagePreviewTarget')+'"></canvas><div class="image-adjust-tabs">'+[['light','imageLight'],['color','imageColor'],['curve','imageCurve']].map(([id,k])=>'<button type="button" data-im-page="'+id+'">'+t(k)+'</button>').join('')+'</div>'+
        '<div class="image-adjust-page" data-im-content="light">'+slider('brightness','brightness',-1,1,.01)+slider('exposure','imageExposure',-5,5,.05)+slider('contrast','contrast',-1,1,.01)+'</div>'+
        '<div class="image-adjust-page" data-im-content="color" hidden>'+slider('saturation','saturation',-1,1,.01)+slider('vibrance','imageVibrance',-1,1,.01)+'<label class="check"><input id="im-adjust-monochrome" type="checkbox" '+(recipe.monochrome?'checked':'')+'>'+t('grayscale')+'</label></div>'+
        '<div class="image-adjust-page" data-im-content="curve" hidden><canvas id="im-tone-curve" width="480" height="220" aria-label="'+t('imageCurveHint')+'"></canvas><p class="image-adjust-caption">'+t('imageCurveHint')+'</p></div>'+
        btn('im-adjust-reset',t('resetAdjust'))+'<p id="im-adjust-state" class="image-adjust-caption">'+t('imageLivePreview')+'</p></div>';
      const pending=modal(content,[{id:'cancel',text:t('cancel')},{id:'apply',text:t('apply'),cls:'primary'}],true);
      drawPreview();
      const preview=()=>{
        const local=structuredClone(recipe);$('#im-adjust-state').textContent=t('imageRendering');
        engine.previewImageRecipe(local).then(applied=>{if(applied&&$('#im-adjust-state')){$('#im-adjust-state').textContent=t('imageLivePreview');drawPreview();}})
          .catch(error=>{failed=error;if($('#im-adjust-state'))$('#im-adjust-state').textContent=t('rasterFailed');});
      };
      const selectPage=id=>{document.querySelectorAll('[data-im-content]').forEach(p=>p.hidden=p.dataset.imContent!==id);document.querySelectorAll('[data-im-page]').forEach(b=>b.classList.toggle('primary',b.dataset.imPage===id));};
      document.querySelectorAll('[data-im-page]').forEach(b=>b.onclick=()=>selectPage(b.dataset.imPage));selectPage('light');
      for(const key of ['brightness','contrast','exposure','saturation','vibrance'])$('#im-adjust-'+key).oninput=e=>{recipe[key]=Number(e.target.value);$('#im-value-'+key).textContent=e.target.value;preview();};
      $('#im-adjust-monochrome').onchange=e=>{recipe.monochrome=e.target.checked;recipe.monochromeMode='luminosity';preview();};
      const graph=$('#im-tone-curve'),g=graph.getContext('2d'),pad=15,w=graph.width-pad*2,h=graph.height-pad*2;
      const draw=()=>{
        g.fillStyle='#17241f';g.fillRect(0,0,graph.width,graph.height);g.strokeStyle='#41574c';g.lineWidth=1;
        for(let i=0;i<=4;i++){g.beginPath();g.moveTo(pad+i*w/4,pad);g.lineTo(pad+i*w/4,pad+h);g.moveTo(pad,pad+i*h/4);g.lineTo(pad+w,pad+i*h/4);g.stroke();}
        g.strokeStyle='#b7edb8';g.lineWidth=2;g.beginPath();recipe.curve.forEach(([x,y],i)=>i?g.lineTo(pad+x/255*w,pad+(255-y)/255*h):g.moveTo(pad+x/255*w,pad+(255-y)/255*h));g.stroke();
        for(const [x,y] of recipe.curve){g.beginPath();g.arc(pad+x/255*w,pad+(255-y)/255*h,5,0,Math.PI*2);g.fillStyle='#e6ffea';g.fill();}
      };
      const graphPoint=e=>{const r=graph.getBoundingClientRect();return {x:(e.clientX-r.left)*graph.width/r.width,y:(e.clientY-r.top)*graph.height/r.height};};
      const updateCurve=e=>{if(curveDragging===null)return;const p=graphPoint(e);recipe.curve[curveDragging][1]=Math.max(0,Math.min(255,Math.round(255-(p.y-pad)*255/h)));draw();preview();};
      graph.onpointerdown=e=>{const p=graphPoint(e);curveDragging=recipe.curve.reduce((best,v,i)=>Math.abs(pad+v[0]/255*w-p.x)<Math.abs(pad+recipe.curve[best][0]/255*w-p.x)?i:best,0);graph.setPointerCapture(e.pointerId);updateCurve(e);};
      graph.onpointermove=updateCurve;graph.onpointerup=()=>{curveDragging=null;};graph.onpointercancel=()=>{curveDragging=null;};draw();
      $('#im-adjust-reset').onclick=()=>{recipe=normalizeImageRecipe();for(const key of ['brightness','contrast','exposure','saturation','vibrance']){$('#im-adjust-'+key).value=recipe[key];$('#im-value-'+key).textContent=recipe[key];}$('#im-adjust-monochrome').checked=false;draw();preview();};
      const result=await pending;
      if(result.id==='apply'&&failed)throw failed;
      await engine.finishImagePreview(result.id==='apply');
    }catch(error){await engine.finishImagePreview(false);throw error;}
    finally{open=false;renderProps();}
  }
  async function chooseTarget(kind) {
    if(kind!=='image'&&!engine.selected[keyOf(kind)])await engine.createImageMask(kind);
    await engine.setImageEditingTarget(kind);
    if(kind!=='image')setTool('brush');
    else setTool('select');
    renderProps();
  }
  function enhanceProps() {
    style();const image=engine.selected;if(!isImage(image))return;
    $('#adjust-brightness')?.closest('.props-section')?.remove();$('#imaging-controls')?.remove();
    const panel=document.createElement('section');panel.className='props-section';panel.id='imaging-controls';panel.dataset.panelCategory='effects';
    let target=engine._imageEditingTarget||'image';if(target!=='image'&&!image[keyOf(target)])target=engine._imageEditingTarget='image';
    if(target!=='image'&&$('#hint'))$('#hint').textContent=t(engine._maskSelectionMode?'imageMaskSelectionHint':'imageMaskGestureHint');
    const kind=target==='adjustmentMask'?'adjustmentMask':'layerMask',mask=image[keyOf(kind)],brush=engine._maskBrush||{mode:'hide',size:60,hardness:.5};
    panel.innerHTML='<div class="section-label">'+t('imageNonDestructive')+'</div>'+btn('image-adjust-open',t('imageRecipeTitle'),'widebtn')+
      '<div class="imaging-summary">'+t(image.dakeImageEdit?'imageRecipeSaved':'imageSourceRetained')+'</div>'+
      '<div class="section-label">'+t('imageEditTarget')+'</div><div class="imaging-actions">'+[['image','imagePixels'],['layerMask','imageLayerMask'],['adjustmentMask','imageAdjustmentMask']].map(([id,k])=>'<button data-im-target="'+id+'" class="'+(target===id?'active':'')+'">'+t(k)+'</button>').join('')+'</div>'+
      '<p class="imaging-target-label">'+t(target==='image'?'imagePixelsNote':target==='layerMask'?'imageLayerMaskNote':'imageAdjustmentMaskNote')+'</p>'+
      '<div class="imaging-mask-settings" '+(target==='image'?'hidden':'')+'><div class="imaging-actions">'+btn('im-mask-hide',t('imageMaskHide'),brush.mode==='hide'?'active':'')+btn('im-mask-restore',t('imageMaskRestore'),brush.mode==='restore'?'active':'')+'</div>'+
      field('im-mask-size',t('brushSize'),brush.size,'range','min="1" max="500" step="1"')+
      field('im-mask-hardness',t('imageMaskHardness'),Math.round(brush.hardness*100),'range','min="0" max="100"')+
      field('im-mask-feather',t('imageMaskFeather'),mask?.feather||0,'range','min="0" max="100" step="1"')+
      '<div class="imaging-actions">'+btn('im-mask-select',t('imageMaskSelection'))+btn('im-mask-toggle',t(mask?.enabled?'imageMaskDisable':'imageMaskEnable'))+btn('im-mask-remove',t('maskRelease'))+'</div>'+
      (mask?'<p class="imaging-summary"><img class="imaging-mask-thumb" src="'+esc(mask.src)+'" alt=""> '+t('imageMaskStored')+'</p>':'')+'</div>';
    $('#props').prepend(panel);$('#image-adjust-open').onclick=()=>run(openAdjustment);
    panel.querySelectorAll('[data-im-target]').forEach(b=>b.onclick=()=>run(()=>chooseTarget(b.dataset.imTarget)));
    $('#im-mask-hide').onclick=()=>{engine.setMaskBrush({mode:'hide'});setTool('brush');renderProps();};
    $('#im-mask-restore').onclick=()=>{engine.setMaskBrush({mode:'restore'});setTool('brush');renderProps();};
    $('#im-mask-size').oninput=e=>engine.setMaskBrush({size:Number(e.target.value)});
    $('#im-mask-hardness').oninput=e=>engine.setMaskBrush({hardness:Number(e.target.value)/100});
    const feather=$('#im-mask-feather');
    feather.oninput=e=>{
      if(!engine._imagingPreview)engine.beginImagePreview();
      image[keyOf(kind)]={...image[keyOf(kind)],feather:Number(e.target.value)};
      engine._renderImageEditing(image).catch(()=>{engine.finishImagePreview(false);});
    };
    feather.onchange=()=>run(async()=>{await engine.finishImagePreview(true);renderProps();});
    feather.onpointercancel=()=>engine.finishImagePreview(false).then(renderProps);
    $('#im-mask-select').onclick=()=>run(()=>{engine.setMaskSelectionMode(kind);setTool('brush');});
    $('#im-mask-toggle').onclick=()=>run(async()=>{await engine.configureImageMask(kind,{enabled:!mask.enabled});renderProps();});
    $('#im-mask-remove').onclick=()=>run(async()=>{await engine.removeImageMask(kind);setTool('select');renderProps();});
  }
  function setup(){
    style();
    document.addEventListener('keydown',e=>{
      if(e.key!=='Escape'||open)return;
      if(engine._maskSelection){
        e.preventDefault();e.stopImmediatePropagation();engine.canvas.remove(engine._maskSelection.rect);engine._maskSelection=null;engine._maskSelectionMode=null;setTool('select');return;
      }
      if(engine._maskStroke){e.preventDefault();e.stopImmediatePropagation();engine._finishMaskStroke(false).then(()=>{setTool('select');renderProps();});}
      else if(engine._imagingPreview){e.preventDefault();e.stopImmediatePropagation();engine.finishImagePreview(false).then(renderProps);}
    },true);
  }
  return {setup,enhanceProps,openAdjustment,chooseTarget,finishPreview:apply=>open?Promise.resolve():engine.finishImagePreview(apply),get previewing(){return engine.imagingPreviewing;}};
}

