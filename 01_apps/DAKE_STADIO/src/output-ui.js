import UI_TEXT from './ui-text.json';
export function createOutputUI(ctx){
 const {engine,bridge,modal,field,btn,esc,run,fit,status,toast,renderProps,production}=ctx,T=UI_TEXT.output,$=s=>document.querySelector(s);
 const bytes=n=>n<1024?n+' B':n<1048576?(n/1024).toFixed(1)+' KiB':(n/1048576).toFixed(2)+' MiB';
 async function exportDocument(){
  engine.finishPath(false);const name=engine.name,sourceWidth=engine.width,sourceHeight=engine.height,sourceDpi=engine.meta?.dpi||96;
  const pending=modal('<h1>'+T.title+'</h1><div class="output-grid"><div class="output-image"><img id="output-image" alt="'+T.preview+'"><span id="output-wait"></span></div><div><label class="field"><span>'+T.format+'</span><select id="export-format"><option value="png">PNG</option><option value="jpeg">JPEG</option><option value="webp">WebP</option><option value="svg">SVG</option><option value="pdf">PDF</option></select></label><div class="fieldgrid">'+field('output-width',T.width,sourceWidth,'number','min="1" max="16384"')+field('output-height',T.height,sourceHeight,'number','min="1" max="16384"')+'</div><small>'+T.aspect+'</small>'+field('export-quality',T.quality,95,'range','min="50" max="100"')+'<label class="check"><input id="output-transparent" type="checkbox" '+(engine.background==='transparent'?'checked':'')+'>'+T.transparent+'</label><p id="output-size"></p><p id="output-dimensions"></p><p id="output-format-note" class="dialog-note"></p></div></div>',[{id:'cancel',text:UI_TEXT.cancel},{id:'export',text:UI_TEXT.export,cls:'primary'}],true);
  let generation=0,timer,working=false,queued=null,prepared=null,closed=false,previewUrl=null;
  const saveButton=$('#dialog-export');saveButton.disabled=true;
  function release(){if(prepared?.token)bridge?.releasePreparedExport(prepared.token);prepared=null;}
  function options(){const format=$('#export-format').value,width=Number($('#output-width').value);if(!Number.isInteger(width)||width<1||width>16384)throw new Error('EXPORT_TOO_LARGE');const height=Math.max(1,Math.round(width*sourceHeight/sourceWidth));if(height>16384||width*height>64000000)throw new Error('EXPORT_TOO_LARGE');return {format,width:format==='svg'?sourceWidth:width,quality:Number($('#export-quality').value)/100,transparent:['png','webp'].includes(format)&&$('#output-transparent').checked};}
  async function pump(){if(working||!queued||closed)return;working=true;const job=queued;queued=null;
   try{const o=job.options,format=o.format;let data,width=sourceWidth,height=sourceHeight;
    if(format==='svg')data=engine.exportSVG();else{data=await engine.exportRaster(format==='pdf'?'png':format,o.width/sourceWidth,o.quality,{transparent:o.transparent});const image=new Image();image.src=data;await image.decode();width=image.naturalWidth;height=image.naturalHeight;}
    if(closed||job.id!==generation)return;
    const result=bridge?await bridge.prepareExport({format,data,name,width,height,dpi:sourceDpi*width/sourceWidth}):{bytes:format==='svg'?new TextEncoder().encode(data).length:Math.floor((data.length-data.indexOf(',')-1)*3/4),format};
    if(closed||job.id!==generation){if(result.token)bridge?.releasePreparedExport(result.token);return;}
    prepared={...result,data};if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=format==='svg'?URL.createObjectURL(new Blob([data],{type:'image/svg+xml'})):null;
    $('#output-image').src=previewUrl||data;$('#output-wait').textContent='';$('#output-size').textContent=T.measured+': '+bytes(result.bytes);$('#output-dimensions').textContent=width+' × '+height+' px'+(format==='pdf'?' · '+(sourceWidth/sourceDpi*25.4).toFixed(2)+' × '+(sourceHeight/sourceDpi*25.4).toFixed(2)+' mm':'');saveButton.disabled=false;
   }catch(error){if(!closed&&job.id===generation){$('#output-wait').textContent=ctx.errorText(error);$('#output-size').textContent=T.failed;saveButton.disabled=true;}}
   finally{working=false;if(queued&&!closed)pump();}
  }
  function schedule(){generation++;saveButton.disabled=true;release();clearTimeout(timer);const format=$('#export-format').value;
   $('#export-quality').disabled=!['jpeg','webp'].includes(format);$('#output-transparent').disabled=!['png','webp'].includes(format);$('#output-width').disabled=$('#output-height').disabled=format==='svg';
   $('#output-format-note').textContent=format==='pdf'?T.pdf:format==='svg'?T.svg:T.raster;
   $('#output-size').textContent=T.pending;$('#output-wait').textContent=T.encoding;
   try{queued={id:generation,options:options()};timer=setTimeout(pump,120);}catch(error){queued=null;$('#output-wait').textContent=ctx.errorText(error);}
  }
  $('#output-width').oninput=()=>{$('#output-height').value=Math.max(1,Math.round(Number($('#output-width').value)*sourceHeight/sourceWidth));schedule();};
  $('#output-height').oninput=()=>{$('#output-width').value=Math.max(1,Math.round(Number($('#output-height').value)*sourceWidth/sourceHeight));schedule();};
  for(const id of ['export-format','export-quality','output-transparent'])$('#'+id).oninput=schedule;schedule();
  const result=await pending;closed=true;clearTimeout(timer);if(previewUrl)URL.revokeObjectURL(previewUrl);
  try{if(result.id==='export'&&prepared){const saved=bridge?await bridge.savePreparedExport(prepared.token):null;if(saved){status('exportDone');toast(T.saved+' '+bytes(saved.bytes));}}}finally{release();}
 }
 async function resize(){
  const pending=modal('<h1>'+T.resizeTitle+'</h1><p>'+T.resizeHint+'</p><label class="field"><span>'+T.mode+'</span><select id="resize-mode"><option value="extent">'+T.extent+'</option><option value="artwork">'+T.artwork+'</option></select></label><div class="fieldgrid">'+field('resize-width',T.width,engine.width,'number','min="1" max="8192"')+field('resize-height',T.height,engine.height,'number','min="1" max="8192"')+'</div><label class="field"><span>'+T.anchor+'</span><select id="resize-anchor"><option value="center">'+T.center+'</option><option value="corner">'+T.corner+'</option></select></label>',[{id:'cancel',text:UI_TEXT.cancel},{id:'apply',text:UI_TEXT.apply,cls:'primary'}]);
  const ratio=engine.height/engine.width;$('#resize-width').oninput=()=>{if($('#resize-mode').value==='artwork')$('#resize-height').value=Math.round(Number($('#resize-width').value)*ratio);};$('#resize-height').oninput=()=>{if($('#resize-mode').value==='artwork')$('#resize-width').value=Math.round(Number($('#resize-height').value)/ratio);};$('#resize-mode').onchange=()=>{$('#resize-anchor').disabled=$('#resize-mode').value==='artwork';if($('#resize-mode').value==='artwork')$('#resize-height').value=Math.round(Number($('#resize-width').value)*ratio);};
  const r=await pending;if(r.id==='apply'){engine.resizeDocument(+r.values['resize-width'],+r.values['resize-height'],{scaleArtwork:r.values['resize-mode']==='artwork',anchor:r.values['resize-anchor']});fit();renderProps();}
 }
 async function cropImage(){const image=engine.selected;if(image?.type?.toLowerCase()!=='image'||image.locked)return;const el=image._originalElement,w=el.naturalWidth||el.width,h=el.naturalHeight||el.height;
  const r=await modal('<h1>'+T.cropImage+'</h1><p>'+T.cropHint+'</p><div class="fieldgrid">'+field('image-crop-x','X',image.cropX||0,'number','min="0"')+field('image-crop-y','Y',image.cropY||0,'number','min="0"')+field('image-crop-w',T.width,image.width,'number','min="1"')+field('image-crop-h',T.height,image.height,'number','min="1"')+'</div>',[{id:'cancel',text:UI_TEXT.cancel},{id:'reset',text:T.resetCrop},{id:'apply',text:UI_TEXT.apply,cls:'primary'}]);
  if(r.id==='reset')engine.updateSelected({cropX:0,cropY:0,width:w,height:h});
  if(r.id==='apply'){const x=+r.values['image-crop-x'],y=+r.values['image-crop-y'],width=+r.values['image-crop-w'],height=+r.values['image-crop-h'];if(![x,y,width,height].every(Number.isInteger)||x<0||y<0||width<1||height<1||x+width>w||y+height>h)throw new Error('INVALID_IMAGE');engine.updateSelected({cropX:x,cropY:y,width,height});}renderProps();
 }
 function enhance(){const settings=$('#document-settings');if(settings&&!$('#resize-document')){const b=document.createElement('button');b.id='resize-document';b.className='widebtn';b.textContent=T.resizeTitle;b.onclick=()=>run(resize);settings.after(b);}if(engine.selected?.type?.toLowerCase()==='image'&&!$('#crop-image')){const b=document.createElement('button');b.id='crop-image';b.className='widebtn';b.textContent=T.cropImage;b.disabled=!!engine.selected.locked;b.onclick=()=>run(cropImage);$('#props').querySelector('.props-section')?.append(b);}}
 function setup(){const old=production.enhanceProps;production.enhanceProps=()=>{old();enhance();};enhance();}
 return {setup,enhance,exportDocument,resize,cropImage};
}
