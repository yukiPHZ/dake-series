import {StudioEngine} from '../src/engine.js';
import {validateImageEditing,normalizeImageRecipe} from '../src/engine-imaging.js';
import {Point} from 'fabric';
const readPixels=async url=>{const image=new Image();image.src=url;await image.decode();const c=document.createElement('canvas');c.width=image.width;c.height=image.height;const ctx=c.getContext('2d');ctx.drawImage(image,0,0);const p=ctx.getImageData(0,0,c.width,c.height).data;return {width:c.width,height:c.height,at:(x,y)=>Array.from(p.slice((y*c.width+x)*4,(y*c.width+x)*4+4))};};
window.runImagingChecks=async()=>{
 const results=[],check=(name,passed,detail)=>{results.push({name,passed:!!passed,...(detail?{detail}:{})});if(!passed)throw new Error(name);};
 const e=new StudioEngine(document.querySelector('#editor'),{onChange(){},onSelection(){},onStatus(){}});window.imagingEngine=e;
 await e.newDocument({width:128,height:96,name:'imaging-regression',background:'transparent'});
 const c=document.createElement('canvas');c.width=128;c.height=96;const ctx=c.getContext('2d');
 for(let x=0;x<128;x++){ctx.fillStyle='rgb('+(40+x)+','+(60+x)+','+(20+x)+')';ctx.fillRect(x,0,1,96);}
 const source=c.toDataURL();let image=await e.importImage(source,'source-photo');e.updateSelected({left:0,top:0,scaleX:1,scaleY:1});
 const legacyId=image.id;await e.applyImageAdjustment({brightness:.12,contrast:.1,saturation:.2,grayscale:true});const legacyPNG=await e.exportRaster();e.beginImagePreview();await e.previewImageRecipe(e.getImageRecipe());check('legacy four filters migrate to recipe with identical pixels',await e.exportRaster()===legacyPNG);await e.finishImagePreview(false);check('cancel migration retains legacy filter definitions',e.selected.filters.length===4);await e.undo();e.select(legacyId);image=e.selected;
 const id=image.id,originalPNG=await e.exportRaster('png'),history=e._historyIndex;
 e.beginImagePreview();await e.previewImageRecipe(normalizeImageRecipe({exposure:1,vibrance:.5,curve:[[0,0],[64,40],[128,150],[192,220],[255,255]]}));
 check('live recipe changes rendered pixels before confirmation',await e.exportRaster('png')!==originalPNG);
 check('live preview keeps original source and makes no history',image.toObject().src===source&&e._historyIndex===history);
 await e.finishImagePreview(false);check('cancel restores pixels and recipe exactly',await e.exportRaster('png')===originalPNG&&!image.dakeImageEdit&&e._historyIndex===history);
 e.beginImagePreview();const a=e.previewImageRecipe(normalizeImageRecipe({exposure:-2})),b=e.previewImageRecipe(normalizeImageRecipe({exposure:1}));await Promise.all([a,b]);await e.finishImagePreview(true);
 check('latest recipe wins and commits one history entry',image.dakeImageEdit.exposure===1&&e._historyIndex===history+1);
 const adjustedPNG=await e.exportRaster('png'),adjusted=await readPixels(adjustedPNG),original=await readPixels(originalPNG);
 check('exposure acts on pixel luminance',adjusted.at(30,30)[0]>original.at(30,30)[0]);
 await e.undo();check('recipe undo returns original pixels',await e.exportRaster('png')===originalPNG);await e.redo();check('recipe redo returns adjusted pixels',await e.exportRaster('png')===adjustedPNG);e.select(id);image=e.selected;
 await e.createImageMask('layerMask');await e.configureImageMask('layerMask',{feather:15});
 check('white feathered mask does not fade image borders',await e.exportRaster('png')===adjustedPNG);
 await e.createImageMask('layerMask',{rect:{left:24,top:16,width:80,height:64}});
 let pixels=await readPixels(await e.exportRaster('png'));check('rect selection creates independent transparency mask',pixels.at(5,5)[3]===0&&pixels.at(64,48)[3]===255&&image.toObject().src===source);
 await e.configureImageMask('layerMask',{feather:5});pixels=await readPixels(await e.exportRaster('png'));check('feather produces intermediate alpha boundary',pixels.at(24,48)[3]>0&&pixels.at(24,48)[3]<255,pixels.at(24,48));
 await e.configureImageMask('layerMask',{enabled:false});check('mask disable restores corrected image',await e.exportRaster('png')===adjustedPNG);
 await e.configureImageMask('layerMask',{enabled:true,feather:0});
 await e.setImageEditingTarget('layerMask');e.setMaskBrush({mode:'hide',size:18,hardness:1});e.setTool('brush');
 let h=e._historyIndex;e._pointerDown({e:{button:0},scenePoint:new Point(64,48)});e._pointerMove({e:{},scenePoint:new Point(74,48)});await e._pointerUp();
 pixels=await readPixels(await e.exportRaster('png'));check('hide brush changes mask alpha only, one undo',pixels.at(64,48)[3]===0&&image.toObject().src===source&&e._historyIndex===h+1);
 e.setMaskBrush({mode:'restore',size:18,hardness:1});e._pointerDown({e:{button:0},scenePoint:new Point(64,48)});await e._pointerUp();
 pixels=await readPixels(await e.exportRaster('png'));check('white brush restores original image through mask',pixels.at(64,48)[3]===255&&pixels.at(64,48)[0]===adjusted.at(64,48)[0]);
 await e.removeImageMask('layerMask');
 await e.createImageMask('adjustmentMask',{rect:{left:64,top:0,width:64,height:96}});
 pixels=await readPixels(await e.exportRaster('png'));check('partial correction blends original outside and adjusted inside',JSON.stringify(pixels.at(20,40))===JSON.stringify(original.at(20,40))&&JSON.stringify(pixels.at(100,40))===JSON.stringify(adjusted.at(100,40)));
 const saved=e.serialize(),savedPNG=await e.exportRaster('png');await e.loadDocument(saved);check('recipe and correction mask save/load preserve exact pixels',await e.exportRaster('png')===savedPNG);
 e.select(id);image=e.selected;check('source remains editable original after reopen',image.toObject().src===source&&image.dakeImageEdit.exposure===1&&!!image.dakeAdjustmentMask);
 const bad=structuredClone(saved);bad.canvas.objects[0].dakeAdjustmentMask.src='data:image/png;base64,';let rejected=false;try{await e.loadDocument(bad);}catch{rejected=true;}
 check('invalid mask decode cannot replace current artwork',rejected&&await e.exportRaster('png')===savedPNG);
 for(const badRecipe of [{...normalizeImageRecipe(),exposure:6},{...normalizeImageRecipe(),curve:[[0,0],[0,255]]}]){
  let no=false;try{validateImageEditing({dakeImageEdit:badRecipe});}catch{no=true;}check('invalid recipe rejected '+results.length,no);
 }
 e.select(id);image=e.selected;await e.removeImageMask('adjustmentMask');await e.applyImageRecipe(normalizeImageRecipe({monochrome:true}));
 pixels=await readPixels(await e.exportRaster('png'));check('monochrome uses equal RGB channels',pixels.at(30,30)[0]===pixels.at(30,30)[1]&&pixels.at(30,30)[1]===pixels.at(30,30)[2]);
 await e.applyImageRecipe(normalizeImageRecipe());check('recipe reset reproduces unchanged source',await e.exportRaster('png')===originalPNG);
 await e.applyImageRecipe(normalizeImageRecipe({exposure:.5,saturation:.2}));await e.createImageMask('layerMask');
 await e.setImageEditingTarget('image');e.setTool('brush');e.setStyle({fill:'#1122cc',brushSize:14});e._pointerDown({e:{button:0},scenePoint:new Point(30,30)});await e._pointerUp();
 check('destructive image brush preserves independent recipe and mask',image.dakeImageEdit.exposure===.5&&image.dakeLayerMask&&image.toObject().src!==source);
 const brushPNG=await e.exportRaster('png'),brushSave=e.serialize();await e.loadDocument(brushSave);check('painted source with recipe and mask survives reopen',await e.exportRaster('png')===brushPNG);
 e.select(id);const originalElement=e.selected._element.toDataURL(),originalSrc=e.selected.toObject().src;const duplicates=await e.duplicate();check('duplicate restores derived recipe and mask appearance before use',duplicates[0]._element.toDataURL()===originalElement&&duplicates[0].toObject().src===originalSrc);await e.loadDocument(brushSave);
 e.select(id);await e.setImageEditingTarget('layerMask');e.setMaskBrush({mode:'hide',size:24,hardness:1});e.setTool('brush');e.canvas.setDimensions({width:500,height:320});e.canvas.setViewportTransform([2,0,0,2,30,30]);e.canvas.renderAll();
 window.nativeMaskBefore={history:e._historyIndex,src:e.selected.toObject().src};const r=e.canvas.upperCanvasEl.getBoundingClientRect();
 window.nativeMaskCoords={x:Math.round(r.left+30+60*2),y:Math.round(r.top+30+50*2)};
 return {passed:true,results,native:window.nativeMaskCoords,save:brushSave,png:brushPNG};
};
window.finishNativeMaskCheck=async()=>{
 const e=window.imagingEngine;await e._strokeTask;const pixel=(await readPixels(await e.exportRaster('png'))).at(60,50);
 return {passed:pixel[3]===0&&e._historyIndex===window.nativeMaskBefore.history+1&&e.selected.toObject().src===window.nativeMaskBefore.src,pixel};
};

