import { FabricImage, Point, Rect, util } from 'fabric';

export const IMAGE_EDIT_PROPERTIES = ['dakeImageEdit','dakeLayerMask','dakeAdjustmentMask'];
const clone = value => value == null ? value : structuredClone(value);
const clamp = (n,a,b) => Math.max(a,Math.min(b,Number(n)||0));
const maskKey = kind => kind === 'adjustmentMask' ? 'dakeAdjustmentMask' : 'dakeLayerMask';
const isImage = image => image instanceof FabricImage;
const canvas = (w,h) => { const c=document.createElement('canvas');c.width=w;c.height=h;return c; };
const sizeOf = image => { const el=image._originalElement;return {width:el.naturalWidth||el.width,height:el.naturalHeight||el.height}; };
const imagePoint = (image,point) => {
  const p=new Point(point.x,point.y).transform(util.invertTransform(image.calcTransformMatrix()));
  return {x:p.x+image.width/2+(image.cropX||0),y:p.y+image.height/2+(image.cropY||0)};
};
export function normalizeImageRecipe(value={}) {
  return {version:1,brightness:clamp(value.brightness,-1,1),contrast:clamp(value.contrast,-1,1),
    exposure:clamp(value.exposure,-5,5),saturation:clamp(value.saturation,-1,1),vibrance:clamp(value.vibrance,-1,1),
    monochrome:!!(value.monochrome??value.grayscale),monochromeMode:value.monochromeMode==='average'?'average':'luminosity',
    curve:(Array.isArray(value.curve)?value.curve:[[0,0],[64,64],[128,128],[192,192],[255,255]]).map(p=>[Number(p[0]),Number(p[1])])};
}
export function validateImageEditing(raw) {
  if(raw.dakeImageEdit!=null){
    const p=raw.dakeImageEdit;
    if(p.version!==1||typeof p.monochrome!=='boolean'||!['average','luminosity'].includes(p.monochromeMode))throw new Error('INVALID_DOCUMENT');
    for(const [key,limit] of [['brightness',1],['contrast',1],['exposure',5],['saturation',1],['vibrance',1]])if(!Number.isFinite(p[key])||Math.abs(p[key])>limit)throw new Error('INVALID_DOCUMENT');
    if(!Array.isArray(p.curve)||p.curve.length<2||p.curve.length>16||p.curve[0]?.[0]!==0||p.curve.at(-1)?.[0]!==255)throw new Error('INVALID_DOCUMENT');
    let last=-1;for(const point of p.curve){if(!Array.isArray(point)||point.length!==2||point.some(n=>!Number.isFinite(n)||n<0||n>255)||point[0]<=last)throw new Error('INVALID_DOCUMENT');last=point[0];}
  }
  for(const key of ['dakeLayerMask','dakeAdjustmentMask'])if(raw[key]!=null){
    const p=raw[key];
    if(p.version!==1||typeof p.enabled!=='boolean'||typeof p.src!=='string'||!/^data:image\/png;base64,[A-Za-z0-9+/]*={0,2}$/.test(p.src)||p.src.length>96*1024*1024||!Number.isFinite(p.feather)||p.feather<0||p.feather>100)throw new Error('INVALID_DOCUMENT');
  }
  return raw;
}
export function installImaging(Engine) {
  if(Engine.prototype._imagingInstalled)return;
  const P=Engine.prototype;
  const original={newDocument:P.newDocument,applyImageAdjustment:P.applyImageAdjustment,pointerDown:P._pointerDown,pointerMove:P._pointerMove,pointerUp:P._pointerUp,startStroke:P._startStroke};
  Object.assign(P,{
    _imagingInstalled:true,
    newDocument(options){this._imageEditingTarget='image';this._maskSelectionMode=null;return original.newDocument.call(this,options);},
    get imagingPreviewing(){return !!this._imagingPreview||!!this._maskStroke||!!this._maskSelection;},
    getImageRecipe(image=this.selected){
      if(!isImage(image))throw new Error('NO_IMAGE_SELECTED');
      return normalizeImageRecipe(image.dakeImageEdit||{...image.adjustments,monochromeMode:'average'});
    },
    async _renderImageEditing(image) {
      validateImageEditing(image);
      const version=(image._dakeRenderVersion||0)+1;image._dakeRenderVersion=version;
      const request={recipe:clone(image.dakeImageEdit),layerMask:clone(image.dakeLayerMask),adjustmentMask:clone(image.dakeAdjustmentMask),
        filters:(image.filters||[]).map(f=>f.toObject())};
      const task=(this._imagingQueue||Promise.resolve()).catch(()=>{}).then(async()=>{
        if(image._dakeRenderVersion!==version)return false;
        const {width,height}=sizeOf(image);
        const bitmap=await createImageBitmap(image._originalElement);
        const specs=[...request.filters,{type:'DakeImageEditing',recipe:request.recipe,layerMask:request.layerMask,adjustmentMask:request.adjustmentMask}];
        const result=await this._filterInWorker(bitmap,specs);
        try{
          if(image._dakeRenderVersion!==version)return false;
          const rendered=canvas(width,height);rendered.getContext('2d').drawImage(result,0,0);
          image._filteredEl=rendered;image._element=rendered;
          image._filterScalingX=image._filterScalingY=image._lastScaleX=image._lastScaleY=1;
          image.set('dirty',true);this.canvas.requestRenderAll();return true;
        }finally{result.close();}
      });
      this._imagingQueue=task;return task;
    },
    async restoreImageEditing(image,raw) {
      if(!isImage(image)||!IMAGE_EDIT_PROPERTIES.some(key=>raw[key]))return;
      validateImageEditing(raw);
      for(const key of IMAGE_EDIT_PROPERTIES)image[key]=clone(raw[key]);
      for(const key of ['dakeLayerMask','dakeAdjustmentMask'])if(image[key]){
        const decoded=await this._imageMaskCanvas(image,key);
        const size=sizeOf(image);if(decoded.width!==size.width||decoded.height!==size.height)throw new Error('INVALID_DOCUMENT');
      }
      await this._renderImageEditing(image);
    },
    beginImagePreview() {
      this._guard();const image=this.selected;
      if(!isImage(image)||image.locked)throw new Error('NO_IMAGE_SELECTED');
      if(this._imagingPreview){if(this._imagingPreview.image===image)return;throw new Error('BUSY');}
      this._imagingPreview={image,before:Object.fromEntries(IMAGE_EDIT_PROPERTIES.map(k=>[k,clone(image[k])])),filters:image.filters,
        adjustments:clone(image.adjustments),rendered:image._element,filtered:image._filteredEl};
    },
    async previewImageRecipe(recipe) {
      if(!this._imagingPreview)this.beginImagePreview();
      const image=this._imagingPreview.image;const normalized=normalizeImageRecipe(recipe);validateImageEditing({dakeImageEdit:normalized});
      image.dakeImageEdit=normalized;image.filters=[];image.adjustments={brightness:normalized.brightness,contrast:normalized.contrast,saturation:normalized.saturation,grayscale:normalized.monochrome};
      return this._renderImageEditing(image);
    },
    async finishImagePreview(apply=true) {
      const preview=this._imagingPreview;if(!preview)return;
      if(!apply){
        const image=preview.image;++image._dakeRenderVersion;
        for(const key of IMAGE_EDIT_PROPERTIES)image[key]=preview.before[key];
        image.filters=preview.filters;image.adjustments=preview.adjustments;
        image._element=preview.rendered;image._filteredEl=preview.filtered;image.set('dirty',true);
        this._imagingPreview=null;await this._imagingQueue?.catch(()=>{});this.canvas.requestRenderAll();return;
      }
      try{
        await this._imagingQueue;
        this._imagingPreview=null;this.commit('image-adjustment');
      }catch(error){await this.finishImagePreview(false);throw error;}
    },
    async applyImageRecipe(recipe) {
      this.beginImagePreview();
      try{await this.previewImageRecipe(recipe);await this.finishImagePreview(true);}
      catch(error){await this.finishImagePreview(false);throw error;}
    },
    async applyImageAdjustment(options={}) {
      const image=this.selected;
      if(!image?.dakeImageEdit&&!image?.dakeLayerMask&&!image?.dakeAdjustmentMask)return original.applyImageAdjustment.call(this,options);
      return this.applyImageRecipe({...this.getImageRecipe(image),...options,monochrome:!!options.grayscale});
    },
    async _imageMaskCanvas(image,key) {
      const value=image[key];if(!value)throw new Error('INVALID_MASK');
      image._dakeMaskCanvases??={};const cached=image._dakeMaskCanvases[key];if(cached?.src===value.src)return cached.canvas;
      const element=new Image();element.src=value.src;await element.decode();
      const {width,height}=sizeOf(image);if(element.naturalWidth!==width||element.naturalHeight!==height)throw new Error('INVALID_DOCUMENT');
      const buffer=canvas(width,height);buffer.getContext('2d').drawImage(element,0,0);
      image._dakeMaskCanvases[key]={src:value.src,canvas:buffer};return buffer;
    },
    async createImageMask(kind='layerMask',{rect=null,sceneRect=null,invert=false}={}) {
      this._guard();const image=this.selected;if(!isImage(image)||image.locked)throw new Error('NO_IMAGE_SELECTED');
      const key=maskKey(kind),size=sizeOf(image),buffer=canvas(size.width,size.height),ctx=buffer.getContext('2d');
      const points=sceneRect?[
        {x:sceneRect.left,y:sceneRect.top},{x:sceneRect.left+sceneRect.width,y:sceneRect.top},
        {x:sceneRect.left+sceneRect.width,y:sceneRect.top+sceneRect.height},{x:sceneRect.left,y:sceneRect.top+sceneRect.height}
      ].map(p=>imagePoint(image,p)):null;
      ctx.fillStyle=rect||points?(invert?'#fff':'#000'):'#fff';ctx.fillRect(0,0,size.width,size.height);
      if(rect){ctx.fillStyle=invert?'#000':'#fff';ctx.fillRect(rect.left,rect.top,rect.width,rect.height);}
      if(points){ctx.fillStyle=invert?'#000':'#fff';ctx.beginPath();points.forEach((p,i)=>i?ctx.lineTo(p.x,p.y):ctx.moveTo(p.x,p.y));ctx.closePath();ctx.fill();}
      const before=clone(image[key]);image[key]={version:1,enabled:true,src:buffer.toDataURL('image/png'),feather:0};
      image._dakeMaskCanvases??={};image._dakeMaskCanvases[key]={src:image[key].src,canvas:buffer};
      return this._withBusy(async()=>{try{await this._renderImageEditing(image);this.commit('image-mask');return image[key];}catch(error){image[key]=before;await this._renderImageEditing(image);throw error;}});
    },
    async configureImageMask(kind,properties) {
      this._guard();const image=this.selected,key=maskKey(kind);if(!isImage(image)||!image[key]||image.locked)throw new Error('INVALID_MASK');
      const before=clone(image[key]),candidate={...image[key],...properties};validateImageEditing({[key]:candidate});image[key]=candidate;
      return this._withBusy(async()=>{try{await this._renderImageEditing(image);this.commit('image-mask');}catch(error){image[key]=before;await this._renderImageEditing(image);throw error;}});
    },
    async removeImageMask(kind) {
      this._guard();const image=this.selected,key=maskKey(kind);if(!isImage(image)||!image[key]||image.locked)throw new Error('INVALID_MASK');
      const before=image[key];delete image[key];
      return this._withBusy(async()=>{try{await this._renderImageEditing(image);this._imageEditingTarget='image';this.commit('image-mask');}catch(error){image[key]=before;await this._renderImageEditing(image);throw error;}});
    },
    async setImageEditingTarget(kind='image') {
      if(!['image','layerMask','adjustmentMask'].includes(kind))throw new Error('INVALID_MASK');
      const image=this.selected;if(!isImage(image)||image.locked)throw new Error('NO_IMAGE_SELECTED');
      if(kind!=='image')await this._imageMaskCanvas(image,maskKey(kind));
      this._imageEditingTarget=kind;this._maskSelectionMode=null;this._maskBrush??={mode:'hide',size:60,hardness:.5};
      this._emit('tool');
    },
    setMaskBrush(values={}) {
      this._maskBrush={mode:values.mode==='restore'?'restore':values.mode==='hide'?'hide':this._maskBrush?.mode||'hide',
        size:clamp(values.size??this._maskBrush?.size??60,1,500),hardness:clamp(values.hardness??this._maskBrush?.hardness??.5,0,1)};
    },
    setMaskSelectionMode(kind='layerMask') {
      if(!isImage(this.selected)||this.selected.locked)throw new Error('NO_IMAGE_SELECTED');
      this._maskSelectionMode=kind;this._maskSelectionTarget=this.selected;this.setTool('brush');
    },
    _startMaskStroke(image,point) {
      const key=maskKey(this._imageEditingTarget),cached=image._dakeMaskCanvases?.[key];
      if(!image[key]||!cached){this.onStatus('maskSelectFirst');return;}
      const {width,height}=sizeOf(image),buffer=canvas(width,height);buffer.getContext('2d').drawImage(cached.canvas,0,0);
      this._maskStroke={image,key,buffer,context:buffer.getContext('2d'),before:clone(image[key]),last:null,paintVersion:0,renderVersion:0,rendering:false};
      this.canvas.setActiveObject(image);this._paintMaskTo(point);
    },
    _paintMaskTo(scenePoint) {
      const stroke=this._maskStroke;if(!stroke)return;
      const {image,context}=stroke,p=imagePoint(image,scenePoint),matrix=image.calcTransformMatrix();
      const scale=Math.sqrt(Math.abs(matrix[0]*matrix[3]-matrix[1]*matrix[2]))||1;
      const brush=this._maskBrush||{mode:'hide',size:60,hardness:.5},radius=brush.size/scale/2,last=stroke.last||p;
      const steps=Math.max(1,Math.ceil(Math.hypot(p.x-last.x,p.y-last.y)/Math.max(1,radius*.25)));
      const rgb=brush.mode==='restore'?'255,255,255':'0,0,0';
      for(let step=1;step<=steps;step++){
        const x=last.x+(p.x-last.x)*step/steps,y=last.y+(p.y-last.y)*step/steps;
        if(brush.hardness>=.999)context.fillStyle='rgb('+rgb+')';
        else{const gradient=context.createRadialGradient(x,y,radius*brush.hardness,x,y,radius);gradient.addColorStop(0,'rgba('+rgb+',1)');gradient.addColorStop(1,'rgba('+rgb+',0)');context.fillStyle=gradient;}
        context.beginPath();context.arc(x,y,radius,0,Math.PI*2);context.fill();
      }
      stroke.last=p;stroke.paintVersion++;this._previewMaskStroke(stroke);
    },
    _previewMaskStroke(stroke) {
      if(stroke.rendering)return;stroke.rendering=true;stroke.renderVersion=stroke.paintVersion;
      stroke.image[stroke.key]={...stroke.image[stroke.key],src:stroke.buffer.toDataURL('image/png')};
      stroke.task=this._renderImageEditing(stroke.image).catch(error=>{stroke.error=error;}).finally(()=>{
        stroke.rendering=false;if(this._maskStroke===stroke&&stroke.renderVersion!==stroke.paintVersion)this._previewMaskStroke(stroke);
      });
    },
    async _finishMaskStroke(apply=true) {
      const stroke=this._maskStroke;if(!stroke)return;
      this._maskStroke=null;await stroke.task;
      if(!apply||stroke.error){stroke.image[stroke.key]=stroke.before;await this._renderImageEditing(stroke.image);if(stroke.error)this.onStatus('rasterFailed');return;}
      stroke.image[stroke.key]={...stroke.before,src:stroke.buffer.toDataURL('image/png')};
      try{
        await this._renderImageEditing(stroke.image);
        stroke.image._dakeMaskCanvases[stroke.key]={src:stroke.image[stroke.key].src,canvas:stroke.buffer};
        this.commit('mask-brush');
      }catch(error){stroke.image[stroke.key]=stroke.before;await this._renderImageEditing(stroke.image);this.onStatus('rasterFailed');}
    },
    _pointerDown(event) {
      if(!this._busy&&this._maskSelectionMode&&(event.e.button===undefined||event.e.button===0)){
        const image=this._maskSelectionTarget,point=this._point(event);this.canvas.setActiveObject(image);
        this._maskSelection={image,kind:this._maskSelectionMode,start:point,end:point,rect:new Rect({left:point.x,top:point.y,width:1,height:1,originX:'left',originY:'top',fill:'#5cd7bc22',stroke:'#5cd7bc',strokeWidth:1/this.zoom,strokeDashArray:[5/this.zoom,5/this.zoom],selectable:false,evented:false,excludeFromExport:true})};
        this.canvas.add(this._maskSelection.rect);return;
      }
      if(!this._busy&&this._imageEditingTarget&&this._imageEditingTarget!=='image'&&['brush','eraser'].includes(this.activeTool)){
        const image=this._gestureTarget||this.selected;this._gestureTarget=null;
        if(isImage(image)&&!image.locked&&(event.e.button===undefined||event.e.button===0)){this._startMaskStroke(image,this._point(event));return;}
      }
      return original.pointerDown.call(this,event);
    },
    _pointerMove(event) {
      if(this._maskSelection){
        const s=this._maskSelection,p=this._point(event);s.end=p;
        s.rect.set({left:Math.min(s.start.x,p.x),top:Math.min(s.start.y,p.y),width:Math.abs(p.x-s.start.x),height:Math.abs(p.y-s.start.y)});
        this.canvas.requestRenderAll();return;
      }
      if(this._maskStroke){this._paintMaskTo(this._point(event));return;}
      return original.pointerMove.call(this,event);
    },
    _pointerUp() {
      if(this._maskSelection){
        const s=this._maskSelection;this._maskSelection=null;this._maskSelectionMode=null;this.canvas.remove(s.rect);this.canvas.setActiveObject(s.image);
        const rect={left:Math.min(s.start.x,s.end.x),top:Math.min(s.start.y,s.end.y),width:Math.abs(s.end.x-s.start.x),height:Math.abs(s.end.y-s.start.y)};
        if(rect.width<2||rect.height<2)return;
        this._strokeTask=this.createImageMask(s.kind,{sceneRect:rect}).then(()=>this.setImageEditingTarget(s.kind)).catch(()=>this.onStatus('rasterFailed')).finally(()=>{this._strokeTask=null;});return this._strokeTask;
      }
      if(this._maskStroke){
        // Existing busy guard protects document changes until the gesture is committed.
        const pending=this._finishMaskStroke(true);this._busy=true;this._applyTool();this._emit('busy');
        this._strokeTask=pending.finally(()=>{this._busy=false;this._strokeTask=null;this._applyTool();this._emit('ready');});return this._strokeTask;
      }
      // A destructive image stroke still edits the original copy; restore its recipe and masks before committing.
      if(this._stroke&&IMAGE_EDIT_PROPERTIES.some(k=>this._stroke.image[k])){
        const stroke=this._stroke;this._stroke=null;stroke.image.filters=[];
        stroke.image.setElement(stroke.buffer,{width:stroke.width,height:stroke.height});stroke.image.filters=stroke.filters;
        const pending=this._withBusy(async()=>{try{await this._renderImageEditing(stroke.image);this.commit('raster');}
          catch(error){Object.assign(stroke.image,stroke.previous);this.canvas.requestRenderAll();this.onStatus('rasterFailed');}});
        this._strokeTask=pending;pending.finally(()=>{if(this._strokeTask===pending)this._strokeTask=null;});return pending;
      }
      return original.pointerUp.call(this);
    }
  });
  // Object.assign evaluates accessors; define dynamic transaction status explicitly.
  Object.defineProperty(P,'imagingPreviewing',{get(){return !!this._imagingPreview||!!this._maskStroke||!!this._maskSelection;}});
}

