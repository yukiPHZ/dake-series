import { FabricImage, Group, ActiveSelection, Path, Textbox, Point, util, cache } from 'fabric';
import UI_TEXT from './ui-text.json';

export const FOUNDATION_PROPERTIES = ['dakeMask','dakeOutline'];
const newId=()=>crypto.randomUUID();
const clone=value=>structuredClone(value);
const finite=(value,fallback)=>Number.isFinite(Number(value))?Number(value):fallback;
const vectorTypes=new Set(['path','rect','ellipse','circle','triangle','polygon','polyline','line']);
const label=key=>UI_TEXT.layerNames?.[key]||UI_TEXT.layerNames?.path;
export function normalizeMeta(input={},width=1200,height=800) {
  const dpi=finite(input.dpi,96), unit=input.unit||'px',bleed=finite(input.bleed,0),safe=finite(input.safe,0);
  if(dpi<36||dpi>2400||!['px','mm'].includes(unit)||bleed<0||safe<0||2*(bleed+safe)>=Math.min(width,height))throw new Error('INVALID_LAYOUT');
  const guides={horizontal:[],vertical:[]};
  for(const axis of ['horizontal','vertical']){
    const list=input.guides?.[axis]||[];
    if(!Array.isArray(list)||list.length>200||list.some(n=>!Number.isFinite(n)||n<0||n>(axis==='horizontal'?height:width)))throw new Error('INVALID_LAYOUT');
    guides[axis]=[...new Set(list)].sort((a,b)=>a-b);
  }
  const snap={enabled:true,objects:true,artboard:true,guides:true,grid:false,gridSize:16,...input.snap};
  for(const key of ['enabled','objects','artboard','guides','grid'])snap[key]=!!snap[key];
  snap.gridSize=finite(snap.gridSize,16);
  if(snap.gridSize<1||snap.gridSize>4096)throw new Error('INVALID_LAYOUT');
  return {dpi,unit,bleed,safe,guides,snap,guideLocked:!!input.guideLocked};
}
export function validateV2Extras(data){
  normalizeMeta(data.meta||{},data.width,data.height);
  const fonts=data.fonts||[];
  if(!Array.isArray(fonts)||fonts.length>32)throw new Error('INVALID_DOCUMENT');
  let bytes=0;
  for(const font of fonts){
    if(!font||typeof font.family!=='string'||font.family.length>200||!['google','local'].includes(font.source))throw new Error('INVALID_DOCUMENT');
    if(font.data!==undefined){
      if(font.source!=='google'||typeof font.data!=='string'||!/^data:(?:font\/(?:ttf|otf|woff2?|sfnt)|application\/(?:font-sfnt|octet-stream));base64,/i.test(font.data)||font.data.length>48*1024*1024)throw new Error('INVALID_DOCUMENT');
      bytes+=font.data.length;
    }
    if(typeof font.license!=='undefined'&&(typeof font.license!=='string'||font.license.length>128000))throw new Error('INVALID_DOCUMENT');
    for(const key of ['__proto__','prototype','constructor'])if(Object.hasOwn(font,key))throw new Error('INVALID_DOCUMENT');
  }
  if(bytes>96*1024*1024)throw new Error('INVALID_DOCUMENT');
}
function packet(object){
  if(object instanceof Group){
    if(object.dakeMask)throw new Error('NO_VECTOR_SELECTED');
    return {children:object.getObjects().map(packet)};
  }
  const type=String(object.type).toLowerCase();
  if(!vectorTypes.has(type))throw new Error('NO_VECTOR_SELECTED');
  const item={type,matrix:object.calcTransformMatrix(),width:object.width,height:object.height,rx:object.rx,ry:object.ry,fillRule:object.fillRule};
  if(object.path){item.path=clone(object.path);item.pathOffset={x:object.pathOffset.x,y:object.pathOffset.y};}
  if(object.points){item.points=clone(object.points);item.pathOffset={x:object.pathOffset.x,y:object.pathOffset.y};}
  if(type==='line')Object.assign(item,object.calcLinePoints());
  return item;
}
const styleOf=object=>({fill:object.fill,stroke:object.stroke,strokeWidth:object.strokeWidth,strokeUniform:true,opacity:object.opacity,strokeLineCap:object.strokeLineCap,strokeLineJoin:object.strokeLineJoin,paintFirst:object.paintFirst});
function invalidateGeometry(object,commands){
  const oldOffset=object.pathOffset;
  const oldCenter=new Point(oldOffset.x,oldOffset.y).subtract(oldOffset).transform(object.calcTransformMatrix());
  object.path=commands;object.setDimensions();
  const nextCenter=new Point(oldOffset.x,oldOffset.y).subtract(object.pathOffset).transform(object.calcTransformMatrix());
  object.left+=oldCenter.x-nextCenter.x;object.top+=oldCenter.y-nextCenter.y;object.set('dirty',true);object.setCoords();
}
const endpoint=command=>({x:command[command.length-2],y:command[command.length-1]});
const lerp=(a,b,t=.5)=>({x:a.x+(b.x-a.x)*t,y:a.y+(b.y-a.y)*t});
function cubic(command,start){
  const end=endpoint(command);
  if(command[0]==='C')return [start,{x:command[1],y:command[2]},{x:command[3],y:command[4]},end];
  if(command[0]==='Q'){const q={x:command[1],y:command[2]};return[start,lerp(start,q,2/3),lerp(end,q,2/3),end];}
  return[start,lerp(start,end,1/3),lerp(start,end,2/3),end];
}

export function installFoundation(Engine){
  Object.assign(Engine.prototype,{
    _initFoundation(){
      this.meta=normalizeMeta({},this.width,this.height);this.fonts=[];
      this.view={guidesVisible:true,gridVisible:false,snapLines:[]};
      this.activeNodeIndex=null;this._vectorWorker=null;this._vectorPending=null;this._vectorSequence=0;
      this.canvas.on('object:moving',({target,e})=>this._snapMoving(target,e));
      this.canvas.on('object:modified',()=>{this.view.snapLines=[];this._emit('view');});
    },
    updateDocumentSettings(partial){
      this._guard();
      const next={...this.meta,...partial,guides:{...this.meta.guides,...partial.guides},snap:{...this.meta.snap,...partial.snap}};
      this.meta=normalizeMeta(next,this.width,this.height);this.commit('layout');return this.meta;
    },
    setView(partial){
      for(const key of ['guidesVisible','gridVisible'])if(partial[key]!==undefined)this.view[key]=!!partial[key];
      this._emit('view');return this.view;
    },
    addGuide(axis,position){
      const key=axis==='x'||axis==='vertical'?'vertical':'horizontal';
      this.updateDocumentSettings({guides:{[key]:[...this.meta.guides[key],Number(position)]}});this.onStatus('guidesChanged');
    },
    removeGuide(axis,index){
      const key=axis==='x'||axis==='vertical'?'vertical':'horizontal';
      this.updateDocumentSettings({guides:{[key]:this.meta.guides[key].filter((_,i)=>i!==index)}});
    },
    setFonts(fonts){this._guard();validateV2Extras({width:this.width,height:this.height,meta:this.meta,fonts});this.fonts=clone(fonts);this.commit('fonts');},
    renameLayer(id,name){
      this._guard();const object=this.layers.find(layer=>layer.id===id);if(!object)return;
      object.name=String(name).trim().slice(0,200)||label(object.type);this.commit('rename-layer');
    },
    selectMany(ids){this._guard();return this._selectMany(ids);},
    _selectMany(ids){this.canvas.discardActiveObject();
      const objects=this.layers.filter(object=>ids.includes(object.id)&&!object.locked&&object.visible);
      if(objects.length)this.canvas.setActiveObject(objects.length===1?objects[0]:new ActiveSelection(objects,{canvas:this.canvas}));
      this.canvas.requestRenderAll();this.onSelection(this.selected);return objects;
    },
    moveLayer(id,index){
      this._guard();const object=this.layers.find(layer=>layer.id===id);if(!object)return;
      this.canvas.moveObjectTo(object,Math.min(this.layers.length-1,Math.max(0,Math.round(index))));this.commit('reorder');
    },
    alignSelected(direction){
      this._guard();const objects=this.canvas.getActiveObjects().filter(object=>!object.locked);
      if(!objects.length)return;
      this.canvas.discardActiveObject();
      const bounds=objects.map(object=>object.getBoundingRect());
      const left=objects.length===1?this.meta.bleed:Math.min(...bounds.map(rect=>rect.left));
      const top=objects.length===1?this.meta.bleed:Math.min(...bounds.map(rect=>rect.top));
      const right=objects.length===1?this.width-this.meta.bleed:Math.max(...bounds.map(rect=>rect.left+rect.width));
      const bottom=objects.length===1?this.height-this.meta.bleed:Math.max(...bounds.map(rect=>rect.top+rect.height));
      objects.forEach((object,index)=>{
        const rect=bounds[index];let dx=0,dy=0;
        if(direction==='left')dx=left-rect.left;
        if(direction==='center')dx=(left+right)/2-rect.left-rect.width/2;
        if(direction==='right')dx=right-rect.left-rect.width;
        if(direction==='top')dy=top-rect.top;
        if(direction==='middle')dy=(top+bottom)/2-rect.top-rect.height/2;
        if(direction==='bottom')dy=bottom-rect.top-rect.height;
        object.set({left:object.left+dx,top:object.top+dy});object.setCoords();
      });
      this.selectMany(objects.map(object=>object.id));this.commit('align');
    },
    distributeSelected(axis){
      this._guard();const objects=this.canvas.getActiveObjects().filter(object=>!object.locked);
      if(objects.length<3)throw new Error('NEED_THREE_LAYERS');
      this.canvas.discardActiveObject();
      const horizontal=axis==='horizontal',position=horizontal?'left':'top',size=horizontal?'width':'height';
      const ordered=objects.map(object=>({object,rect:object.getBoundingRect()})).sort((a,b)=>a.rect[position]-b.rect[position]);
      const first=ordered[0].rect[position],last=ordered.at(-1).rect;
      const gap=(last[position]+last[size]-first-ordered.reduce((sum,item)=>sum+item.rect[size],0))/(ordered.length-1);
      let cursor=first;
      for(const {object,rect} of ordered){object.set(position,object[position]+cursor-rect[position]);object.setCoords();cursor+=rect[size]+gap;}
      this.selectMany(objects.map(object=>object.id));this.commit('distribute');
    },
    getEffectiveDPI(object=this.selected){
      if(object instanceof Group){
        const values=object.getObjects().map(child=>this.getEffectiveDPI(child)).filter(Boolean);
        return values.length?{x:Math.min(...values.map(v=>v.x)),y:Math.min(...values.map(v=>v.y)),min:Math.min(...values.map(v=>v.min))}:null;
      }
      if(!(object instanceof FabricImage))return null;
      const matrix=object.calcTransformMatrix();const x=this.meta.dpi/Math.max(.00001,Math.hypot(matrix[0],matrix[1]));const y=this.meta.dpi/Math.max(.00001,Math.hypot(matrix[2],matrix[3]));
      return {x,y,min:Math.min(x,y)};
    },
    _snapMoving(target,event){
      if(!target||!this.meta.snap.enabled||event?.altKey){this.view.snapLines=[];return;}
      const selected=new Set(this.canvas.getActiveObjects()),snap=this.meta.snap;
      const rect=target.getBoundingRect(),xValues=[rect.left,rect.left+rect.width/2,rect.left+rect.width],yValues=[rect.top,rect.top+rect.height/2,rect.top+rect.height];
      const xs=[],ys=[];
      if(snap.artboard){xs.push(0,this.width/2,this.width,this.meta.bleed,this.width-this.meta.bleed);ys.push(0,this.height/2,this.height,this.meta.bleed,this.height-this.meta.bleed);}
      if(snap.guides){xs.push(...this.meta.guides.vertical);ys.push(...this.meta.guides.horizontal);}
      if(snap.objects)for(const object of this.layers){
        if(object===target||selected.has(object)||!object.visible)continue;
        const box=object.getBoundingRect();xs.push(box.left,box.left+box.width/2,box.left+box.width);ys.push(box.top,box.top+box.height/2,box.top+box.height);
      }
      if(snap.grid){for(const x of xValues)xs.push(Math.round(x/snap.gridSize)*snap.gridSize);for(const y of yValues)ys.push(Math.round(y/snap.gridSize)*snap.gridSize);}
      const nearest=(from,to)=>{
        let result=null;const tolerance=6/this.zoom;
        for(const value of from)for(const candidate of to){const delta=candidate-value;if(Math.abs(delta)<=tolerance&&(!result||Math.abs(delta)<Math.abs(result.delta)))result={delta,position:candidate};}
        return result;
      };
      const x=nearest(xValues,xs),y=nearest(yValues,ys);this.view.snapLines=[];
      if(x){target.left+=x.delta;this.view.snapLines.push({axis:'x',position:x.position});}
      if(y){target.top+=y.delta;this.view.snapLines.push({axis:'y',position:y.position});}
      target.setCoords();this._emit('view');
    },
    _vectorRequest(method,payload,transfer=[]){
      if(this._vectorPending)return Promise.reject(new Error('BUSY'));
      if(!this._vectorWorker){
        this._vectorWorker=new Worker(new URL('./vector-worker.js',document.baseURI));
        this._vectorWorker.onmessage=({data})=>{
          const pending=this._vectorPending;if(!pending||pending.id!==data.id)return;
          clearTimeout(pending.timer);this._vectorPending=null;
          data.error?pending.reject(new Error(data.error)):pending.resolve(data.result);
        };
        this._vectorWorker.onerror=()=>{
          const pending=this._vectorPending;this._vectorPending=null;
          if(pending){clearTimeout(pending.timer);pending.reject(new Error('VECTOR_TOO_COMPLEX'));}
          this._vectorWorker?.terminate();this._vectorWorker=null;
        };
      }
      return new Promise((resolve,reject)=>{
        const id=++this._vectorSequence;
        const timer=setTimeout(()=>{this._vectorWorker?.terminate();this._vectorWorker=null;this._vectorPending=null;reject(new Error('VECTOR_TOO_COMPLEX'));},45000);
        this._vectorPending={id,resolve,reject,timer};
        try{this._vectorWorker.postMessage({id,method,payload},transfer);}catch(error){clearTimeout(timer);this._vectorPending=null;reject(error);}
      });
    },


    _replacePaths(objects,result,kind='path'){
      if(!result.paths?.length||result.paths.length>2000)throw new Error('VECTOR_EMPTY');
      const base=objects[0],style=styleOf(base);
      const paths=result.paths.map((data,index)=>this._decorate(new Path(data,{...style,fillRule:result.fillRule||base.fillRule||'nonzero',objectCaching:false}),kind,index===0?base.name:undefined));
      if(paths.some(path=>!Number.isFinite(path.width)||!Number.isFinite(path.height)||path.path.length>40000))throw new Error('VECTOR_TOO_COMPLEX');
      const insertion=Math.min(...objects.map(object=>this.layers.indexOf(object)));
      this.canvas.discardActiveObject();this.canvas.remove(...objects);
      this.canvas.insertAt(insertion,...paths);this._selectMany(paths.map(path=>path.id));
      if(this.activeTool==='node')this._enableNodeControls(paths[0]);
      return paths;
    },
    async booleanSelected(operation){
      const objects=this.layers.filter(object=>this.canvas.getActiveObjects().includes(object)&&!object.locked);
      if(objects.length<2)throw new Error('NEED_TWO_SHAPES');
      const items=objects.map(packet);
      return this._withBusy(async()=>{
        const result=await this._vectorRequest('boolean',{operation,items});
        const paths=this._replacePaths(objects,result,operation==='compound'?'compound':'path');
        this.commit('boolean');this.onStatus('vectorDone');return paths;
      });
    },
    async splitCompound(){
      const object=this.selected;if(!(object instanceof Path)||object.locked)throw new Error('NO_VECTOR_SELECTED');
      const item=packet(object);
      return this._withBusy(async()=>{
        const result=await this._vectorRequest('boolean',{operation:'split',items:[item]});
        const paths=this._replacePaths([object],result);this.commit('split');return paths;
      });
    },
    async convertSelectedToPath(){
      const objects=this.canvas.getActiveObjects().filter(object=>!object.locked);
      if(!objects.length||objects.some(object=>object instanceof Group))throw new Error('NO_VECTOR_SELECTED');
      const items=objects.map(packet);
      return this._withBusy(async()=>{
        // Each object keeps its own appearance; conversion is not a boolean union.
        const replacements=[];
        for(let i=0;i<items.length;i++){
          const result=await this._vectorRequest('boolean',{operation:'convert',items:[items[i]]});
          if(!result.paths?.length)throw new Error('VECTOR_EMPTY');
          const path=this._decorate(new Path(result.paths[0],{...styleOf(objects[i]),fillRule:objects[i].fillRule,objectCaching:false}),'path',objects[i].name);
          replacements.push({object:objects[i],path,index:this.layers.indexOf(objects[i])});
        }
        this.canvas.discardActiveObject();
        for(const {object,path,index} of replacements){this.canvas.remove(object);this.canvas.insertAt(index,path);}
        this._selectMany(replacements.map(item=>item.path.id));this.commit('convert-path');return replacements.map(item=>item.path);
      });
    },
    async createMask(){
      const objects=this.layers.filter(object=>this.canvas.getActiveObjects().includes(object)&&!object.locked);
      if(objects.length<2)throw new Error('INVALID_MASK');
      const source=objects.at(-1),type=String(source.type).toLowerCase();
      if(!vectorTypes.has(type)||type==='line'||type==='polyline'||!source.fill||source.clipPath)throw new Error('INVALID_MASK');
      if(source instanceof Path&&!source.path.some(command=>command[0]==='Z'))throw new Error('INVALID_MASK');
      return this._withBusy(async()=>{
        const original=await source.clone();
        const clip=await source.clone();
        // A selection supplies world transform through its temporary group.
        util.applyTransformToObject(original,source.calcTransformMatrix());
        const contents=objects.slice(0,-1),index=Math.min(...objects.map(object=>this.layers.indexOf(object)));
        this.canvas.discardActiveObject();this.canvas.remove(...objects);
        const group=this._decorate(new Group(contents,{originX:'left',originY:'top',objectCaching:true}),'mask',label('mask'));
        util.sendObjectToPlane(original,undefined,group.calcTransformMatrix());
        const sourceData=original.toObject();
        util.applyTransformToObject(clip,original.calcTransformMatrix());
        clip.set({fill:'#000000',stroke:null,strokeWidth:0,opacity:1,visible:true,absolutePositioned:false});
        group.dakeMask={enabled:true,source:sourceData};group.clipPath=clip;group._dakeMaskClip=clip;
        this.canvas.insertAt(index,group);this.canvas.setActiveObject(group);
        this.commit('mask');this.onStatus('maskCreated');return group;
      });
    },
    async toggleMask(){
      const group=this.selected;if(!(group instanceof Group)||!group.dakeMask||group.locked)throw new Error('INVALID_MASK');
      return this._withBusy(async()=>{
        let clip=group.clipPath||group._dakeMaskClip;
        if(!clip){[clip]=await util.enlivenObjects([clone(group.dakeMask.source)]);clip.set({fill:'#000000',stroke:null,strokeWidth:0,opacity:1,visible:true,absolutePositioned:false});}
        const enabled=!group.dakeMask.enabled;group.dakeMask={...group.dakeMask,enabled};
        group._dakeMaskClip=clip;group.set('clipPath',enabled?clip:undefined);group.set('dirty',true);
        this.commit('toggle-mask');this.onStatus('maskToggled');return enabled;
      });
    },
    async releaseMask(){
      const group=this.selected;if(!(group instanceof Group)||!group.dakeMask||group.locked)throw new Error('INVALID_MASK');
      return this._withBusy(async()=>{
        const [source]=await util.enlivenObjects([clone(group.dakeMask.source)]);
        const index=this.layers.indexOf(group);this.canvas.discardActiveObject();
        util.sendObjectToPlane(source,group.calcTransformMatrix(),undefined);
        const children=group.removeAll();this.canvas.remove(group);
        this.canvas.insertAt(index,...children,source);this._selectMany([...children,source].map(object=>object.id));
        this.commit('release-mask');this.onStatus('maskReleased');return [...children,source];
      });
    },
    editNode(operation,index=this.activeNodeIndex){
      this._guard();const object=this.selected;
      if(!(object instanceof Path)||object.locked)throw new Error('NO_VECTOR_SELECTED');
      const path=clone(object.path);
      if(index!=null&&(!Number.isInteger(index)||index<0||index>=path.length))throw new Error('NO_NODE_SELECTED');
      // Compound contours retain their independent topology. Work on the contour containing the selected node.
      let start=Number.isInteger(index)?index:0;
      while(start>0&&path[start][0]!=='M')start--;
      let end=start+1;while(end<path.length&&path[end][0]!=='M')end++;
      const isClosed=path[end-1]?.[0]==='Z';
      // Keep one explicit closing segment when editing the first node. The
      // duplicated endpoint is a seam, not an additional logical anchor.
      let seamIndex=-1;
      if(isClosed&&end-start>2){const first=endpoint(path[start]),lastPoint=endpoint(path[end-2]);
        if(first.x===lastPoint.x&&first.y===lastPoint.y)seamIndex=end-2;
      }
      if(index===seamIndex&&seamIndex>=0)index=start;
      if(operation==='close'){
        if(!isClosed){if(end-start<3)throw new Error('NO_NODE_SELECTED');path.splice(end,0,['Z']);}
      }else if(operation==='open'){
        if(isClosed){path.splice(end-1,1);if(seamIndex>=0)path.splice(seamIndex,1);}
      }else{
        if(!Number.isInteger(index)||index<start||index>=end||path[index][0]==='Z')throw new Error('NO_NODE_SELECTED');
        const last=seamIndex>=0?seamIndex-1:isClosed?end-2:end-1;
        if(operation==='insert'){
          const next=index<last?index+1:isClosed?start:-1;
          if(next<0)throw new Error('NO_NODE_SELECTED');
          const a=endpoint(path[index]),command=next===start?(seamIndex>=0?path[seamIndex]:['L',...Object.values(endpoint(path[start]))]):path[next];
          const [p0,p1,p2,p3]=cubic(command,a),a1=lerp(p0,p1),a2=lerp(p1,p2),a3=lerp(p2,p3),b1=lerp(a1,a2),b2=lerp(a2,a3),mid=lerp(b1,b2);
          const first=['C',a1.x,a1.y,b1.x,b1.y,mid.x,mid.y],second=['C',b2.x,b2.y,a3.x,a3.y,p3.x,p3.y];
          if(next===start)path.splice(last+1,seamIndex>=0?1:0,first,second);else path.splice(next,1,first,second);
          this.activeNodeIndex=index+1;
        }else if(operation==='remove'){
          if(last-start+1<=(isClosed?3:2))throw new Error('NO_NODE_SELECTED');
          if(index===start){const point=endpoint(path[start+1]);if(seamIndex>=0){path[seamIndex][path[seamIndex].length-2]=point.x;path[seamIndex][path[seamIndex].length-1]=point.y;}path.splice(start,2,['M',point.x,point.y]);}
          else path.splice(index,1);
          this.activeNodeIndex=Math.max(start,index-1);
        }else if(operation==='smooth'||operation==='corner'){
          const previous=index>start?index-1:isClosed?last:-1,next=index<last?index+1:isClosed?start:-1;
          if(previous<0&&next<0)throw new Error('NO_NODE_SELECTED');
          const current=endpoint(path[index]),before=previous>=0?endpoint(path[previous]):current,after=next>=0?endpoint(path[next]):current;
          const length=Math.hypot(after.x-before.x,after.y-before.y)||1,ux=(after.x-before.x)/length,uy=(after.y-before.y)/length;
          const inLength=operation==='corner'?0:Math.hypot(current.x-before.x,current.y-before.y)/3;
          const outLength=operation==='corner'?0:Math.hypot(after.x-current.x,after.y-current.y)/3;
          if(index>start){const c=cubic(path[index],before);path[index]=['C',c[1].x,c[1].y,current.x-ux*inLength,current.y-uy*inLength,current.x,current.y];}
          else if(isClosed){const c=seamIndex>=0?cubic(path[seamIndex],before):cubic(['L',current.x,current.y],before);path.splice(last+1,seamIndex>=0?1:0,['C',c[1].x,c[1].y,current.x-ux*inLength,current.y-uy*inLength,current.x,current.y]);}
          if(next>start){const c=cubic(path[next],current);path[next]=['C',current.x+ux*outLength,current.y+uy*outLength,c[2].x,c[2].y,c[3].x,c[3].y];}
          else if(isClosed){const to=endpoint(path[start]),c=seamIndex>=0?cubic(path[seamIndex],current):cubic(['L',to.x,to.y],current);path.splice(last+1,seamIndex>=0?1:0,['C',current.x+ux*outLength,current.y+uy*outLength,c[2].x,c[2].y,to.x,to.y]);}
        }else throw new Error('NO_NODE_SELECTED');
      }
      invalidateGeometry(object,path);this._enableNodeControls(object);this.commit('node');this.onStatus('nodeChanged');return object;
    },
    async outlineText(fontBuffer,descriptor={}){
      const object=this.selected;
      if(!(object instanceof Textbox)&&!['text','i-text','itext'].includes(String(object?.type).toLowerCase()))throw new Error('FONT_OUTLINE_FAILED');
      if(object.locked||object.path||object.direction==='rtl'||Object.keys(object.styles||{}).length||object.textBackgroundColor||object.backgroundColor||object.clipPath||object.shadow||object.stroke&&object.strokeWidth>0&&(object.underline||object.overline||object.linethrough)||object.textDecorationColor&&object.textDecorationColor!==object.fill)throw new Error('FONT_UNSUPPORTED');
      if(!(fontBuffer instanceof ArrayBuffer))throw new Error('FONT_OUTLINE_FAILED');
      const weight=object.fontWeight==='bold'?700:object.fontWeight==='normal'?400:finite(object.fontWeight,400);
      return this._withBusy(async()=>{
        if(object.isEditing)object.exitEditing();
        await document.fonts.load(object.fontStyle+' '+weight+' '+object.fontSize+'px "'+object.fontFamily+'"');
        cache.clearFontCache(object.fontFamily);object.initDimensions();object.setCoords();
        const runs=[],decorations=[];let y=-object.height/2;
        for(let i=0;i<object._textLines.length;i++){
          const chars=object._textLines[i],lineHeight=object.getHeightOfLine(i),baseline=y+lineHeight/object.lineHeight*(1-object._fontSizeFraction);
          const x=-object.width/2+object._getLineLeftOffset(i);
          if(!object.charSpacing&&!String(object.textAlign).includes('justify'))runs.push({text:chars.join(''),chars,positions:chars.map((char,j)=>object.__charBounds[i][j].left),x,y:baseline,ligatures:true});
          else chars.forEach((char,j)=>runs.push({text:char,x:x+object.__charBounds[i][j].left,y:baseline,ligatures:false}));
          for(const type of ['underline','overline','linethrough'])if(object[type]&&chars.length){
            const thickness=object.fontSize*(object.textDecorationThickness||66.667)/1000,align=type==='linethrough'?.5:type==='overline'?1:0;
            const rectY=baseline+object.offsets[type]*object.fontSize-align*thickness;
            const width=object.__charBounds[i].slice(0,chars.length).reduce((sum,box)=>sum+box.kernedWidth,0)-object._getWidthOfCharSpacing();
            decorations.push({x,y:rectY,width,height:thickness});
          }
          y+=lineHeight;
        }
        const buffer=fontBuffer.slice(0);
        const result=await this._vectorRequest('outline',{font:buffer,weight,style:object.fontStyle,fontSize:object.fontSize,matrix:object.calcTransformMatrix(),runs,decorations},[buffer]);
        const [path]=this._replacePaths([object],result,'outline');
        path.name=object.name;path.dakeOutline={family:descriptor.family||object.fontFamily,weight,style:object.fontStyle,originalText:object.text,sourceId:descriptor.id||'',fsType:descriptor.fsType||0};
        this.commit('outline');this.onStatus('outlineDone');return path;
      });
    },
  });
  const dispose=Engine.prototype.dispose;
  Engine.prototype.dispose=async function(){
    if(this._vectorPending){clearTimeout(this._vectorPending.timer);this._vectorPending.reject(new Error('BUSY'));this._vectorPending=null;}
    this._vectorWorker?.terminate();this._vectorWorker=null;
    return dispose.call(this);
  };
}

