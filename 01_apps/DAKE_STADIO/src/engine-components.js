import {validateFontIdentity} from '../desktop/font-identity.cjs';
import { Group, Rect, StaticCanvas, LayoutManager, FixedLayout, util } from 'fabric';
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


const copy = value => structuredClone(value);
const uid = () => crypto.randomUUID();
export const COMPONENT_PROPERTIES = ['dakeComponent'];
export { validateComponents };
const position = object => Object.fromEntries(['id','name','left','top','scaleX','scaleY','angle','skewX','skewY','flipX','flipY','opacity','visible','locked','originX','originY'].filter(k => object[k] !== undefined).map(k => [k, object[k]]));
export function componentFonts(document) {
  const result = new Map();
  const visit = doc => {
    for (const font of doc.fonts || []) result.set([font.source,font.family,font.weight,font.style,font.sha256].join('|'),copy(font));
    const walk = objects => {for (const object of objects || []) {if(object.dakeComponent)visit(object.dakeComponent.document);else walk(object.objects);}};
    walk(doc.canvas.objects);
  }; visit(document); return [...result.values()];
}
function mergeFonts(engine, document) {
  const result = new Map();
  for (const font of [...(engine.fonts||[]),...componentFonts(document)]) result.set([font.source,font.family,font.weight,font.style,font.sha256].join('|'),copy(font));
  engine.fonts = [...result.values()];
}
export function installComponents(Engine) {
  const baseNew = Engine.prototype.newDocument;
  const baseRestore = Engine.prototype._restore;
  const baseSerialize = Engine.prototype.serialize;
  const baseUngroup = Engine.prototype.ungroup;
  Engine.prototype.newDocument = function(options) {
    this.documentId = uid(); return baseNew.call(this,options);
  };
  Engine.prototype.serialize = function() {
    const data = baseSerialize.call(this);
    data.version = 3; data.documentId = this.documentId ||= uid();
    return data;
  };
  Engine.prototype._buildComponent = async function(document, metadata, existing={}) {
    validateComponents(document);
    const el = globalThis.document.createElement('canvas');
    // Reuse the same complete image/mask/filter restoration path as ordinary documents.
    const staging = new Engine(el);
    staging.onChange=()=>{}; staging.onSelection=()=>{}; staging.onStatus=()=>{};
    try {
      await staging._restore(copy(document));
      const children=staging.layers; staging.canvas.remove(...children);
      const background = new Rect({left:0,top:0,originX:'left',originY:'top',width:document.width,height:document.height,
        fill:document.background==='transparent'?'rgba(0,0,0,0)':document.background,strokeWidth:0,selectable:false,evented:false});
      const group = new Group([background,...children],{
        left:0,top:0,originX:'left',originY:'top',width:document.width,height:document.height,
        layoutManager:new LayoutManager(new FixedLayout()),objectCaching:true,subTargetCheck:false,interactive:false
      });
      // Explicit document-space origin; out-of-artboard objects do not change placement.
      group.set({width:document.width,height:document.height});
      group.clipPath = new Rect({left:-document.width/2,top:-document.height/2,width:document.width,height:document.height,originX:'left',originY:'top',fill:'#000',strokeWidth:0});
      group.dakeComponent=copy(metadata);group.dakeComponent.document=copy(document);
      group.set({...existing,objectCaching:true});this._decorate(group,'group',existing.name||document.name);
      return group;
    } finally {await staging.dispose();}
  };
  Engine.prototype._restore = async function(data) {
    validateComponents(data);validateFontIdentity(data);
    const prepared=copy(data);
    const expand = async objects => {
      const result=[];
      for(const raw of objects||[]) {
        if(raw.dakeComponent) {
          const group=await this._buildComponent(raw.dakeComponent.document,raw.dakeComponent,position(raw));
          result.push(group.toObject(['dakeComponent','id','name','locked','objectCaching']));
        } else result.push({...raw,...(raw.objects?{objects:await expand(raw.objects)}:{})});
      } return result;
    };
    prepared.canvas.objects=await expand(prepared.canvas.objects);
    prepared.fonts=componentFonts(prepared);
    await baseRestore.call(this,prepared);this.documentId=data.documentId||uid();
  };
  Engine.prototype.insertComponent=async function(document, {mode='embedded',link,originPath,name}={}) {
    this._guard();
    const candidate=copy(document);candidate.documentId ||= uid();
    const metadata={id:uid(),mode,document:candidate,...(originPath?{originPath}:{}),...(link?{link:copy(link)}:{})};
    validateComponents({...this.serialize(),canvas:{objects:[...this.serialize().canvas.objects,{type:'Group',objects:[],dakeComponent:metadata}]}});
    validateFontIdentity({...this.serialize(),canvas:{objects:[...this.serialize().canvas.objects,{type:'Group',dakeComponent:metadata}]}});
    return this._withBusy(async()=>{
      const size=Math.min(1,this.width*.65/candidate.width,this.height*.65/candidate.height);
      const group=await this._buildComponent(candidate,metadata,{left:this.width*.12,top:this.height*.12,scaleX:size,scaleY:size,name:name||candidate.name});
      mergeFonts(this,candidate);this.canvas.add(group);this.canvas.setActiveObject(group);this.commit('component-place');return group;
    });
  };
  Engine.prototype.replaceComponent=async function(layerId, document, {commit=true, link, mode}={}) {
    this._guard();const old=this.layers.find(o=>o.id===layerId);
    if(!old?.dakeComponent)throw new Error('COMPONENT_MISSING');
    const metadata={...copy(old.dakeComponent),document:copy(document)};
    if(link)metadata.link=copy(link);if(mode){metadata.mode=mode;if(mode==='embedded')delete metadata.link;}
    const data=this.serialize(),raw=data.canvas.objects.find(o=>o.id===layerId);
    raw.dakeComponent=metadata;validateComponents(data);validateFontIdentity(data);
    return this._withBusy(async()=>{
      const group=await this._buildComponent(document,metadata,position(old));
      const index=this.layers.indexOf(old);this.canvas.discardActiveObject();this.canvas.remove(old);this.canvas.insertAt(index,group);
      mergeFonts(this,document);this.canvas.setActiveObject(group);
      if(commit)this.commit('component-update');else this.canvas.requestRenderAll();
      return group;
    });
  };
  Engine.prototype.ungroup=function(){
    if(this.selected?.dakeComponent)throw new Error('COMPONENT_EDIT_TAB');
    return baseUngroup.call(this);
  };
}
