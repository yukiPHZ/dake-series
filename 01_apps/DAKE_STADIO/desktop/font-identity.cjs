'use strict';
// Shared by renderer and native validation. No filesystem or browser side effects.
function validateFontIdentity(document){
 const faces=new Map();let documents=0;
 function visit(doc,depth=0){
  if(depth>8||++documents>64)throw new Error('COMPONENT_LIMIT');
  for(const f of doc.fonts||[]){const key=String(f.family).trim().toLocaleLowerCase()+'|'+(f.style||'normal'),range=String(f.weightRange||f.weight||400).split(/\s+/).map(Number),min=range[0],max=range.at(-1);if(!Number.isFinite(min)||!Number.isFinite(max)||min<1||max>1000||min>max)throw new Error('INVALID_DOCUMENT');
   const known=faces.get(key)||[];if(f.sha256&&known.some(other=>other.sha256&&other.sha256!==f.sha256&&min<=other.max&&max>=other.min))throw new Error('FONT_IDENTITY_CONFLICT');known.push({min,max,sha256:f.sha256});faces.set(key,known);
  }
  const walk=objects=>{for(const object of objects||[]){if(object.dakeComponent)visit(object.dakeComponent.document,depth+1);else walk(object.objects);}};walk(doc.canvas?.objects);
 }
 visit(document);return document;
}
module.exports={validateFontIdentity};
