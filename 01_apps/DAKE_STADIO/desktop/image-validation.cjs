'use strict';
function validateImageEditing(raw) {
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

module.exports={validateImageEditing};
