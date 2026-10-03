// Read only the Unicode cmap. No text or artwork is sent to a font service.
const cache=new Map();
export function missingFontGlyphs(descriptor,text){
  const chars=[...new Set([...String(text||'')].filter(ch=>{const cp=ch.codePointAt(0);return !/\s/.test(ch)&&cp!==0x200d&&!(cp>=0xfe00&&cp<=0xfe0f);}))];
  if(Array.isArray(descriptor.coverage))return chars.filter(ch=>!descriptor.coverage.some(([a,b])=>ch.codePointAt(0)>=a&&ch.codePointAt(0)<=b));
  const key=descriptor.sha256||descriptor.id;let tables=cache.get(key);
  if(!tables){
    if(!descriptor.data)return null;
    try{
      const encoded=descriptor.data.slice(descriptor.data.indexOf(',')+1),binary=atob(encoded),bytes=Uint8Array.from(binary,c=>c.charCodeAt(0)),view=new DataView(bytes.buffer),count=view.getUint16(4);let cmap;
      if(count>200)throw Error('invalid');
      for(let i=0;i<count;i++){const at=12+i*16;if(view.getUint32(at)===0x636d6170){const offset=view.getUint32(at+8),size=view.getUint32(at+12);cmap=new DataView(bytes.slice(offset,offset+size).buffer);break;}}
      if(!cmap)return null;tables=[];const n=cmap.getUint16(2);if(n>256)throw Error('invalid');
      for(let i=0;i<n;i++){const at=4+i*8,platform=cmap.getUint16(at),encoding=cmap.getUint16(at+2),sub=cmap.getUint32(at+4);if(platform!==0&&!(platform===3&&[1,10].includes(encoding)))continue;const format=cmap.getUint16(sub);
        if(format===12){const groups=cmap.getUint32(sub+12);if(groups>200000)throw Error('invalid');const ranges=[];for(let j=0;j<groups;j++){const g=sub+16+j*12;ranges.push([cmap.getUint32(g),cmap.getUint32(g+4),cmap.getUint32(g+8)]);}tables.push({format,ranges});}
        if(format===4){const size=cmap.getUint16(sub+2),data=new DataView(cmap.buffer.slice(sub,sub+size)),segments=data.getUint16(6)/2;if(!Number.isInteger(segments)||segments>32767)throw Error('invalid');tables.push({format,data,segments});}
      }if(!tables.length)return null;if(key)cache.set(key,tables);
    }catch{return null;}
  }
  function has(cp){return tables.some(t=>{if(t.format===12)return t.ranges.some(([a,b,g])=>cp>=a&&cp<=b&&g+(cp-a)!==0);if(cp>65535)return false;const ends=14,starts=ends+t.segments*2+2,deltas=starts+t.segments*2,offsets=deltas+t.segments*2;for(let j=0;j<t.segments;j++){const a=t.data.getUint16(starts+j*2),b=t.data.getUint16(ends+j*2);if(cp<a||cp>b)continue;const delta=t.data.getInt16(deltas+j*2),offset=t.data.getUint16(offsets+j*2);if(!offset)return((cp+delta)&65535)!==0;const pos=offsets+j*2+offset+2*(cp-a);if(pos+2>t.data.byteLength)return false;const glyph=t.data.getUint16(pos);return glyph!==0&&((glyph+delta)&65535)!==0;}return false;});}
  return chars.filter(ch=>!has(ch.codePointAt(0)));
}
