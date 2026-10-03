/* Original four filters retain Fabric.js 7.4 CPU rounding for existing .dake files.
 * Dake recipes and masks use the same worker for display, restore and export.
 * No artwork or pixels leave the local process. */
function legacyFilter(pixels,filter) {
  if(filter.type==='Brightness'){
    const amount=Math.round(filter.brightness*255);for(let i=0;i<pixels.length;i+=4){pixels[i]+=amount;pixels[i+1]+=amount;pixels[i+2]+=amount;}
  }else if(filter.type==='Contrast'){
    const amount=Math.floor(filter.contrast*255),factor=(259*(amount+255))/(255*(259-amount));
    for(let i=0;i<pixels.length;i+=4){pixels[i]=factor*(pixels[i]-128)+128;pixels[i+1]=factor*(pixels[i+1]-128)+128;pixels[i+2]=factor*(pixels[i+2]-128)+128;}
  }else if(filter.type==='Saturation'){
    const adjust=-filter.saturation;for(let i=0;i<pixels.length;i+=4){const r=pixels[i],g=pixels[i+1],b=pixels[i+2],max=Math.max(r,g,b);pixels[i]+=max!==r?(max-r)*adjust:0;pixels[i+1]+=max!==g?(max-g)*adjust:0;pixels[i+2]+=max!==b?(max-b)*adjust:0;}
  }else if(filter.type==='Grayscale'){
    for(let i=0;i<pixels.length;i+=4){const r=pixels[i],g=pixels[i+1],b=pixels[i+2];const value=filter.mode==='lightness'?(Math.min(r,g,b)+Math.max(r,g,b))/2:filter.mode==='luminosity'?r*.21+g*.72+b*.07:(r+g+b)/3;pixels[i]=pixels[i+1]=pixels[i+2]=value;}
  }else throw new Error('INVALID_DOCUMENT');
}
function curveTable(points) {
  const table=new Uint8ClampedArray(256);let segment=0;
  for(let x=0;x<256;x++){
    while(segment<points.length-2&&x>points[segment+1][0])segment++;
    const a=points[segment],b=points[segment+1];table[x]=a[1]+(b[1]-a[1])*(x-a[0])/(b[0]-a[0]);
  }
  return table;
}
function applyRecipe(pixels,recipe) {
  if(!recipe)return;
  if(recipe.exposure){
    const gain=2**recipe.exposure;
    for(let i=0;i<pixels.length;i+=4)for(let channel=0;channel<3;channel++){
      const srgb=pixels[i+channel]/255,linear=srgb<=.04045?srgb/12.92:((srgb+.055)/1.055)**2.4;
      const v=linear*gain;pixels[i+channel]=(v<=.0031308?v*12.92:1.055*(v**(1/2.4))-.055)*255;
    }
  }
  if(recipe.brightness)legacyFilter(pixels,{type:'Brightness',brightness:recipe.brightness});
  if(recipe.contrast)legacyFilter(pixels,{type:'Contrast',contrast:recipe.contrast});
  if(recipe.saturation)legacyFilter(pixels,{type:'Saturation',saturation:recipe.saturation});
  if(recipe.vibrance){
    for(let i=0;i<pixels.length;i+=4){
      const r=pixels[i],g=pixels[i+1],b=pixels[i+2],maximum=Math.max(r,g,b),minimum=Math.min(r,g,b);
      const saturation=maximum?(maximum-minimum)/maximum:0,amount=recipe.vibrance*(1-saturation),luma=.2126*r+.7152*g+.0722*b;
      pixels[i]=luma+(r-luma)*(1+amount);pixels[i+1]=luma+(g-luma)*(1+amount);pixels[i+2]=luma+(b-luma)*(1+amount);
    }
  }
  if(recipe.curve?.some(p=>p[0]!==p[1])){const table=curveTable(recipe.curve);for(let i=0;i<pixels.length;i+=4){pixels[i]=table[pixels[i]];pixels[i+1]=table[pixels[i+1]];pixels[i+2]=table[pixels[i+2]];}}
  if(recipe.monochrome)legacyFilter(pixels,{type:'Grayscale',mode:recipe.monochromeMode||'luminosity'});
}
async function maskPixels(mask,width,height) {
  if(!mask?.enabled)return null;
  if(typeof mask.src!=='string'||!/^data:image\/png;base64,/.test(mask.src))throw new Error('INVALID_DOCUMENT');
  const binary=atob(mask.src.slice(mask.src.indexOf(',')+1)),bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);const bitmap=await createImageBitmap(new Blob([bytes],{type:'image/png'}));
  try{
    if(bitmap.width!==width||bitmap.height!==height)throw new Error('INVALID_DOCUMENT');
    const c=new OffscreenCanvas(width,height),ctx=c.getContext('2d',{willReadFrequently:true});
    if(mask.feather>0){
      // Extend border pixels before blur so an all-white mask stays white at the image edges.
      const padding=Math.ceil(mask.feather*3),padded=new OffscreenCanvas(width+padding*2,height+padding*2),p=padded.getContext('2d');
      p.drawImage(bitmap,padding,padding);p.drawImage(bitmap,0,0,1,height,0,padding,padding,height);p.drawImage(bitmap,width-1,0,1,height,padding+width,padding,padding,height);
      p.drawImage(bitmap,0,0,width,1,padding,0,width,padding);p.drawImage(bitmap,0,height-1,width,1,padding,padding+height,width,padding);
      for(const [sx,sy,dx,dy] of [[0,0,0,0],[width-1,0,padding+width,0],[0,height-1,0,padding+height],[width-1,height-1,padding+width,padding+height]])p.drawImage(bitmap,sx,sy,1,1,dx,dy,padding,padding);
      ctx.filter='blur('+mask.feather+'px)';ctx.drawImage(padded,-padding,-padding);ctx.filter='none';padded.width=padded.height=1;
    }else ctx.drawImage(bitmap,0,0);
    return ctx.getImageData(0,0,width,height).data;
  }finally{bitmap.close();}
}
self.onmessage=async({data:{id,bitmap,filters}})=>{
  try{
    const width=bitmap.width,height=bitmap.height;
    if(width<1||height<1||width>8192||height>8192||width*height>32000000)throw new Error('DOCUMENT_TOO_LARGE');
    const canvas=new OffscreenCanvas(width,height),context=canvas.getContext('2d',{willReadFrequently:true});
    context.drawImage(bitmap,0,0);bitmap.close();const imageData=context.getImageData(0,0,width,height),pixels=imageData.data;
    for(const filter of filters){
      if(filter.type!=='DakeImageEditing'){legacyFilter(pixels,filter);continue;}
      const before=filter.recipe&&filter.adjustmentMask?.enabled?new Uint8ClampedArray(pixels):null;
      applyRecipe(pixels,filter.recipe);
      if(before){
        const mask=await maskPixels(filter.adjustmentMask,width,height);
        for(let i=0;i<pixels.length;i+=4){const amount=mask[i]/255*mask[i+3]/255;for(let c=0;c<3;c++)pixels[i+c]=before[i+c]+(pixels[i+c]-before[i+c])*amount;}
      }
      const mask=await maskPixels(filter.layerMask,width,height);
      if(mask)for(let i=0;i<pixels.length;i+=4)pixels[i+3]*=mask[i]/255*mask[i+3]/255;
    }
    context.putImageData(imageData,0,0);const result=canvas.transferToImageBitmap();self.postMessage({id,bitmap:result},[result]);
  }catch(error){bitmap?.close();self.postMessage({id,error:error.message||'INVALID_IMAGE'});}
};

