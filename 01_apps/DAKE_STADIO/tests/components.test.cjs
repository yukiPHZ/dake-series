'use strict';
const test=require('node:test'),assert=require('node:assert/strict');
const fs=require('node:fs/promises'),os=require('node:os'),path=require('node:path');
const {validateComponents,createComponentService}=require('../desktop/components.cjs');
const leaf=id=>({format:'dake-stadio',version:3,documentId:id,name:id,width:100,height:80,background:'transparent',fonts:[],canvas:{objects:[]}});
const component=(id,doc,extra={})=>({type:'Group',objects:[],dakeComponent:{id,mode:'embedded',document:doc,...extra}});
test('component source keeps editable text and vectors; validates legacy child and rejects same-document cycles',()=>{
 const child=leaf('logo');child.version=2;child.canvas.objects=[{type:'Textbox',text:'DAKE',fontFamily:'Yu Gothic'},{type:'Path',path:[['M',0,0],['L',10,10]]}];
 const parent=leaf('card');parent.canvas.objects=[component('instance',child)];
 const before=JSON.stringify(parent);assert.equal(validateComponents(parent),parent);assert.equal(JSON.stringify(parent),before);
 parent.canvas.objects[0].dakeComponent.document=leaf('card');assert.throws(()=>validateComponents(parent),/COMPONENT_CYCLE/);
});
test('recursive definition cycle, wrong component types and oversized nesting are refused',()=>{
 const a=leaf('a'),b=leaf('b'),c=leaf('a');a.canvas.objects=[component('i',b)];b.canvas.objects=[component('j',c)];
 assert.throws(()=>validateComponents(a),/COMPONENT_CYCLE/);
 const wrong=leaf('x');wrong.canvas.objects=[{...component('i',leaf('y')),type:'Image'}];assert.throws(()=>validateComponents(wrong),/COMPONENT_INVALID/);
 let root=leaf('d0'),pointer=root;for(let i=1;i<=10;i++){const n=leaf('d'+i);pointer.canvas.objects=[component('i'+i,n)];pointer=n;}
 assert.throws(()=>validateComponents(root),/COMPONENT_INVALID/);
});
test('linked snapshot validates explicit path/hash; arbitrary file reads require a fresh capability',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'stadio-component-'));
 try{
  const file=path.join(dir,'logo.dake');await fs.writeFile(file,JSON.stringify(leaf('logo')));
  const remembered=[];const options={readProject:async p=>JSON.parse(await fs.readFile(p,'utf8')),rememberSource:async p=>remembered.push(p)};
  const service=createComponentService(options);
  await assert.rejects(async()=>service.refresh(file),/COMPONENT_RELINK/);
  const imported=await service.importFile(file,'linked');assert.equal(imported.data.documentId,'logo');assert.equal(imported.link.hash.length,64);assert.equal(remembered.length,1);
  const original=await fs.readFile(file,'utf8');imported.data.name='Edited child';
  const refreshed=await service.refresh(imported.link.token);assert.equal(refreshed.data.name,'logo');assert.equal(await fs.readFile(file,'utf8'),original);
  const changed=leaf('logo');changed.name='Externally edited';await fs.writeFile(file,JSON.stringify(changed));
  const updated=await service.refresh(imported.link.token);assert.notEqual(updated.link.hash,imported.link.hash);assert.equal(updated.data.name,'Externally edited');
  const afterRestart=createComponentService(options);assert.throws(()=>afterRestart.refresh(imported.link.token),/COMPONENT_RELINK/);
  const p=leaf('parent');p.canvas.objects=[component('linked',updated.data,{mode:'linked',link:updated.link})];assert.doesNotThrow(()=>validateComponents(p));
  p.canvas.objects[0].dakeComponent.link.hash='bad';assert.throws(()=>validateComponents(p),/COMPONENT_INVALID/);
 }finally{assert.equal(path.dirname(path.resolve(dir)),path.resolve(os.tmpdir()));assert.ok(path.basename(dir).startsWith('stadio-component-'));await fs.rm(dir,{recursive:true,force:true});}
});
