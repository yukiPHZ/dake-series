'use strict';
const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs/promises'),path=require('node:path'),crypto=require('node:crypto'),https=require('node:https');
const {EventEmitter}=require('node:events'),{PassThrough}=require('node:stream');
const {temp}=require('../scripts/local-test-env.cjs')();
const {createFontService}=require('../desktop/fonts.cjs'),{extractFont}=require('../desktop/font-binary.cjs');
const storage=require('../desktop/storage.cjs');
const sha=bytes=>crypto.createHash('sha256').update(bytes).digest('hex');
const git=bytes=>crypto.createHash('sha1').update(Buffer.from('blob '+bytes.length+'\0')).update(bytes).digest('hex');
function fixture(){const bytes=Buffer.alloc(162);bytes.writeUInt32BE(0x10000);bytes.writeUInt16BE(2,4);bytes.write('head',12);bytes.writeUInt32BE(44,20);bytes.writeUInt32BE(54,24);bytes.write('OS/2',28);bytes.writeUInt32BE(98,36);bytes.writeUInt32BE(64,40);bytes.writeUInt16BE(500,102);return bytes}
async function setup(t,{invalidLicense=false,invalidHash=false,cancel=false,large=false}={}){
 const output=await fs.mkdtemp(path.join(temp,'font-unit-')),font=fixture(),license=Buffer.from(invalidLicense?'not a license':'SIL OPEN FONT LICENSE Version 1.1\nSynthetic font test fixture.');
 const catalogPath=path.join(output,'catalog.json');await fs.writeFile(catalogPath,JSON.stringify({sourceCommit:'a'.repeat(40),families:[{family:'Fixture',directory:'ofl/fixture',licenseType:'OFL',licensePath:'ofl/fixture/OFL.txt',licenseBlobSha:git(license),files:[{path:'ofl/fixture/Fixture.ttf',style:'normal',weight:400,size:large?60*1024*1024:font.length,blobSha:invalidHash?'0'.repeat(40):git(font)}]}]}));
 const urls=[],originalGet=https.get;let service;
 https.get=(url,options,callback)=>{urls.push(url);const request=new EventEmitter();request.setTimeout=()=>{};request.destroy=error=>request.emit('error',error);options.signal.addEventListener('abort',()=>{const error=new Error('aborted');error.name='AbortError';request.emit('error',error)},{once:true});
  process.nextTick(()=>{const response=new PassThrough();response.statusCode=200;response.headers={};callback(response);const bytes=url.endsWith('OFL.txt')?license:font;if(cancel)response.write(bytes.subarray(0,1));else response.end(bytes)});return request};
 t.after(()=>{https.get=originalGet});
 service=createFontService({userData:output,catalogPath,progress:()=>{if(cancel)service.cancel('cancel-me')}});
 return {service,output,font,license,urls};
}
test('font acquisition pins URLs, verifies hashes and license, and reuses complete cache without network',async t=>{
 const {service,urls,font}=await setup(t);const first=await service.download({family:'Fixture',weight:400,text:'must not leave this PC'});
 assert.equal(first.weight,500);assert.equal(first.sha256,sha(font));assert.equal(first.cached,false);
 assert.equal(urls.length,2);assert.ok(urls.every(url=>url.startsWith('https://raw.githubusercontent.com/google/fonts/'+'a'.repeat(40)+'/ofl/fixture/')&&!url.includes('?')&&!url.includes('must')));
 https.get=()=>{throw new Error('Network forbidden for cached read')};const second=await service.download({family:'Fixture'});assert.equal(second.cached,true);assert.equal(second.data,first.data);
});
test('unknown families, missing license, and changed pinned font bytes fail without publishing cache',async t=>{
 const {service,urls}=await setup(t,{invalidLicense:true});await assert.rejects(service.download({family:'absent'}));assert.equal(urls.length,0);await assert.rejects(service.download({family:'Fixture'}));assert.deepEqual(await service.cachedMetadata(),[]);
});
test('pinned font content mismatch fails before cache publication',async t=>{const {service}=await setup(t,{invalidHash:true});await assert.rejects(service.download({family:'Fixture'}));assert.deepEqual(await service.cachedMetadata(),[])});
test('canceling both partial font and license downloads publishes no descriptor',async t=>{const {service}=await setup(t,{cancel:true});await assert.rejects(service.download({family:'Fixture',requestId:'cancel-me'}),/中止/);assert.deepEqual(await service.cachedMetadata(),[])});
test('untrusted font tables and TTC face selection fail within bounds',()=>{assert.throws(()=>extractFont(Buffer.alloc(30)));const bytes=fixture();bytes.writeUInt32BE(999999,20);assert.throws(()=>extractFont(bytes));const ttc=Buffer.alloc(16);ttc.write('ttcf');ttc.writeUInt32BE(129,8);assert.throws(()=>extractFont(ttc,'missing'))});
test('v2 project stores metadata and licensed Google bytes, rejects local bytes, hash changes, and invalid layouts',()=>{
 const bytes=fixture(),font={family:'Fixture',source:'google',weight:500,data:'data:font/ttf;base64,'+bytes.toString('base64'),license:'SIL OPEN FONT LICENSE Version 1.1',sha256:sha(bytes)};
 const data={format:'dake-stadio',version:2,width:600,height:400,meta:{dpi:300,unit:'mm',bleed:35,safe:12,guides:{horizontal:[100],vertical:[200]},snap:{gridSize:10}},fonts:[font],canvas:{objects:[]}};
 assert.equal(storage.validateProject(data),data);
 assert.throws(()=>storage.validateProject({...data,fonts:[{...font,source:'local'}]}),error=>error.code==='LOCAL_FONT_EMBED');
 assert.throws(()=>storage.validateProject({...data,fonts:[{...font,sha256:'0'.repeat(64)}]}),error=>error.code==='FONT_HASH_MISMATCH');
 assert.throws(()=>storage.validateProject({...data,fonts:[{...font,license:''}]}),error=>error.code==='INVALID_FONT_LICENSE');
 assert.throws(()=>storage.validateProject({...data,meta:{...data.meta,guides:{horizontal:[500]}}}));
 assert.throws(()=>storage.validateProject({...data,meta:{...data.meta,bleed:200}}));
 assert.equal(storage.validateProject({...data,fonts:[{family:'Local',source:'local',postscriptName:'Local-Regular'}]}).version,2);
});

test('oversize catalog fonts are rejected before network or document mutation',async t=>{const {service,urls}=await setup(t,{large:true});await assert.rejects(service.download({family:'Fixture'}));assert.equal(urls.length,0);assert.deepEqual(await service.cachedMetadata(),[])});
