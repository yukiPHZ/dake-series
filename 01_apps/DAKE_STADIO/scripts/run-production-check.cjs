'use strict';
require('./local-test-env.cjs')();
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { spawn } = require('node:child_process');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const electron = require('electron');
const root = path.resolve(__dirname, '..');
const output = path.resolve(process.env.STADIO_PRODUCTION_OUTPUT || path.join(root, 'test-output', `production-${Date.now()}`));
const phaseOnly = process.argv[2];
const reportPath = path.join(root, 'evidence', 'production-results.json');
if (!output.startsWith(path.join(root, 'test-output') + path.sep)) throw new Error('Expected isolated output under test-output');
const report = { output, started: new Date().toISOString(), checks: [], phaseResults: {} };
async function execute(command,args,environment,timeoutMs=600000){
  return new Promise((resolve,reject)=>{
    const child=spawn(command,args,{cwd:root,env:environment,stdio:['ignore','pipe','pipe'],windowsHide:true});
    let stdout='',stderr='';child.stdout.on('data',chunk=>{stdout+=chunk;process.stdout.write(chunk);});child.stderr.on('data',chunk=>{stderr+=chunk;process.stderr.write(chunk);});
    const timeout=setTimeout(()=>{child.kill();reject(new Error(`QA child exceeded ${timeoutMs} ms`));},timeoutMs);
    child.once('error',error=>{clearTimeout(timeout);reject(error);});child.once('exit',code=>{clearTimeout(timeout);code===0?resolve({stdout,stderr}):reject(new Error(`QA child exited with ${code}`));});
  });
}
async function phase(name){
  const environment={...process.env,STADIO_PRODUCTION_OUTPUT:output};delete environment.ELECTRON_RUN_AS_NODE;
  await execute(electron,['scripts/check-production.cjs',name],environment);
  const result=JSON.parse(await fs.readFile(path.join(output,name+'.json'),'utf8'));report.phaseResults[name]=result;assert.ok(result.passed);return result;
}
async function validatePdfs(prepare){
  const dependencyRoot=path.join(os.homedir(),'.cache','codex-runtimes','codex-primary-runtime','dependencies');
  const python=process.env.STADIO_PYTHON||path.join(dependencyRoot,'python','python.exe');
  const poppler=process.env.STADIO_PDFTOPPM||path.join(dependencyRoot,'native','poppler','Library','bin','pdftoppm.exe');
  await fs.access(python);await fs.access(poppler);
  const source=String.raw`
import json,sys,pathlib,subprocess
from pypdf import PdfReader
from PIL import Image,ImageChops,ImageStat
root=pathlib.Path(sys.argv[1]); poppler=sys.argv[2]
prepare=json.loads((root/'prepare.json').read_text(encoding='utf-8'))
results={}
for kind in ('banner','flyer','logo','card'):
    source=root/(kind+'.pdf'); reader=PdfReader(str(source))
    assert len(reader.pages)==1, (kind,'Expected one PDF page')
    page=reader.pages[0]; box=page.mediabox
    actual=[float(box.width),float(box.height)]
    work=prepare['works'][kind]; expected=[work['width']/work['meta']['dpi']*72,work['height']/work['meta']['dpi']*72]
    assert all(abs(a-b)<0.2 for a,b in zip(actual,expected)), (kind,'PDF physical size mismatch',actual,expected)
    prefix=root/(kind+'-pdf-render')
    subprocess.run([poppler,'-singlefile','-r','96','-png',str(source),str(prefix)],check=True,capture_output=True)
    rendered=Image.open(str(prefix)+'.png').convert('RGB')
    original=Image.open(root/(kind+'.png')).convert('RGBA')
    flattened=Image.new('RGBA',original.size,'white');flattened.alpha_composite(original)
    reference=flattened.convert('RGB').resize(rendered.size,Image.Resampling.LANCZOS)
    difference=ImageChops.difference(reference,rendered)
    mean=sum(ImageStat.Stat(difference).mean)/3
    assert mean<6.0,(kind,'Independent PDF rendering differs from exported artwork',mean)
    objects=page.get('/Resources',{}).get('/XObject',{})
    images=[]
    for key,value in objects.items():
        item=value.get_object()
        if item.get('/Subtype')=='/Image':images.append({'width':item.get('/Width'),'height':item.get('/Height'),'colorSpace':str(item.get('/ColorSpace'))})
    results[kind]={'pages':1,'mediaBoxPoints':actual,'expectedPoints':expected,'physicalMm':[v/72*25.4 for v in actual],'explicitTrimBox':'/TrimBox' in page,'explicitBleedBox':'/BleedBox' in page,'images':images,'renderedPng':str(prefix)+'.png','renderMeanAbsoluteError':round(mean,5),'renderer':'Poppler pdftoppm,96dpi','parser':'pypdf'}
(root/'pdf-validation.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(results,ensure_ascii=False,indent=2))
`;
  await execute(python,['-c',source,output,poppler],process.env,120000);
  return JSON.parse(await fs.readFile(path.join(output,'pdf-validation.json'),'utf8'));
}
async function publishExamples(){
  const directory=path.join(root,'artifacts','production');await fs.mkdir(directory,{recursive:true});
  const files=[];
  for(const kind of ['banner','flyer','logo','card'])for(const suffix of ['-edited.dake','.png','.jpg','.webp','.svg','.pdf','-edited-screen.png','-pdf-render.png']){
    const from=path.join(output,kind+suffix),to=path.join(directory,kind+suffix);await fs.copyFile(from,to);const bytes=await fs.readFile(to);files.push({file:to,bytes:bytes.length,sha256:crypto.createHash('sha256').update(bytes).digest('hex')});
  }
  return files;
}
async function main(){
  await fs.mkdir(output,{recursive:true});
  if(['prepare','resume','safety'].includes(phaseOnly)){await phase(phaseOnly);report.passed=true;await fs.writeFile(reportPath,JSON.stringify(report,null,2));return;}
  const prepare=await phase('prepare'),resume=await phase('resume');
  assert.notEqual(prepare.processId,resume.processId);report.processes=[prepare.processId,resume.processId];report.checks=[...prepare.checks,...resume.checks,{name:'Preparation and offline reopening used distinct Electron PIDs'}];
  report.pdfValidation=await validatePdfs(prepare);report.checks.push({name:'All four PDFs independently parsed and rendered with correct physical size and matching raster appearance'});
  report.artifacts=await publishExamples();report.passed=true;report.finished=new Date().toISOString();
  report.notes=['Native file-dialog answers are replaced by isolated paths; DOM controls, CDP file drops, Canvas, workers, preload, IPC, storage and exporters are real.','Prepare downloads a Google font via native HTTPS; resume has fresh user data and blocks external native requests.','PDF exports are RGB raster pages; this does not establish PDF/X, CMYK or commercial print acceptance.','Automated pixel and geometry checks do not replace final human visual review of generated artwork.'];
  await fs.writeFile(reportPath,JSON.stringify(report,null,2));console.log(JSON.stringify({passed:true,checks:report.checks.length,output,artifacts:report.artifacts.length}));
}
main().catch(async error=>{report.passed=false;report.failure=error.stack;for(const name of ['prepare','resume'])try{report.phaseResults[name]=JSON.parse(await fs.readFile(path.join(output,name+'.json'),'utf8'));}catch{}await fs.writeFile(reportPath,JSON.stringify(report,null,2));console.error(error);process.exitCode=1;});