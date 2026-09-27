'use strict';
require('./local-test-env.cjs')();
const {app}=require('electron'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),distribution=path.resolve(process.argv[2]||path.join(root,'dist','DAKE_STADIO-win32-x64'));
const output=path.join(root,'test-output','packaged-native-'+Date.now());fs.mkdirSync(output,{recursive:true});app.setPath('userData',path.join(output,'profile'));
const report={distribution,output,checks:[]},check=(name,value)=>{assert.ok(value,name);report.checks.push(name)};
app.whenReady().then(async()=>{
 const archive=path.join(distribution,'resources','app.asar');
 const native=require(path.join(archive,'desktop','print-artwork.cjs'));
 const png=fs.readFileSync(path.join(distribution,'examples','card.png')),project=JSON.parse(fs.readFileSync(path.join(distribution,'examples','card.dake'),'utf8'));
 const started=Date.now();let prior=performance.now(),maxGap=0;
 const tick=setInterval(()=>{const now=performance.now();maxGap=Math.max(maxGap,now-prior);prior=now},10);
 let bytes;try{bytes=await native.toPDF({data:'data:image/png;base64,'+png.toString('base64'),width:project.width,height:project.height,dpi:project.meta.dpi})}finally{clearInterval(tick)}
 check('native PDF worker executes from inside packaged app.asar',bytes.toString('ascii',0,5)==='%PDF-');
 const media=/\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/.exec(bytes.toString('latin1'));
 check('packaged PDF retains complete physical card size',media&&Math.abs(Number(media[1])-project.width/project.meta.dpi*72)<.00001&&Math.abs(Number(media[2])-project.height/project.meta.dpi*72)<.00001);
 check('packaged PDF work does not block main event loop for 500ms',maxGap<500);
 fs.writeFileSync(path.join(output,'packaged-card.pdf'),bytes);report.pdfBytes=bytes.length;report.elapsedMs=Date.now()-started;report.maxMainGapMs=Math.round(maxGap);report.passed=true;
 fs.writeFileSync(path.join(root,'evidence','packaged-native-results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));app.exit(0);
}).catch(error=>{report.passed=false;report.failure=error.stack;fs.writeFileSync(path.join(root,'evidence','packaged-native-results.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1)});
